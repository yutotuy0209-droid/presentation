#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Finalize a generated PPTX with Microsoft PowerPoint through Python COM.

The planning path is platform-independent and covered by CI.
The apply path is Windows-only and imports pywin32 lazily.

Supported today:
- entrance animations: appear / fade / wipe / fly
- triggers: click / with_previous / after_previous
- slide transitions: none / fade / wipe_left / push_left
- hyperlinks on generated shapes

The base PPTX is still generated deterministically by Grid DSL / Pixel Layout.
This module only handles PowerPoint-native features that are awkward or unavailable
in the normal writer.
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any


class FinalizeError(RuntimeError):
    pass


SUPPORTED_ANIMATIONS = {
    "appear": ("msoAnimEffectAppear",),
    "fade": ("msoAnimEffectFade",),
    "wipe": ("msoAnimEffectWipe",),
    "fly": ("msoAnimEffectFly",),
}

SUPPORTED_TRIGGERS = {
    "click": ("msoAnimTriggerOnPageClick",),
    "with_previous": ("msoAnimTriggerWithPrevious",),
    "after_previous": ("msoAnimTriggerAfterPrevious",),
}

SUPPORTED_TRANSITIONS = {
    "none": ("ppEffectNone",),
    "fade": ("ppEffectFadeSmoothly", "ppEffectFade"),
    "wipe_left": ("ppEffectWipeLeft",),
    "push_left": ("ppEffectPushLeft", "ppEffectPushLeftDown"),
}


def _shape_index(
    shape_map: dict[str, Any],
    slide_id: str,
    element_id: str,
) -> int | None:
    raw = (shape_map.get(slide_id) or {}).get(element_id)
    if raw is None:
        return None
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return None
    return value if value > 0 else None


def build_finalize_plan(
    deck: dict[str, Any],
    generation_report: dict[str, Any],
) -> dict[str, Any]:
    """Build a deterministic PowerPoint-native post-processing plan."""
    shape_map = generation_report.get("shapeMap") or {}
    warnings: list[dict[str, Any]] = []
    slides_out: list[dict[str, Any]] = []

    raw_template = (deck.get("meta") or {}).get("template")
    template_plan = None
    if raw_template:
        if isinstance(raw_template, str):
            template_plan = {"path": raw_template, "applyOnWindows": True}
        elif isinstance(raw_template, dict):
            template_plan = {
                "path": str(raw_template.get("path") or ""),
                "applyOnWindows": bool(raw_template.get("applyOnWindows", True)),
            }

    for slide_index, slide in enumerate(deck.get("slides") or [], start=1):
        slide_id = str(slide.get("id") or f"slide-{slide_index}")
        transition = slide.get("transition")
        normalized_transition = None
        if transition:
            if isinstance(transition, str):
                transition = {"type": transition}
            ttype = str(transition.get("type") or "none").lower()
            if ttype not in SUPPORTED_TRANSITIONS:
                raise FinalizeError(f"unsupported transition: {ttype}")
            normalized_transition = {
                "type": ttype,
                "duration": float(transition.get("duration", 0.5)),
            }

        animations: list[dict[str, Any]] = []
        hyperlinks: list[dict[str, Any]] = []

        for element in slide.get("elements") or []:
            element_id = str(element.get("id") or "")
            if not element_id:
                continue
            shape_index = _shape_index(shape_map, slide_id, element_id)

            animation = element.get("animation")
            if animation:
                entrance = str(animation.get("entrance") or "appear").lower()
                trigger = str(animation.get("trigger") or "click").lower()
                if entrance not in SUPPORTED_ANIMATIONS:
                    raise FinalizeError(f"unsupported animation: {entrance}")
                if trigger not in SUPPORTED_TRIGGERS:
                    raise FinalizeError(f"unsupported animation trigger: {trigger}")
                if shape_index is None:
                    warnings.append({
                        "slide": slide_id,
                        "element": element_id,
                        "reason": "missing shape mapping for animation",
                    })
                else:
                    animations.append({
                        "element": element_id,
                        "shapeIndex": shape_index,
                        "entrance": entrance,
                        "order": int(animation.get("order", len(animations) + 1)),
                        "trigger": trigger,
                        "duration": float(animation.get("duration", 0.45)),
                    })

            interaction = element.get("interaction") or {}
            hyperlink = interaction.get("hyperlink")
            if hyperlink:
                if shape_index is None:
                    warnings.append({
                        "slide": slide_id,
                        "element": element_id,
                        "reason": "missing shape mapping for hyperlink",
                    })
                else:
                    hyperlinks.append({
                        "element": element_id,
                        "shapeIndex": shape_index,
                        "address": str(hyperlink),
                    })

        animations.sort(key=lambda item: (item["order"], item["shapeIndex"]))
        slides_out.append({
            "id": slide_id,
            "slideIndex": slide_index,
            "transition": normalized_transition,
            "animations": animations,
            "hyperlinks": hyperlinks,
        })

    return {
        "version": 1,
        "backend": "python-win32com-powerpoint",
        "template": template_plan,
        "slides": slides_out,
        "warnings": warnings,
    }


def _resolve_constant(constants: Any, candidates: tuple[str, ...]) -> Any:
    for name in candidates:
        try:
            return getattr(constants, name)
        except AttributeError:
            continue
    raise FinalizeError(f"PowerPoint constant not available: {candidates}")


def _transition_speed(constants: Any, duration: float) -> Any:
    if duration <= 0.35:
        return _resolve_constant(constants, ("ppTransitionSpeedFast",))
    if duration >= 0.85:
        return _resolve_constant(constants, ("ppTransitionSpeedSlow",))
    return _resolve_constant(constants, ("ppTransitionSpeedMedium",))


def apply_finalize_plan(
    input_pptx: Path,
    output_pptx: Path,
    plan: dict[str, Any],
    *,
    visible: bool = False,
) -> Path:
    """Apply the plan through installed desktop PowerPoint on Windows."""
    if os.name != "nt":
        raise FinalizeError("PowerPoint COM finalization requires Windows.")

    try:
        import win32com.client  # type: ignore
    except ImportError as exc:
        raise FinalizeError(
            "pywin32 is required. Install with: python -m pip install pywin32"
        ) from exc

    input_pptx = input_pptx.resolve()
    output_pptx = output_pptx.resolve()
    if not input_pptx.exists():
        raise FileNotFoundError(input_pptx)
    output_pptx.parent.mkdir(parents=True, exist_ok=True)

    app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
    constants = win32com.client.constants
    app.Visible = bool(visible)
    presentation = None
    try:
        presentation = app.Presentations.Open(
            str(input_pptx),
            ReadOnly=False,
            Untitled=False,
            WithWindow=bool(visible),
        )

        template = plan.get("template")
        if template and template.get("applyOnWindows") and template.get("path"):
            template_path = Path(str(template["path"])).expanduser().resolve()
            if not template_path.exists():
                raise FileNotFoundError(template_path)
            presentation.ApplyTemplate(str(template_path))

        for slide_plan in plan.get("slides") or []:
            slide = presentation.Slides(int(slide_plan["slideIndex"]))

            transition = slide_plan.get("transition")
            if transition:
                transition_type = str(transition["type"])
                slide.SlideShowTransition.EntryEffect = _resolve_constant(
                    constants, SUPPORTED_TRANSITIONS[transition_type]
                )
                slide.SlideShowTransition.Speed = _transition_speed(
                    constants, float(transition.get("duration", 0.5))
                )

            sequence = slide.TimeLine.MainSequence
            for anim in slide_plan.get("animations") or []:
                shape = slide.Shapes(int(anim["shapeIndex"]))
                effect_const = _resolve_constant(
                    constants, SUPPORTED_ANIMATIONS[str(anim["entrance"])]
                )
                trigger_const = _resolve_constant(
                    constants, SUPPORTED_TRIGGERS[str(anim["trigger"])]
                )
                effect = sequence.AddEffect(shape, effect_const, 0, trigger_const)
                try:
                    effect.Timing.Duration = float(anim.get("duration", 0.45))
                except Exception:
                    pass

            for link in slide_plan.get("hyperlinks") or []:
                shape = slide.Shapes(int(link["shapeIndex"]))
                mouse_click = _resolve_constant(constants, ("ppMouseClick",))
                action_hyperlink = _resolve_constant(constants, ("ppActionHyperlink",))
                action = shape.ActionSettings(mouse_click)
                action.Action = action_hyperlink
                action.Hyperlink.Address = str(link["address"])

        if output_pptx == input_pptx:
            presentation.Save()
        else:
            presentation.SaveCopyAs(str(output_pptx))
        return output_pptx
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass


def main() -> int:
    ap = argparse.ArgumentParser(description="Finalize PPTX using Python + desktop PowerPoint")
    sub = ap.add_subparsers(dest="command", required=True)

    p_plan = sub.add_parser("plan", help="IR + generation report -> finalization plan")
    p_plan.add_argument("ir", type=Path)
    p_plan.add_argument("report", type=Path)
    p_plan.add_argument("--out", type=Path, required=True)

    p_apply = sub.add_parser("apply", help="Apply a prepared plan on Windows PowerPoint")
    p_apply.add_argument("pptx", type=Path)
    p_apply.add_argument("plan", type=Path)
    p_apply.add_argument("--out", type=Path, required=True)
    p_apply.add_argument("--visible", action="store_true")

    args = ap.parse_args()
    if args.command == "plan":
        deck = json.loads(args.ir.read_text(encoding="utf-8"))
        report = json.loads(args.report.read_text(encoding="utf-8"))
        plan = build_finalize_plan(deck, report)
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[powerpoint-finalize] plan: {args.out}")
        if plan["warnings"]:
            print(json.dumps(plan["warnings"], ensure_ascii=False, indent=2))
        return 0

    plan = json.loads(args.plan.read_text(encoding="utf-8"))
    result = apply_finalize_plan(args.pptx, args.out, plan, visible=args.visible)
    print(f"[powerpoint-finalize] PPTX: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
