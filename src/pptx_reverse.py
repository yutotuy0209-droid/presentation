#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Reverse an existing PPTX into inspectable/editable text artifacts.

Outputs:
- slide_ir.json: geometry/text/table/media references for supported objects
- slides.md: human-readable slide text + geometry summary
- source_bundle/: optional loss-aware package preserving master/theme/media/unknown OOXML

This does NOT pretend that arbitrary PPTX visual semantics can always be reconstructed
as Grid DSL. Existing PowerPoint files usually do not contain concepts such as
"input role" or "this is a 1x3 grid".
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from pptx_to_ir import parse_pptx


EMU_PER_INCH = 914400


def _inch(v: int | float) -> float:
    return float(v or 0) / EMU_PER_INCH


def _title_candidate(slide: dict[str, Any]) -> str:
    candidates = []
    for el in slide.get("elements") or []:
        text = str(el.get("text") or "").strip()
        if not text:
            continue
        runs = el.get("runs") or []
        max_size = max((float(r.get("fontSizePt") or 0) for r in runs), default=0)
        candidates.append((
            max_size,
            -int(el.get("y") or 0),
            text.replace("\n", " ").strip(),
        ))
    if not candidates:
        return ""
    candidates.sort(reverse=True)
    return candidates[0][2]


def ir_to_markdown(deck: dict[str, Any]) -> str:
    lines = [
        "# PPTX reverse text",
        "",
        f"Source: {deck.get('meta', {}).get('filename', '')}",
        "",
        "This file is an inspection/editing aid. The original PPTX package remains authoritative.",
        "",
    ]
    for slide in deck.get("slides") or []:
        number = slide.get("number") or ""
        title = _title_candidate(slide)
        lines += [f"## Slide {number}" + (f"｜{title}" if title else ""), ""]

        for el in slide.get("elements") or []:
            etype = str(el.get("type") or "unknown")
            name = str(el.get("name") or el.get("id") or etype)
            text = str(el.get("text") or "").strip()
            lines += [
                f"### {name}",
                "",
                f"- type: {etype}",
                (
                    "- box: "
                    f"x={_inch(el.get('x', 0)):.3f}in, "
                    f"y={_inch(el.get('y', 0)):.3f}in, "
                    f"w={_inch(el.get('width', 0)):.3f}in, "
                    f"h={_inch(el.get('height', 0)):.3f}in"
                ),
            ]
            if el.get("mediaTarget"):
                lines.append(f"- media: {el['mediaTarget']}")
            if el.get("table"):
                lines.append("- table: yes")
            if text:
                lines += ["", text]
            lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def reverse_pptx(source: Path, out_dir: Path, *, make_bundle: bool = True) -> dict[str, Any]:
    source = source.resolve()
    out_dir = out_dir.resolve()
    if out_dir.exists() and any(out_dir.iterdir()):
        raise FileExistsError(f"Choose an empty output directory: {out_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    deck = parse_pptx(source)
    ir_path = out_dir / "slide_ir.json"
    md_path = out_dir / "slides.md"
    ir_path.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
    md_path.write_text(ir_to_markdown(deck), encoding="utf-8")

    bundle_status: dict[str, Any] = {"requested": make_bundle, "created": False}
    if make_bundle:
        try:
            from pptx_bundle import unpack
            bundle_dir = out_dir / "source_bundle"
            unpack(source, bundle_dir)
            bundle_status = {
                "requested": True,
                "created": True,
                "path": str(bundle_dir),
                "text": str(bundle_dir / "slides.md"),
                "manifest": str(bundle_dir / "deck.json"),
            }
        except ImportError as exc:
            bundle_status = {
                "requested": True,
                "created": False,
                "reason": f"optional dependency missing: {exc}",
            }

    summary = {
        "source": str(source),
        "slideCount": len(deck.get("slides") or []),
        "ir": str(ir_path),
        "markdown": str(md_path),
        "bundle": bundle_status,
        "limitations": [
            "Arbitrary PPTX does not contain Grid DSL roles/semantic intent.",
            "slide_ir.json is editable but not guaranteed lossless for every PowerPoint object.",
            "source_bundle preserves unknown/master/theme/media parts and is the safer round-trip path.",
        ],
    }
    (out_dir / "reverse-report.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return summary


def main() -> int:
    ap = argparse.ArgumentParser(description="PPTX -> text/IR/source bundle")
    ap.add_argument("pptx", type=Path)
    ap.add_argument("--out-dir", type=Path, required=True)
    ap.add_argument("--no-bundle", action="store_true")
    args = ap.parse_args()

    result = reverse_pptx(args.pptx, args.out_dir, make_bundle=not args.no_bundle)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
