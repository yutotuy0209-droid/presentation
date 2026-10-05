#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""2D Grid DSL -> existing pixel layout engine.

Goal:
- Humans/AI choose semantic blocks and a small 2D grid.
- Python deterministically turns that into editable PowerPoint geometry.
- Coordinates/fonts/colors are owned by presets, not re-invented per prompt.

Example:
    grid:
      [["input", "process", "output"]]

Repeated IDs create a rectangular span:
    [["photo", "result"],
     ["photo", "note"]]

Connectors are optional and separate from placement.
"""
from __future__ import annotations

import argparse
import json
import math
from copy import deepcopy
from pathlib import Path
from typing import Any

from pixel_layout_engine import (
    DEFAULT_THEME,
    DeckCompiler,
    LayoutError,
    Rect,
    build_pptx,
    deep_merge,
    layout_zones,
    resolve_canvas,
    run_ir_lint,
)

ROOT = Path(__file__).resolve().parents[1]
PRESET_DIR = ROOT / "presets"


def load_design_preset(name: str) -> dict[str, Any]:
    path = PRESET_DIR / f"{name}.json"
    if not path.exists():
        raise LayoutError(f"design preset not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def _normalize_design(spec: dict[str, Any]) -> dict[str, Any]:
    raw = spec.get("design", "glanceable-scientific")
    if isinstance(raw, str):
        preset = load_design_preset(raw)
    elif isinstance(raw, dict):
        preset_name = str(raw.get("preset") or "glanceable-scientific")
        preset = load_design_preset(preset_name)
        preset = deep_merge(preset, {k: v for k, v in raw.items() if k != "preset"})
    else:
        preset = load_design_preset("glanceable-scientific")
    return preset


def _partition(total: int, count: int, gap: int, weights: list[float] | None = None) -> tuple[list[int], list[int]]:
    if count <= 0:
        raise LayoutError("grid must have at least one row/column")
    usable = total - gap * (count - 1)
    if usable <= 0:
        raise LayoutError("grid gap is too large for the content region")

    weights = list(weights or [1.0] * count)
    if len(weights) != count or any(float(w) <= 0 for w in weights):
        raise LayoutError("grid weights must be positive and match row/column count")

    weight_sum = sum(float(w) for w in weights)
    raw = [usable * float(w) / weight_sum for w in weights]
    sizes = [int(math.floor(v)) for v in raw]
    remainder = usable - sum(sizes)
    for i in range(remainder):
        sizes[i % count] += 1

    starts = [0] * count
    cursor = 0
    for i, size in enumerate(sizes):
        starts[i] = cursor
        cursor += size + gap
    return starts, sizes


def _validate_grid(grid: list[list[Any]]) -> tuple[int, int]:
    if not grid or not isinstance(grid, list):
        raise LayoutError("slide.grid must be a non-empty 2D array")
    cols = len(grid[0])
    if cols <= 0:
        raise LayoutError("slide.grid must have at least one column")
    for row in grid:
        if not isinstance(row, list) or len(row) != cols:
            raise LayoutError("slide.grid must be rectangular")
    return len(grid), cols


def _block_rects(
    grid: list[list[Any]],
    content: Rect,
    gap: int,
    row_weights: list[float] | None = None,
    col_weights: list[float] | None = None,
) -> dict[str, Rect]:
    rows, cols = _validate_grid(grid)
    x_starts, widths = _partition(content.w, cols, gap, col_weights)
    y_starts, heights = _partition(content.h, rows, gap, row_weights)

    coordinates: dict[str, list[tuple[int, int]]] = {}
    for r, row in enumerate(grid):
        for c, raw_id in enumerate(row):
            if raw_id in (None, "", ".", "-"):
                continue
            bid = str(raw_id)
            coordinates.setdefault(bid, []).append((r, c))

    rects: dict[str, Rect] = {}
    for bid, cells in coordinates.items():
        rs = [r for r, _ in cells]
        cs = [c for _, c in cells]
        r0, r1 = min(rs), max(rs)
        c0, c1 = min(cs), max(cs)

        # Repeated IDs must form a full rectangle. This makes span semantics deterministic.
        for rr in range(r0, r1 + 1):
            for cc in range(c0, c1 + 1):
                if str(grid[rr][cc]) != bid:
                    raise LayoutError(
                        f"grid block '{bid}' is non-rectangular; repeated IDs must fill a rectangle"
                    )

        x = content.x + x_starts[c0]
        y = content.y + y_starts[r0]
        w = sum(widths[c0 : c1 + 1]) + gap * (c1 - c0)
        h = sum(heights[r0 : r1 + 1]) + gap * (r1 - r0)
        rects[bid] = Rect(x, y, w, h)
    return rects


def _role_style(preset: dict[str, Any], role: str) -> dict[str, Any]:
    roles = preset.get("roles") or {}
    return deepcopy(roles.get(role) or roles.get("neutral") or {})


def _text_block(
    *,
    text: str,
    x: int,
    y: int,
    w: int,
    h: int,
    font_pt: float,
    color: str,
    bold: bool = False,
    align: str = "left",
    z: int = 20,
    animation: dict[str, Any] | None = None,
    hyperlink: str | None = None,
) -> dict[str, Any]:
    out = {
        "type": "text",
        "text": text,
        "position": {"x": x, "y": y, "w": w, "h": h, "mode": "fixed"},
        "fontSizePt": font_pt,
        "minFontSizePt": max(10, min(font_pt, 14)),
        "style": {
            "fontSizePt": font_pt,
            "color": color,
            "bold": bold,
            "align": align,
            "valign": "middle",
        },
        "z": z,
    }
    if animation:
        out["animation"] = deepcopy(animation)
    if hyperlink:
        out["hyperlink"] = str(hyperlink)
    return out


def _card_blocks(
    bid: str,
    block: dict[str, Any],
    rect: Rect,
    preset: dict[str, Any],
) -> list[dict[str, Any]]:
    grid_cfg = preset.get("grid") or {}
    pad = int(block.get("paddingPx", grid_cfg.get("cellPaddingPx", 14)))
    role = str(block.get("role") or "neutral")
    role_style = _role_style(preset, role)
    title = str(block.get("title") or bid)
    body = str(block.get("body") or block.get("text") or "")
    animation = block.get("animation")

    title_size = float(block.get("titleSizePt", grid_cfg.get("cardTitleSizePt", 20)))
    body_size = float(block.get("bodySizePt", grid_cfg.get("cardBodySizePt", 15)))
    title_h = int(block.get("titleHeightPx", max(34, title_size * 1.8)))
    inner_x = rect.x + pad
    inner_w = max(1, rect.w - 2 * pad)

    result: list[dict[str, Any]] = [{
        "type": "box",
        "decorative": True,
        "x": rect.x,
        "y": rect.y,
        "width": rect.w,
        "height": rect.h,
        "style": {
            "fillColor": role_style.get("fillColor", "FFFFFF"),
            "lineColor": role_style.get("borderColor", "CBD5E1"),
            "widthPt": float(block.get("borderWidthPt", 1.2)),
        },
        "z": int(block.get("boxZ", 4)),
    }]

    result.append(_text_block(
        text=title,
        x=inner_x,
        y=rect.y + pad,
        w=inner_w,
        h=min(title_h, max(1, rect.h - 2 * pad)),
        font_pt=title_size,
        color=role_style.get("titleColor", "172033"),
        bold=True,
        z=20,
        animation=animation,
        hyperlink=str(block.get("hyperlink")) if block.get("hyperlink") else None,
    ))

    if body:
        body_y = rect.y + pad + title_h + 6
        body_h = max(1, rect.bottom - pad - body_y)
        result.append(_text_block(
            text=body,
            x=inner_x,
            y=body_y,
            w=inner_w,
            h=body_h,
            font_pt=body_size,
            color=role_style.get("bodyColor", "475467"),
            bold=False,
            z=20,
        ))
    return result


def _metric_blocks(
    bid: str,
    block: dict[str, Any],
    rect: Rect,
    preset: dict[str, Any],
) -> list[dict[str, Any]]:
    grid_cfg = preset.get("grid") or {}
    pad = int(block.get("paddingPx", grid_cfg.get("cellPaddingPx", 14)))
    role_style = _role_style(preset, str(block.get("role") or "output"))
    value = str(block.get("value") or "")
    label = str(block.get("label") or block.get("title") or bid)
    value_size = float(block.get("valueSizePt", grid_cfg.get("metricValueSizePt", 30)))
    label_size = float(block.get("labelSizePt", grid_cfg.get("metricLabelSizePt", 14)))
    result = [{
        "type": "box",
        "decorative": True,
        "x": rect.x,
        "y": rect.y,
        "width": rect.w,
        "height": rect.h,
        "style": {
            "fillColor": role_style.get("fillColor", "FFFFFF"),
            "lineColor": role_style.get("borderColor", "7EB58E"),
            "widthPt": 1.4,
        },
        "z": 4,
    }]
    label_h = max(36, int(math.ceil(label_size * 2.2)))
    result.append(_text_block(
        text=value,
        x=rect.x + pad,
        y=rect.y + pad,
        w=max(1, rect.w - 2 * pad),
        h=max(1, rect.h - 2 * pad - label_h - 6),
        font_pt=value_size,
        color=role_style.get("titleColor", "246B3C"),
        bold=True,
        align="center",
        animation=block.get("animation"),
        hyperlink=str(block.get("hyperlink")) if block.get("hyperlink") else None,
    ))
    result.append(_text_block(
        text=label,
        x=rect.x + pad,
        y=max(rect.y + pad, rect.bottom - pad - label_h),
        w=max(1, rect.w - 2 * pad),
        h=label_h,
        font_pt=label_size,
        color=role_style.get("bodyColor", "344054"),
        align="center",
    ))
    return result


def _media_block(
    block: dict[str, Any],
    rect: Rect,
) -> dict[str, Any]:
    raw_type = str(block.get("type") or "image").lower()
    normalized = raw_type if raw_type in {"video", "audio"} else "image"
    media_type = raw_type if raw_type in {"gif", "svg", "video", "audio"} else str(block.get("mediaType") or "image")
    out = {
        "type": normalized,
        "path": str(block.get("path") or ""),
        "slot": str(block.get("_slot") or ""),
        "position": {
            "x": rect.x,
            "y": rect.y,
            "w": rect.w,
            "h": rect.h,
            "mode": "fixed",
        },
        "mediaType": media_type,
        "fit": str(block.get("fit") or "contain"),
        "z": int(block.get("z", 15)),
    }
    if block.get("poster"):
        out["poster"] = str(block["poster"])
    if block.get("caption"):
        out["caption"] = str(block["caption"])
    if block.get("animation"):
        out["animation"] = deepcopy(block["animation"])
    if block.get("hyperlink"):
        out["hyperlink"] = str(block["hyperlink"])
    return out


def _connector_blocks(
    connector: dict[str, Any],
    rects: dict[str, Rect],
    preset: dict[str, Any],
    index: int,
) -> list[dict[str, Any]]:
    source_id = str(connector.get("from") or "")
    target_id = str(connector.get("to") or "")
    if source_id not in rects or target_id not in rects:
        raise LayoutError(f"connector references unknown block: {source_id} -> {target_id}")

    a = rects[source_id]
    b = rects[target_id]
    conn_cfg = preset.get("connector") or {}
    color = str(connector.get("color") or conn_cfg.get("color") or "667085")
    width_pt = float(connector.get("widthPt", conn_cfg.get("widthPt", 2)))
    arrow = str(connector.get("arrowType") or conn_cfg.get("arrowType") or "triangle")

    acx, acy = a.x + a.w // 2, a.y + a.h // 2
    bcx, bcy = b.x + b.w // 2, b.y + b.h // 2

    blocks: list[dict[str, Any]] = []

    # Prefer orthogonal connectors for readability.
    vertical_overlap = max(a.y, b.y) <= min(a.bottom, b.bottom)
    horizontal_overlap = max(a.x, b.x) <= min(a.right, b.right)

    def line(x: int, y: int, w: int, h: int, *, begin: str = "none", end: str = "none") -> dict[str, Any]:
        return {
            "type": "line",
            "decorative": True,
            "x": x,
            "y": y,
            "width": max(0, w),
            "height": max(0, h),
            "style": {
                "color": color,
                "widthPt": width_pt,
                "beginArrowType": begin,
                "endArrowType": end,
            },
            "z": int(connector.get("z", 8)),
            "connectorId": f"connector-{index}",
        }

    if vertical_overlap:
        y = int((acy + bcy) / 2)
        if acx <= bcx:
            x1, x2 = a.right, b.x
            blocks.append(line(x1, y, max(1, x2 - x1), 0, end=arrow))
        else:
            x1, x2 = b.right, a.x
            blocks.append(line(x1, y, max(1, x2 - x1), 0, begin=arrow))
        return blocks

    if horizontal_overlap:
        x = int((acx + bcx) / 2)
        if acy <= bcy:
            y1, y2 = a.bottom, b.y
            blocks.append(line(x, y1, 0, max(1, y2 - y1), end=arrow))
        else:
            y1, y2 = b.bottom, a.y
            blocks.append(line(x, y1, 0, max(1, y2 - y1), begin=arrow))
        return blocks

    # Elbow: horizontal first, then vertical. The arrow sits on the target segment.
    if acx <= bcx:
        x_start = a.right
        x_mid = bcx
        blocks.append(line(x_start, acy, max(1, x_mid - x_start), 0))
    else:
        x_start = bcx
        x_end = a.x
        blocks.append(line(x_start, acy, max(1, x_end - x_start), 0))

    if acy <= bcy:
        blocks.append(line(bcx, acy, 0, max(1, b.y - acy), end=arrow))
    else:
        blocks.append(line(bcx, b.bottom, 0, max(1, acy - b.bottom), begin=arrow))
    return blocks


def compile_grid_spec(spec: dict[str, Any], spec_path: Path) -> dict[str, Any]:
    preset = _normalize_design(spec)
    canvas = deepcopy(spec.get("canvas") or {"preset": "16:9"})
    low: dict[str, Any] = {
        "canvas": canvas,
        "meta": {
            "source": "grid-deck",
            "designPreset": preset.get("name", "glanceable-scientific"),
            **(spec.get("meta") or {}),
        },
        "theme": deep_merge(preset.get("theme") or DEFAULT_THEME, spec.get("theme")),
        "slides": [],
    }

    width, height, _dpi = resolve_canvas(low)
    theme = low["theme"]
    base_content = layout_zones(width, height, theme, "title-body")["content"]
    grid_cfg = preset.get("grid") or {}

    for ordinal, slide in enumerate(spec.get("slides") or [], start=1):
        sid = str(slide.get("id") or f"slide-{ordinal}")
        grid = slide.get("grid")
        if not grid:
            raise LayoutError(f"{sid}: grid is required")

        gap = int(slide.get("gapPx", grid_cfg.get("gapPx", 18)))
        rects = _block_rects(
            grid,
            base_content,
            gap,
            slide.get("rowWeights"),
            slide.get("colWeights"),
        )

        high_blocks = slide.get("blocks") or {}
        missing = [bid for bid in rects if bid not in high_blocks]
        if missing:
            raise LayoutError(f"{sid}: missing block definitions: {missing}")

        regions = {
            f"grid_{bid}": {"x": r.x, "y": r.y, "width": r.w, "height": r.h}
            for bid, r in rects.items()
        }
        blocks: list[dict[str, Any]] = []

        for bid, rect in rects.items():
            block = deepcopy(high_blocks[bid])
            btype = str(block.get("type") or "card").lower()
            block["_slot"] = f"grid_{bid}"

            if btype == "card":
                blocks.extend(_card_blocks(bid, block, rect, preset))
            elif btype == "metric":
                blocks.extend(_metric_blocks(bid, block, rect, preset))
            elif btype in {"image", "gif", "svg", "video", "audio"}:
                if block.get("title"):
                    role_style = _role_style(preset, str(block.get("role") or "neutral"))
                    title_h = int(grid_cfg.get("cardTitleSizePt", 20) * 1.9)
                    pad = int(grid_cfg.get("cellPaddingPx", 14))
                    blocks.append({
                        "type": "box",
                        "decorative": True,
                        "x": rect.x, "y": rect.y, "width": rect.w, "height": rect.h,
                        "style": {
                            "fillColor": role_style.get("fillColor", "FFFFFF"),
                            "lineColor": role_style.get("borderColor", "CBD5E1"),
                            "widthPt": 1.2,
                        },
                        "z": 4,
                    })
                    blocks.append(_text_block(
                        text=str(block["title"]),
                        x=rect.x + pad,
                        y=rect.y + pad,
                        w=max(1, rect.w - 2 * pad),
                        h=title_h,
                        font_pt=float(grid_cfg.get("cardTitleSizePt", 20)),
                        color=role_style.get("titleColor", "172033"),
                        bold=True,
                        animation=block.get("animation"),
                    ))
                    media_rect = Rect(
                        rect.x + pad,
                        rect.y + pad + title_h + 4,
                        max(1, rect.w - 2 * pad),
                        max(1, rect.h - 2 * pad - title_h - 4),
                    )
                    blocks.append(_media_block(block, media_rect))
                else:
                    blocks.append(_media_block(block, rect))
            elif btype in {"text", "callout"}:
                card_like = dict(block)
                card_like["type"] = "card"
                blocks.extend(_card_blocks(bid, card_like, rect, preset))
            else:
                raise LayoutError(f"{sid}/{bid}: unsupported grid block type '{btype}'")

        for index, connector in enumerate(slide.get("connectors") or [], start=1):
            blocks.extend(_connector_blocks(connector, rects, preset, index))

        low["slides"].append({
            "id": sid,
            "title": str(slide.get("title") or ""),
            "layout": "title-body",
            "regions": regions,
            "blocks": blocks,
            **({"transition": deepcopy(slide["transition"])} if slide.get("transition") else {}),
        })

    if not low["slides"]:
        raise LayoutError("spec.slides is empty")
    return low


def main() -> int:
    ap = argparse.ArgumentParser(description="2D Grid DSL -> editable PowerPoint")
    sub = ap.add_subparsers(dest="command", required=True)

    build = sub.add_parser("build")
    build.add_argument("spec", type=Path)
    build.add_argument("--out-spec", type=Path, required=True)
    build.add_argument("--out-ir", type=Path)
    build.add_argument("--pptx", type=Path)
    build.add_argument("--strict", action="store_true")
    args = ap.parse_args()

    raw = json.loads(args.spec.read_text(encoding="utf-8"))
    pixel_spec = compile_grid_spec(raw, args.spec)
    args.out_spec.parent.mkdir(parents=True, exist_ok=True)
    args.out_spec.write_text(json.dumps(pixel_spec, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[grid-deck] pixel spec: {args.out_spec}")

    if args.out_ir or args.pptx:
        ir_path = args.out_ir or args.out_spec.with_suffix(".ir.json")
        compiler = DeckCompiler(pixel_spec, args.spec)
        deck, report = compiler.compile()
        ir_path.parent.mkdir(parents=True, exist_ok=True)
        ir_path.write_text(json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[grid-deck] IR: {ir_path}")

        if args.strict and report["errors"]:
            print(json.dumps(report["errors"], ensure_ascii=False, indent=2))
            return 1

        lint_code, lint_report = run_ir_lint(ir_path)
        print(f"[grid-deck] lint: {lint_report}")
        if args.strict and lint_code != 0:
            return 1

        if args.pptx:
            args.pptx.parent.mkdir(parents=True, exist_ok=True)
            build_pptx(ir_path, args.pptx)
            print(f"[grid-deck] PPTX: {args.pptx}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
