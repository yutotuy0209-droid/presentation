#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pixel_layout_engine.py

AIの仕事量を減らすための、1px単位・決定論的 PowerPoint レイアウトエンジン。

AI/人間は短い JSON spec に「内容」と「大まかなレイアウト」だけを書く。
本ツールが以下を担当する。

- 4:3 / 16:9 canvas の決定
- px -> PowerPoint native EMU 変換
- タイトル/本文/2カラム/画像の自動配置
- 文字量から必要高さを推定
- 最小フォントまでの自動縮小
- 長文の自動改ページ
- 1px単位の衝突回避
- スライド外はみ出し防止
- 画像のアスペクト比維持
- 表の高さ推定
- ページ番号/フッター
- IR v2(JSON)生成
- 既存 ir_to_pptx.js を呼んでPPTX生成
- レイアウトレポート出力

Python側は標準ライブラリのみ。
PPTX出力時のみ、このリポジトリ既存の Node/PptxGenJS writer を利用する。
"""
from __future__ import annotations

import argparse
import json
import math
import os
import struct
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

EMU_PER_INCH = 914400
DEFAULT_DPI = 96
PX_TO_EMU = EMU_PER_INCH / DEFAULT_DPI

PRESETS = {
    "4:3": (960, 720),
    "standard": (960, 720),
    "16:9": (1280, 720),
    "wide": (1280, 720),
}


class LayoutError(RuntimeError):
    pass


@dataclass(frozen=True)
class Rect:
    x: int
    y: int
    w: int
    h: int

    @property
    def right(self) -> int:
        return self.x + self.w

    @property
    def bottom(self) -> int:
        return self.y + self.h

    def intersects(self, other: "Rect", pad: int = 0) -> bool:
        return (
            self.x < other.right + pad
            and self.right + pad > other.x
            and self.y < other.bottom + pad
            and self.bottom + pad > other.y
        )

    def inside(self, outer: "Rect") -> bool:
        return (
            self.x >= outer.x
            and self.y >= outer.y
            and self.right <= outer.right
            and self.bottom <= outer.bottom
        )


def px_to_emu(value: int | float, dpi: int = DEFAULT_DPI) -> int:
    return int(round(float(value) * EMU_PER_INCH / dpi))


def clamp(v: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, v))


def deep_merge(base: dict[str, Any], override: dict[str, Any] | None) -> dict[str, Any]:
    out = dict(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


DEFAULT_THEME: dict[str, Any] = {
    "fontFamily": "Yu Gothic",
    "titleFontFamily": "Yu Gothic Light",
    "titleSizePt": 28,
    "bodySizePt": 20,
    "minBodySizePt": 14,
    "captionSizePt": 13,
    "footerSizePt": 10,
    "textColor": "111111",
    "mutedColor": "666666",
    "accentColor": "244B3A",
    "background": "FFFFFF",
    "lineColor": "D1D5DB",
    "marginLeftPx": 64,
    "marginRightPx": 64,
    "marginTopPx": 36,
    "marginBottomPx": 42,
    "titleHeightPx": 54,
    "titleGapPx": 22,
    "blockGapPx": 16,
    "columnGapPx": 26,
    "safePadPx": 4,
    "lineHeight": 1.34,
    "footerText": "",
    "showPageNumber": True,
}


def resolve_canvas(spec: dict[str, Any]) -> tuple[int, int, int]:
    canvas = spec.get("canvas") or {}
    dpi = int(canvas.get("dpi", DEFAULT_DPI))
    if canvas.get("width") and canvas.get("height"):
        return int(canvas["width"]), int(canvas["height"]), dpi
    preset = str(canvas.get("preset", "4:3")).lower()
    width, height = PRESETS.get(preset, PRESETS["4:3"])
    return width, height, dpi


def char_width_px(ch: str, font_px: float) -> float:
    if ch == "\t":
        return font_px * 2.2
    if ord(ch) < 128:
        if ch in " il.,:;!'|":
            return font_px * 0.32
        if ch in "MW@#%":
            return font_px * 0.86
        return font_px * 0.56
    return font_px


def wrap_text(text: str, width_px: int, font_pt: float, dpi: int = DEFAULT_DPI) -> list[str]:
    if not text:
        return [""]
    font_px = font_pt * dpi / 72.0
    max_w = max(1.0, float(width_px))
    lines: list[str] = []
    for para in text.replace("\r", "").split("\n"):
        if para == "":
            lines.append("")
            continue
        buf = ""
        used = 0.0
        for ch in para:
            cw = char_width_px(ch, font_px)
            if buf and used + cw > max_w:
                lines.append(buf.rstrip())
                buf = ch
                used = cw
            else:
                buf += ch
                used += cw
        lines.append(buf.rstrip())
    return lines


def required_text_height_px(
    text: str,
    width_px: int,
    font_pt: float,
    line_height: float,
    dpi: int = DEFAULT_DPI,
    pad_y: int = 4,
) -> int:
    line_count = max(1, len(wrap_text(text, width_px, font_pt, dpi)))
    font_px = font_pt * dpi / 72.0
    return int(math.ceil(line_count * font_px * line_height + 2 * pad_y))


def fit_text(
    text: str,
    width_px: int,
    max_height_px: int,
    preferred_pt: float,
    min_pt: float,
    line_height: float,
    dpi: int,
) -> tuple[float, int, bool]:
    size = float(preferred_pt)
    while size >= min_pt - 1e-9:
        h = required_text_height_px(text, width_px, size, line_height, dpi)
        if h <= max_height_px:
            return size, h, True
        size -= 0.5
    h = required_text_height_px(text, width_px, min_pt, line_height, dpi)
    return float(min_pt), h, h <= max_height_px


def split_text_for_height(
    text: str,
    width_px: int,
    height_px: int,
    font_pt: float,
    line_height: float,
    dpi: int,
) -> tuple[str, str]:
    """Split text deterministically so the first part fits the requested box."""
    if not text:
        return "", ""

    paragraphs = text.replace("\r", "").split("\n")
    first_parts: list[str] = []
    rest_parts: list[str] = []

    for idx, para in enumerate(paragraphs):
        trial = "\n".join(first_parts + [para])
        if required_text_height_px(trial, width_px, font_pt, line_height, dpi) <= height_px:
            first_parts.append(para)
            continue

        if not first_parts:
            # Split one long paragraph by character while preserving 1px geometry.
            lo, hi = 1, len(para)
            best = 0
            while lo <= hi:
                mid = (lo + hi) // 2
                trial = para[:mid]
                if required_text_height_px(trial, width_px, font_pt, line_height, dpi) <= height_px:
                    best = mid
                    lo = mid + 1
                else:
                    hi = mid - 1
            if best <= 0:
                return "", text
            first_parts.append(para[:best].rstrip())
            remaining = para[best:].lstrip()
            rest_parts = ([remaining] if remaining else []) + paragraphs[idx + 1 :]
        else:
            rest_parts = paragraphs[idx:]
        break
    else:
        return text, ""

    return "\n".join(first_parts).rstrip(), "\n".join(rest_parts).lstrip()


def read_png_size(path: Path) -> tuple[int, int] | None:
    with path.open("rb") as f:
        sig = f.read(24)
    if len(sig) >= 24 and sig[:8] == b"\x89PNG\r\n\x1a\n":
        return struct.unpack(">II", sig[16:24])
    return None


def read_jpeg_size(path: Path) -> tuple[int, int] | None:
    with path.open("rb") as f:
        data = f.read()
    if not data.startswith(b"\xff\xd8"):
        return None
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        i += 2
        if marker in (0xD8, 0xD9):
            continue
        if i + 2 > len(data):
            break
        length = int.from_bytes(data[i : i + 2], "big")
        if marker in list(range(0xC0, 0xC4)) + list(range(0xC5, 0xC8)) + list(range(0xC9, 0xCC)) + list(range(0xCD, 0xD0)):
            if i + 7 <= len(data):
                h = int.from_bytes(data[i + 3 : i + 5], "big")
                w = int.from_bytes(data[i + 5 : i + 7], "big")
                return w, h
        i += max(2, length)
    return None


def image_size(path: Path) -> tuple[int, int] | None:
    if not path.exists():
        return None
    try:
        return read_png_size(path) or read_jpeg_size(path)
    except OSError:
        return None


def contain_rect(source_wh: tuple[int, int] | None, target: Rect) -> Rect:
    if not source_wh or source_wh[0] <= 0 or source_wh[1] <= 0:
        return target
    sw, sh = source_wh
    scale = min(target.w / sw, target.h / sh)
    w = max(1, int(round(sw * scale)))
    h = max(1, int(round(sh * scale)))
    x = target.x + (target.w - w) // 2
    y = target.y + (target.h - h) // 2
    return Rect(x, y, w, h)


class CollisionMap:
    """Foreground rectangles with deterministic 1px placement."""

    def __init__(self, bounds: Rect, safe_pad: int = 4):
        self.bounds = bounds
        self.safe_pad = safe_pad
        self.items: list[tuple[str, Rect]] = []

    def colliders(self, rect: Rect) -> list[tuple[str, Rect]]:
        return [(eid, r) for eid, r in self.items if rect.intersects(r, self.safe_pad)]

    def is_free(self, rect: Rect) -> bool:
        return rect.inside(self.bounds) and not self.colliders(rect)

    def add(self, element_id: str, rect: Rect) -> None:
        if not rect.inside(self.bounds):
            raise LayoutError(f"{element_id}: outside canvas {rect}")
        self.items.append((element_id, rect))

    def place_nearest(self, element_id: str, preferred: Rect, zone: Rect) -> Rect:
        return self.place_controlled(
            element_id, preferred, zone, mode="flow", max_nudge_px=None
        )

    def place_controlled(
        self,
        element_id: str,
        preferred: Rect,
        zone: Rect,
        *,
        mode: str = "flow",
        max_nudge_px: int | None = None,
        axis: str = "both",
        allow_overlap: bool = False,
    ) -> Rect:
        """
        Hybrid placement.

        fixed:
            x/y/w/h are authoritative. Never move the object.
            Collision or out-of-zone placement is an error.
        prefer:
            x/y are a design intent. Search around the requested position by 1px,
            but never farther than max_nudge_px.
        flow:
            current automatic behavior. Search the semantic zone until free.
        """
        mode = str(mode or "flow").lower()
        axis = str(axis or "both").lower()
        if mode not in {"fixed", "prefer", "flow"}:
            raise LayoutError(f"{element_id}: unknown placement mode '{mode}'")
        if axis not in {"x", "y", "both"}:
            raise LayoutError(f"{element_id}: unknown nudge axis '{axis}'")

        if preferred.w <= 0 or preferred.h <= 0:
            raise LayoutError(f"{element_id}: invalid size {preferred.w}x{preferred.h}px")

        if mode == "fixed":
            if not preferred.inside(zone):
                raise LayoutError(
                    f"{element_id}: fixed box {preferred} is outside region {zone}"
                )
            hits = self.colliders(preferred)
            if hits and not allow_overlap:
                raise LayoutError(
                    f"{element_id}: fixed box overlaps {[eid for eid, _ in hits]}"
                )
            self.add(element_id, preferred)
            return preferred

        # prefer/flow may clamp to the allowed semantic region.
        preferred = Rect(
            clamp(preferred.x, zone.x, max(zone.x, zone.right - preferred.w)),
            clamp(preferred.y, zone.y, max(zone.y, zone.bottom - preferred.h)),
            preferred.w,
            preferred.h,
        )
        if self.is_free(preferred) or allow_overlap:
            self.add(element_id, preferred)
            return preferred

        if mode == "prefer":
            radius = max(0, int(max_nudge_px if max_nudge_px is not None else 24))
            for d in range(1, radius + 1):
                offsets: list[tuple[int, int]] = []
                if axis in {"y", "both"}:
                    offsets.extend([(0, d), (0, -d)])
                if axis in {"x", "both"}:
                    offsets.extend([(d, 0), (-d, 0)])
                if axis == "both":
                    offsets.extend([(d, d), (-d, d), (d, -d), (-d, -d)])

                for dx, dy in offsets:
                    cand = Rect(preferred.x + dx, preferred.y + dy, preferred.w, preferred.h)
                    if cand.inside(zone) and self.is_free(cand):
                        self.add(element_id, cand)
                        return cand
            raise LayoutError(
                f"{element_id}: preferred position cannot be resolved within {radius}px"
            )

        # flow: first preserve x and move downward exactly one pixel at a time.
        max_y = zone.bottom - preferred.h
        for y in range(preferred.y + 1, max_y + 1):
            cand = Rect(preferred.x, y, preferred.w, preferred.h)
            if self.is_free(cand):
                self.add(element_id, cand)
                return cand

        # Then exhaustive 1px scan inside the semantic zone.
        max_x = zone.right - preferred.w
        for y in range(zone.y, max_y + 1):
            for x in range(zone.x, max_x + 1):
                cand = Rect(x, y, preferred.w, preferred.h)
                if self.is_free(cand):
                    self.add(element_id, cand)
                    return cand

        raise LayoutError(
            f"{element_id}: no free {preferred.w}x{preferred.h}px area in zone {zone}"
        )


def element_ir(
    element_id: str,
    element_type: str,
    rect: Rect,
    dpi: int,
    *,
    text: str = "",
    style: dict[str, Any] | None = None,
    source: dict[str, Any] | None = None,
    runs: list[dict[str, Any]] | None = None,
    table: dict[str, Any] | None = None,
    z: int = 10,
) -> dict[str, Any]:
    out: dict[str, Any] = {
        "id": element_id,
        "name": element_id,
        "type": element_type,
        "z": z,
        "x": px_to_emu(rect.x, dpi),
        "y": px_to_emu(rect.y, dpi),
        "width": px_to_emu(rect.w, dpi),
        "height": px_to_emu(rect.h, dpi),
        "rotation": 0,
        "text": text,
        "style": style or {},
        "source": {
            "engine": "pixel-layout-v1",
            "pixel": {"x": rect.x, "y": rect.y, "width": rect.w, "height": rect.h},
            **(source or {}),
        },
    }
    if runs is not None:
        out["runs"] = runs
    if table is not None:
        out["table"] = table
    return out


def layout_zones(width: int, height: int, theme: dict[str, Any], layout: str) -> dict[str, Rect]:
    ml = int(theme["marginLeftPx"])
    mr = int(theme["marginRightPx"])
    mt = int(theme["marginTopPx"])
    mb = int(theme["marginBottomPx"])
    th = int(theme["titleHeightPx"])
    tg = int(theme["titleGapPx"])
    gap = int(theme["columnGapPx"])

    title = Rect(ml, mt, width - ml - mr, th)
    content_y = title.bottom + tg
    content = Rect(ml, content_y, width - ml - mr, height - mb - content_y)

    if layout in {"two-column", "image-left", "image-right"}:
        left_w = (content.w - gap) // 2
        right_w = content.w - gap - left_w
        left = Rect(content.x, content.y, left_w, content.h)
        right = Rect(content.x + left_w + gap, content.y, right_w, content.h)
    else:
        left = right = content

    return {"title": title, "content": content, "left": left, "right": right}


def normalize_block(block: Any) -> dict[str, Any]:
    if isinstance(block, str):
        return {"type": "text", "text": block}
    if not isinstance(block, dict):
        return {"type": "text", "text": str(block)}
    return dict(block)


def choose_slot(layout: str, block: dict[str, Any], index: int) -> str:
    explicit = block.get("slot")
    if isinstance(explicit, str) and explicit:
        return explicit
    if layout == "image-left":
        return "left" if block.get("type") == "image" else "right"
    if layout == "image-right":
        return "right" if block.get("type") == "image" else "left"
    if layout == "two-column":
        return "left" if index % 2 == 0 else "right"
    return "content"


def rect_from_mapping(
    data: dict[str, Any] | None,
    fallback: Rect,
) -> Rect:
    data = data or {}
    return Rect(
        int(data.get("x", fallback.x)),
        int(data.get("y", fallback.y)),
        int(data.get("w", data.get("width", fallback.w))),
        int(data.get("h", data.get("height", fallback.h))),
    )


def block_position(block: dict[str, Any]) -> dict[str, Any]:
    """
    Read either the compact top-level x/y/width/height form or
    position: {x,y,w,h,mode,maxNudgePx,axis}.
    """
    pos = dict(block.get("position") or {})
    for key in ("x", "y", "width", "height", "w", "h"):
        if key in block and key not in pos:
            pos[key] = block[key]
    if "placement" in block and "mode" not in pos:
        pos["mode"] = block["placement"]
    return pos


def placement_options(block: dict[str, Any]) -> dict[str, Any]:
    pos = block_position(block)
    return {
        "mode": str(pos.get("mode", "flow")).lower(),
        "max_nudge_px": int(pos.get("maxNudgePx", pos.get("max_nudge_px", 24))),
        "axis": str(pos.get("axis", "both")).lower(),
        "allow_overlap": bool(pos.get("allowOverlap", False)),
    }


def make_text_element(
    eid: str,
    text: str,
    rect: Rect,
    font_pt: float,
    theme: dict[str, Any],
    dpi: int,
    style_override: dict[str, Any] | None = None,
    z: int = 20,
) -> dict[str, Any]:
    style = deep_merge(
        {
            "fontFamily": theme["fontFamily"],
            "fontSizePt": font_pt,
            "color": theme["textColor"],
            "bold": False,
            "align": "left",
            "valign": "top",
            "marginPx": 0,
            "lineHeight": theme["lineHeight"],
        },
        style_override,
    )
    runs = [{
        "text": text,
        "fontFamily": style.get("fontFamily", theme["fontFamily"]),
        "fontSizePt": float(style.get("fontSizePt", font_pt)),
        "bold": bool(style.get("bold", False)),
        "italic": bool(style.get("italic", False)),
        "color": style.get("color", theme["textColor"]),
    }]
    return element_ir(eid, "sp", rect, dpi, text=text, style=style, runs=runs, z=z)


def build_table_ir(rows: list[list[Any]], font_pt: float, theme: dict[str, Any]) -> dict[str, Any]:
    normalized_rows = []
    col_count = max((len(r) for r in rows), default=1)
    for ridx, row in enumerate(rows):
        cells = []
        for cell in row:
            txt = str(cell.get("text", "")) if isinstance(cell, dict) else str(cell)
            bold = bool(cell.get("bold", False)) if isinstance(cell, dict) else ridx == 0
            cells.append({
                "text": txt,
                "runs": [{
                    "text": txt,
                    "fontFamily": theme["fontFamily"],
                    "fontSizePt": font_pt,
                    "bold": bold,
                    "color": theme["textColor"],
                }],
                "gridSpan": 1,
                "rowSpan": 1,
            })
        while len(cells) < col_count:
            cells.append({"text": "", "runs": [], "gridSpan": 1, "rowSpan": 1})
        normalized_rows.append({"height": 0, "cells": cells})
    return {"colWidths": [], "rows": normalized_rows}


def estimate_table_height(rows: list[list[Any]], font_pt: float, dpi: int) -> int:
    font_px = font_pt * dpi / 72.0
    return max(42, int(math.ceil(max(1, len(rows)) * max(30, font_px * 1.65))))


class DeckCompiler:
    def __init__(self, spec: dict[str, Any], spec_path: Path):
        self.spec = spec
        self.spec_path = spec_path.resolve()
        self.base_dir = self.spec_path.parent
        self.width, self.height, self.dpi = resolve_canvas(spec)
        self.theme = deep_merge(DEFAULT_THEME, spec.get("theme"))
        self.deck: dict[str, Any] = {
            "version": 2,
            "unit": "emu",
            "meta": {
                "source": "pixel-layout-engine",
                "spec": str(self.spec_path),
                **(spec.get("meta") or {}),
            },
            "slideSize": {
                "width": px_to_emu(self.width, self.dpi),
                "height": px_to_emu(self.height, self.dpi),
            },
            "slides": [],
        }
        self.report: dict[str, Any] = {
            "engine": "pixel-layout-v1",
            "canvas": {"widthPx": self.width, "heightPx": self.height, "dpi": self.dpi},
            "slides": [],
            "errors": [],
            "warnings": [],
        }

    def _new_slide(
        self,
        slide_id: str,
        title: str,
        layout: str,
        slide_spec: dict[str, Any] | None = None,
        continuation: bool = False,
    ) -> dict[str, Any]:
        slide_spec = slide_spec or {}
        zones = layout_zones(self.width, self.height, self.theme, layout)

        # Optional named regions let a human/AI control composition without
        # micromanaging every object coordinate.
        for region_name, raw in (slide_spec.get("regions") or {}).items():
            if not isinstance(raw, dict):
                continue
            zones[str(region_name)] = rect_from_mapping(raw, zones["content"])

        if isinstance(slide_spec.get("titleBox"), dict):
            zones["title"] = rect_from_mapping(slide_spec["titleBox"], zones["title"])
        collision = CollisionMap(Rect(0, 0, self.width, self.height), int(self.theme["safePadPx"]))
        slide = {
            "id": slide_id,
            "number": len(self.deck["slides"]) + 1,
            "background": self.theme["background"],
            "elements": [],
            "_zones": zones,
            "_collision": collision,
            "_cursor": {
                "content": zones["content"].y,
                "left": zones["left"].y,
                "right": zones["right"].y,
            },
        }

        if title:
            title_text = title + ("（続き）" if continuation else "")
            tr = zones["title"]
            collision.add(f"{slide_id}-title", tr)
            slide["elements"].append(
                make_text_element(
                    f"{slide_id}-title",
                    title_text,
                    tr,
                    float(self.theme["titleSizePt"]),
                    self.theme,
                    self.dpi,
                    {
                        "fontFamily": self.theme["titleFontFamily"],
                        "bold": False,
                        "valign": "middle",
                    },
                    z=50,
                )
            )
            # Small accent line: decorative and does not occupy collision space.
            line_rect = Rect(tr.x, tr.bottom + 8, tr.w, 1)
            slide["elements"].append(
                element_ir(
                    f"{slide_id}-title-line",
                    "line",
                    line_rect,
                    self.dpi,
                    style={"color": self.theme["accentColor"], "widthPt": 1.5},
                    z=5,
                )
            )
        return slide

    def _push_slide(self, slide: dict[str, Any]) -> None:
        zones = slide.pop("_zones")
        collision: CollisionMap = slide.pop("_collision")
        slide.pop("_cursor")
        self.deck["slides"].append(slide)
        self.report["slides"].append({
            "id": slide["id"],
            "elementCount": len(slide["elements"]),
            "occupiedPx": [
                {"id": eid, "x": r.x, "y": r.y, "width": r.w, "height": r.h}
                for eid, r in collision.items
            ],
            "zones": {
                k: {"x": v.x, "y": v.y, "width": v.w, "height": v.h}
                for k, v in zones.items()
            },
        })

    def _resolve_path(self, raw: str) -> Path:
        p = Path(raw)
        return p.resolve() if p.is_absolute() else (self.base_dir / p).resolve()

    def compile_slide_spec(self, slide_spec: dict[str, Any], ordinal: int) -> None:
        base_id = str(slide_spec.get("id") or f"slide-{ordinal}")
        title = str(slide_spec.get("title") or "")

        # Token節約用shorthand:
        # AIは blocks を細かく書かなくても body / bullets / image / table だけでよい。
        raw_blocks = slide_spec.get("blocks")
        if raw_blocks is None:
            raw_blocks = []
            if slide_spec.get("body"):
                raw_blocks.append({"type": "text", "text": str(slide_spec["body"])})
            if slide_spec.get("bullets"):
                bullet_text = "\n".join(f"• {x}" for x in slide_spec["bullets"])
                raw_blocks.append({"type": "text", "text": bullet_text})
            if slide_spec.get("image"):
                image_value = slide_spec["image"]
                if isinstance(image_value, dict):
                    raw_blocks.append({"type": "image", **image_value})
                else:
                    raw_blocks.append({"type": "image", "path": str(image_value)})
            if slide_spec.get("table"):
                raw_blocks.append({"type": "table", "rows": slide_spec["table"]})

        blocks = [normalize_block(b) for b in raw_blocks]
        requested_layout = slide_spec.get("layout")
        if not requested_layout or str(requested_layout).lower() == "auto":
            has_image = any(str(b.get("type", "")).lower() == "image" for b in blocks)
            has_text = any(str(b.get("type", "text")).lower() in {"text", "bullets", "callout", "metric"} for b in blocks)
            if int(slide_spec.get("columns", 1)) == 2:
                layout = "two-column"
            elif has_image and has_text:
                layout = "image-right"
            else:
                layout = "title-body"
        else:
            layout = str(requested_layout)

        slide_index = 1
        slide = self._new_slide(base_id, title, layout, slide_spec, continuation=False)
        zones: dict[str, Rect] = slide["_zones"]
        collision: CollisionMap = slide["_collision"]
        cursors: dict[str, int] = slide["_cursor"]

        def continue_slide() -> None:
            nonlocal slide_index, slide, zones, collision, cursors
            self._push_slide(slide)
            slide_index += 1
            sid = f"{base_id}-cont-{slide_index}"
            slide = self._new_slide(sid, title, layout, slide_spec, continuation=True)
            zones = slide["_zones"]
            collision = slide["_collision"]
            cursors = slide["_cursor"]

        for block_index, block in enumerate(blocks):
            btype = str(block.get("type", "text")).lower()

            # Semantic blocks are compiled down to normal text so AI does not
            # need to micromanage typography/geometry.
            if btype == "bullets":
                items = block.get("items") or block.get("bullets") or []
                block["text"] = "\n".join(f"• {x}" for x in items)
                btype = "text"
            elif btype == "callout":
                block["style"] = deep_merge(
                    {"bold": True, "color": self.theme["accentColor"]},
                    block.get("style"),
                )
                btype = "text"
            elif btype == "metric":
                label = str(block.get("label") or "")
                value = str(block.get("value") or "")
                block["text"] = f"{value}\n{label}".strip()
                block["fontSizePt"] = block.get("fontSizePt", max(26, self.theme["bodySizePt"] + 6))
                block["style"] = deep_merge(
                    {"bold": True, "color": self.theme["accentColor"], "align": "center"},
                    block.get("style"),
                )
                btype = "text"

            if btype == "spacer":
                slot = choose_slot(layout, block, block_index)
                cursors[slot] += int(block.get("height", self.theme["blockGapPx"]))
                continue

            if btype in {"box", "line"} and block.get("decorative", True):
                # Decorative items are permitted to overlap and do not consume layout space.
                x = int(block.get("x", zones["content"].x))
                y = int(block.get("y", zones["content"].y))
                w = int(block.get("width", zones["content"].w))
                h = int(block.get("height", 1 if btype == "line" else 40))
                rect = Rect(x, y, w, h)
                slide["elements"].append(
                    element_ir(
                        f"{base_id}-b{block_index}",
                        btype,
                        rect,
                        self.dpi,
                        style=block.get("style") or {},
                        z=int(block.get("z", 2)),
                    )
                )
                continue

            slot = choose_slot(layout, block, block_index)
            if slot not in zones:
                raise LayoutError(
                    f"{base_id}-b{block_index}: unknown slot/region '{slot}'"
                )
            zone = zones[slot]
            if slot not in cursors:
                cursors[slot] = zone.y

            gap = int(block.get("gapAfter", self.theme["blockGapPx"]))
            pos = block_position(block)
            width = int(pos.get("w", pos.get("width", block.get("width", zone.w))))
            width = min(width, zone.w)
            preferred_x = int(pos.get("x", block.get("x", zone.x)))
            preferred_y = int(pos.get("y", block.get("y", cursors[slot])))
            requested_h = pos.get("h", pos.get("height"))
            place_opts = placement_options(block)

            if btype == "text":
                text = str(block.get("text") or "")
                pref = float(block.get("fontSizePt", self.theme["bodySizePt"]))
                min_pt = float(block.get("minFontSizePt", self.theme["minBodySizePt"]))
                max_available = (
                    int(requested_h)
                    if requested_h is not None
                    else zone.bottom - preferred_y
                )
                if max_available <= 8:
                    continue_slide()
                    zone = zones[slot]
                    preferred_y = cursors[slot]
                    max_available = zone.bottom - preferred_y

                font_pt, need_h, fits = fit_text(
                    text, width, max_available, pref, min_pt,
                    float(self.theme["lineHeight"]), self.dpi,
                )

                if not fits:
                    if place_opts["mode"] == "fixed":
                        raise LayoutError(
                            f"{base_id}-b{block_index}: text does not fit fixed "
                            f"{width}x{max_available}px box even at {min_pt}pt"
                        )
                    first, rest = split_text_for_height(
                        text, width, max_available, min_pt,
                        float(self.theme["lineHeight"]), self.dpi,
                    )
                    if not first:
                        continue_slide()
                        zone = zones[slot]
                        preferred_y = cursors[slot]
                        max_available = zone.bottom - preferred_y
                        first, rest = split_text_for_height(
                            text, width, max_available, min_pt,
                            float(self.theme["lineHeight"]), self.dpi,
                        )
                    if not first:
                        raise LayoutError(f"{base_id}: text block cannot fit even on empty slide")

                    part = 1
                    remaining = rest
                    current = first
                    while True:
                        zone = zones[slot]
                        preferred_y = cursors[slot]
                        max_available = zone.bottom - preferred_y
                        h = required_text_height_px(
                            current, width, min_pt,
                            float(self.theme["lineHeight"]), self.dpi,
                        )
                        rect = collision.place_controlled(
                            f"{base_id}-b{block_index}-p{part}",
                            Rect(preferred_x, preferred_y, width, min(h, max_available)),
                            zone,
                            mode="flow",
                        )
                        slide["elements"].append(
                            make_text_element(
                                f"{base_id}-b{block_index}-p{part}",
                                current, rect, min_pt, self.theme, self.dpi,
                                block.get("style"), z=int(block.get("z", 20)),
                            )
                        )
                        cursors[slot] = rect.bottom + gap
                        if not remaining:
                            break
                        continue_slide()
                        zone = zones[slot]
                        preferred_x = zone.x
                        preferred_y = cursors[slot]
                        max_available = zone.bottom - preferred_y
                        current, remaining = split_text_for_height(
                            remaining, width, max_available, min_pt,
                            float(self.theme["lineHeight"]), self.dpi,
                        )
                        if not current:
                            raise LayoutError(f"{base_id}: continuation split failed")
                        part += 1
                    continue

                # If the caller supplied h/height, preserve that design box
                # in both fixed and prefer modes. This keeps coordinate control
                # meaningful while still allowing prefer-mode nudging.
                text_h = int(requested_h) if requested_h is not None else need_h
                rect = collision.place_controlled(
                    f"{base_id}-b{block_index}",
                    Rect(preferred_x, preferred_y, width, text_h),
                    zone,
                    **place_opts,
                )
                slide["elements"].append(
                    make_text_element(
                        f"{base_id}-b{block_index}",
                        text, rect, font_pt, self.theme, self.dpi,
                        block.get("style"), z=int(block.get("z", 20)),
                    )
                )
                cursors[slot] = rect.bottom + gap

            elif btype == "image":
                raw_path = str(block.get("path") or "")
                if not raw_path:
                    self.report["warnings"].append({
                        "slide": base_id, "block": block_index, "rule": "missing-image-path"
                    })
                    continue
                img_path = self._resolve_path(raw_path)
                if not img_path.exists():
                    self.report["warnings"].append({
                        "slide": base_id, "block": block_index,
                        "rule": "image-not-found", "path": str(img_path)
                    })
                max_available = zone.bottom - preferred_y
                desired_h = int(
                    requested_h
                    if requested_h is not None
                    else block.get("height", min(max_available, int(zone.h * 0.58)))
                )
                if desired_h <= 8:
                    continue_slide()
                    zone = zones[slot]
                    preferred_y = cursors[slot]
                    max_available = zone.bottom - preferred_y
                    desired_h = int(block.get("height", min(max_available, int(zone.h * 0.58))))
                target = Rect(preferred_x, preferred_y, width, min(desired_h, max_available))
                actual = contain_rect(image_size(img_path), target)
                rect = collision.place_controlled(
                    f"{base_id}-b{block_index}", actual, zone, **place_opts
                )
                slide["elements"].append(
                    element_ir(
                        f"{base_id}-b{block_index}", "pic", rect, self.dpi,
                        source={"path": str(img_path)},
                        style=block.get("style") or {},
                        z=int(block.get("z", 15)),
                    )
                )
                cursors[slot] = rect.bottom + gap
                caption = str(block.get("caption") or "")
                if caption:
                    cap_pt = float(block.get("captionSizePt", self.theme["captionSizePt"]))
                    cap_h = required_text_height_px(
                        caption, width, cap_pt,
                        float(self.theme["lineHeight"]), self.dpi,
                    )
                    if cursors[slot] + cap_h > zone.bottom:
                        continue_slide()
                        zone = zones[slot]
                    cap_rect = collision.place_nearest(
                        f"{base_id}-b{block_index}-caption",
                        Rect(zone.x, cursors[slot], width, cap_h), zone,
                    )
                    slide["elements"].append(
                        make_text_element(
                            f"{base_id}-b{block_index}-caption",
                            caption, cap_rect, cap_pt, self.theme, self.dpi,
                            {"color": self.theme["mutedColor"], "align": "center"},
                            z=20,
                        )
                    )
                    cursors[slot] = cap_rect.bottom + gap

            elif btype == "table":
                rows = block.get("rows") or []
                font_pt = float(block.get("fontSizePt", max(12, self.theme["bodySizePt"] - 4)))
                desired_h = int(block.get("height", estimate_table_height(rows, font_pt, self.dpi)))
                max_available = zone.bottom - preferred_y
                if desired_h > max_available:
                    continue_slide()
                    zone = zones[slot]
                    preferred_y = cursors[slot]
                    max_available = zone.bottom - preferred_y
                table_h = (
                    int(requested_h)
                    if requested_h is not None
                    else min(desired_h, max_available)
                )
                rect = collision.place_controlled(
                    f"{base_id}-b{block_index}",
                    Rect(preferred_x, preferred_y, width, table_h),
                    zone,
                    **place_opts,
                )
                slide["elements"].append(
                    element_ir(
                        f"{base_id}-b{block_index}",
                        "graphicFrame", rect, self.dpi,
                        table=build_table_ir(rows, font_pt, self.theme),
                        style=deep_merge({
                            "fontSizePt": font_pt,
                            "borderColor": self.theme["lineColor"],
                            "fillColor": "FFFFFF",
                        }, block.get("style")),
                        z=int(block.get("z", 20)),
                    )
                )
                cursors[slot] = rect.bottom + gap

            else:
                self.report["warnings"].append({
                    "slide": base_id, "block": block_index,
                    "rule": "unsupported-block", "type": btype
                })

        self._push_slide(slide)

    def add_footer_and_page_numbers(self) -> None:
        total = len(self.deck["slides"])
        footer_text = str(self.theme.get("footerText") or "")
        size = float(self.theme["footerSizePt"])
        y = self.height - int(self.theme["marginBottomPx"]) + 10
        h = max(18, int(size * self.dpi / 72 * 1.3))
        left = int(self.theme["marginLeftPx"])
        right = int(self.theme["marginRightPx"])

        for idx, slide in enumerate(self.deck["slides"], start=1):
            if footer_text:
                r = Rect(left, y, self.width - left - right - 90, h)
                slide["elements"].append(
                    make_text_element(
                        f"{slide['id']}-footer", footer_text, r, size,
                        self.theme, self.dpi,
                        {"color": self.theme["mutedColor"], "valign": "middle"},
                        z=60,
                    )
                )
            if self.theme.get("showPageNumber", True):
                r = Rect(self.width - right - 78, y, 78, h)
                slide["elements"].append(
                    make_text_element(
                        f"{slide['id']}-page", f"{idx} / {total}", r, size,
                        self.theme, self.dpi,
                        {"color": self.theme["mutedColor"], "align": "right", "valign": "middle"},
                        z=60,
                    )
                )

    def final_validate(self) -> None:
        errors = []
        for slide in self.deck["slides"]:
            foreground = []
            for el in slide["elements"]:
                src = el.get("source") or {}
                px = src.get("pixel")
                if not px:
                    continue
                rect = Rect(int(px["x"]), int(px["y"]), int(px["width"]), int(px["height"]))
                if not rect.inside(Rect(0, 0, self.width, self.height)):
                    errors.append({
                        "slide": slide["id"], "element": el["id"], "rule": "overflow"
                    })
                if el["type"] in {"sp", "pic", "graphicFrame"} and not el["id"].endswith(("-footer", "-page")):
                    for prev_id, prev_rect in foreground:
                        if rect.intersects(prev_rect, int(self.theme["safePadPx"])):
                            errors.append({
                                "slide": slide["id"],
                                "elements": [prev_id, el["id"]],
                                "rule": "overlap",
                            })
                    foreground.append((el["id"], rect))
        self.report["errors"].extend(errors)
        self.report["ok"] = not self.report["errors"]
        self.report["slideCount"] = len(self.deck["slides"])
        self.report["elementCount"] = sum(len(s["elements"]) for s in self.deck["slides"])

    def compile(self) -> tuple[dict[str, Any], dict[str, Any]]:
        slides = self.spec.get("slides") or []
        if not slides:
            raise LayoutError("spec.slides is empty")
        for i, slide_spec in enumerate(slides, start=1):
            self.compile_slide_spec(slide_spec, i)
        self.add_footer_and_page_numbers()
        self.final_validate()
        return self.deck, self.report


def build_pptx(ir_path: Path, pptx_path: Path) -> None:
    repo_root = Path(__file__).resolve().parent.parent
    writer = repo_root / "src" / "ir_to_pptx.js"
    report = pptx_path.with_suffix(".generation-report.json")
    cmd = ["node", str(writer), str(ir_path), str(pptx_path), str(report)]
    subprocess.run(cmd, cwd=repo_root, check=True)


def run_ir_lint(ir_path: Path) -> tuple[int, str]:
    """Run the repository's common IR lint as the second safety gate."""
    repo_root = Path(__file__).resolve().parent.parent
    lint = repo_root / "src" / "lint_ir.js"
    proc = subprocess.run(
        ["node", str(lint), str(ir_path)],
        cwd=repo_root,
        text=True,
        capture_output=True,
        check=False,
    )
    lint_report = ir_path.with_suffix(".lint.json")
    lint_report.write_text(proc.stdout or "[]", encoding="utf-8")
    if proc.stderr:
        sys.stderr.write(proc.stderr)
    return proc.returncode, str(lint_report)


def render_with_powerpoint(pptx_path: Path, output_dir: Path) -> None:
    """Optional Windows/PowerPoint render for final visual inspection."""
    repo_root = Path(__file__).resolve().parent.parent
    renderer = repo_root / "src" / "render_pptx_powerpoint.ps1"
    output_dir.mkdir(parents=True, exist_ok=True)
    candidates = ["pwsh", "powershell"]
    last_error: Exception | None = None
    for shell in candidates:
        try:
            subprocess.run(
                [
                    shell,
                    "-NoProfile",
                    "-ExecutionPolicy", "Bypass",
                    "-File", str(renderer),
                    "-InputPptx", str(pptx_path),
                    "-OutputDir", str(output_dir),
                ],
                cwd=repo_root,
                check=True,
            )
            return
        except (FileNotFoundError, subprocess.CalledProcessError) as exc:
            last_error = exc
    raise RuntimeError(f"PowerPoint render failed: {last_error}")


def main() -> int:
    ap = argparse.ArgumentParser(
        description="1px単位のPowerPoint自動レイアウトエンジン"
    )
    sub = ap.add_subparsers(dest="command", required=True)

    build = sub.add_parser("build", help="compact JSON spec -> IR/PPTX")
    build.add_argument("spec", type=Path)
    build.add_argument("--out-ir", type=Path, required=True)
    build.add_argument("--pptx", type=Path)
    build.add_argument("--report", type=Path)
    build.add_argument("--strict", action="store_true")
    build.add_argument(
        "--skip-lint",
        action="store_true",
        help="共通IR lintを省略する（通常は非推奨）",
    )
    build.add_argument(
        "--render-dir",
        type=Path,
        help="Windows + PowerPoint環境で最終PNGレンダリングする出力先",
    )

    args = ap.parse_args()

    if args.command == "build":
        spec_path: Path = args.spec
        if not spec_path.exists():
            print(f"ERROR: spec not found: {spec_path}", file=sys.stderr)
            return 2
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
        compiler = DeckCompiler(spec, spec_path)
        try:
            deck, report = compiler.compile()
        except LayoutError as exc:
            print(f"LAYOUT ERROR: {exc}", file=sys.stderr)
            return 1

        args.out_ir.parent.mkdir(parents=True, exist_ok=True)
        args.out_ir.write_text(
            json.dumps(deck, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        report_path = args.report or args.out_ir.with_suffix(".layout-report.json")
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        print(
            f"[pixel-layout] slides={report['slideCount']} "
            f"elements={report['elementCount']} "
            f"errors={len(report['errors'])} warnings={len(report['warnings'])}"
        )
        print(f"[pixel-layout] IR: {args.out_ir}")
        print(f"[pixel-layout] report: {report_path}")

        if args.strict and report["errors"]:
            return 1

        if not args.skip_lint:
            lint_code, lint_report = run_ir_lint(args.out_ir)
            print(f"[pixel-layout] lint: {lint_report}")
            if lint_code != 0:
                print("[pixel-layout] ERROR: common IR lint failed", file=sys.stderr)
                return 1

        if args.pptx:
            args.pptx.parent.mkdir(parents=True, exist_ok=True)
            build_pptx(args.out_ir, args.pptx)
            print(f"[pixel-layout] PPTX: {args.pptx}")

            if args.render_dir:
                try:
                    render_with_powerpoint(args.pptx, args.render_dir)
                    print(f"[pixel-layout] rendered PNGs: {args.render_dir}")
                except RuntimeError as exc:
                    print(f"[pixel-layout] render warning: {exc}", file=sys.stderr)
                    if args.strict:
                        return 1
        elif args.render_dir:
            print("ERROR: --render-dir requires --pptx", file=sys.stderr)
            return 2

        return 0

    return 2


if __name__ == "__main__":
    raise SystemExit(main())
