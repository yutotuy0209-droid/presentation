#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Single-entry presentation pipeline.

Input:
- Markdown: # deck title, ## slide title, bullets, paragraphs, ![caption](image)
- JSON: existing compact spec is passed through with small defaults
- Plain text: first non-empty line becomes deck title, paragraphs become slides

Output:
source -> compact spec -> pixel layout IR -> lint -> PPTX -> optional PowerPoint render

The orchestration intentionally contains no LLM dependency. AI tools can create or edit
one concise Markdown file, while this script keeps the rendering path deterministic.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
from pathlib import Path
import re
import subprocess
import sys
from typing import Any

IMAGE_RE = re.compile(r"^!\[(?P<caption>[^\]]*)\]\((?P<path>[^)]+)\)\s*$")
BULLET_RE = re.compile(r"^\s*(?:[-*+]\s+|\d+[.)]\s+)(?P<text>.+?)\s*$")


def _nonempty(lines: list[str]) -> list[str]:
    return [line.strip() for line in lines if line.strip()]


def _section_to_slides(
    title: str,
    lines: list[str],
    *,
    source_dir: Path,
    max_bullets: int,
) -> list[dict[str, Any]]:
    bullets: list[str] = []
    paragraphs: list[str] = []
    images: list[dict[str, str]] = []
    paragraph_buf: list[str] = []

    def flush_paragraph() -> None:
        if paragraph_buf:
            paragraphs.append(" ".join(paragraph_buf).strip())
            paragraph_buf.clear()

    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            flush_paragraph()
            continue

        image_match = IMAGE_RE.match(line.strip())
        if image_match:
            flush_paragraph()
            raw_path = image_match.group("path").strip().strip('"').strip("'")
            p = Path(raw_path)
            if not p.is_absolute():
                p = (source_dir / p).resolve()
            images.append({
                "path": str(p),
                "caption": image_match.group("caption").strip(),
            })
            continue

        bullet_match = BULLET_RE.match(line)
        if bullet_match:
            flush_paragraph()
            bullets.append(bullet_match.group("text").strip())
            continue

        if line.lstrip().startswith(">"):
            flush_paragraph()
            paragraphs.append(line.lstrip()[1:].strip())
            continue

        paragraph_buf.append(line.strip())

    flush_paragraph()

    bullet_chunks = [
        bullets[i : i + max_bullets]
        for i in range(0, len(bullets), max_bullets)
    ] or [[]]
    slide_count = max(len(bullet_chunks), len(images), 1)
    slides: list[dict[str, Any]] = []

    for idx in range(slide_count):
        blocks: list[dict[str, Any]] = []
        if idx == 0 and paragraphs:
            blocks.append({"type": "text", "text": "\n\n".join(paragraphs)})
        if idx < len(bullet_chunks) and bullet_chunks[idx]:
            blocks.append({"type": "bullets", "items": bullet_chunks[idx]})
        if idx < len(images):
            image = dict(images[idx])
            blocks.append({"type": "image", **image})

        if not blocks:
            blocks.append({"type": "text", "text": ""})

        has_image = any(block.get("type") == "image" for block in blocks)
        has_text = any(block.get("type") in {"text", "bullets"} and (
            block.get("text") or block.get("items")
        ) for block in blocks)
        layout = "image-right" if has_image and has_text else "auto"

        suffix = "" if idx == 0 else f"（続き {idx + 1}）"
        slides.append({
            "title": f"{title}{suffix}",
            "layout": layout,
            "blocks": blocks,
        })

    return slides


def parse_markdown(text: str, *, source_dir: Path, max_bullets: int = 5) -> tuple[str, list[dict[str, Any]]]:
    deck_title = ""
    sections: list[tuple[str, list[str]]] = []
    current_title: str | None = None
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_title, current_lines
        if current_title is not None:
            sections.append((current_title, current_lines))
        current_lines = []

    for raw in text.replace("\r", "").split("\n"):
        if raw.startswith("# "):
            if not deck_title:
                deck_title = raw[2:].strip()
            elif current_title is None:
                current_title = raw[2:].strip()
            else:
                current_lines.append(raw)
            continue

        if raw.startswith("## "):
            flush()
            current_title = raw[3:].strip()
            continue

        if raw.strip() == "---":
            if current_title is not None and _nonempty(current_lines):
                flush()
                current_title = "続き"
            continue

        if current_title is None:
            if raw.strip():
                current_title = "概要"
                current_lines.append(raw)
        else:
            current_lines.append(raw)

    flush()

    if not deck_title:
        deck_title = source_dir.name or "Presentation"

    slides: list[dict[str, Any]] = []
    for title, lines in sections:
        slides.extend(_section_to_slides(
            title or "概要", lines, source_dir=source_dir, max_bullets=max_bullets
        ))

    if not slides:
        slides = [{"title": deck_title, "body": "内容がありません。"}]

    return deck_title, slides


def parse_plain_text(text: str, *, max_bullets: int = 5) -> tuple[str, list[dict[str, Any]]]:
    lines = text.replace("\r", "").split("\n")
    nonempty = _nonempty(lines)
    if not nonempty:
        return "Presentation", [{"title": "概要", "body": "内容がありません。"}]

    first_index = next(i for i, line in enumerate(lines) if line.strip())
    deck_title = lines[first_index].strip()
    body = "\n".join(lines[first_index + 1 :]).strip()
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", body) if p.strip()]
    slides: list[dict[str, Any]] = []
    for idx, paragraph in enumerate(paragraphs or [body or deck_title], start=1):
        para_lines = paragraph.split("\n")
        bullet_items = []
        normal = []
        for line in para_lines:
            m = BULLET_RE.match(line)
            if m:
                bullet_items.append(m.group("text").strip())
            else:
                normal.append(line.strip())
        title_seed = next((x for x in normal if x), "") or (bullet_items[0] if bullet_items else f"Slide {idx}")
        title = title_seed[:34] + ("…" if len(title_seed) > 34 else "")
        slide: dict[str, Any] = {"title": title or f"Slide {idx}"}
        if bullet_items:
            slide["bullets"] = bullet_items[:max_bullets]
            if normal:
                slide["body"] = " ".join(normal)
        else:
            slide["body"] = " ".join(normal)
        slides.append(slide)
    return deck_title, slides


def load_source(path: Path, *, aspect: str, footer: str, preset: str, max_bullets: int) -> dict[str, Any]:
    if path.suffix.lower() == ".json":
        spec = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(spec, dict) or not isinstance(spec.get("slides"), list):
            raise ValueError("JSON input must contain a slides array")
        out = dict(spec)
        out.setdefault("canvas", {"preset": aspect})
        out.setdefault("meta", {})
        out["meta"] = {**out["meta"], "autoDeckPreset": preset, "source": str(path.resolve())}
        if footer:
            out["theme"] = {**(out.get("theme") or {}), "footerText": footer}
        return out

    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() in {".md", ".markdown"}:
        deck_title, slides = parse_markdown(
            text, source_dir=path.parent.resolve(), max_bullets=max_bullets
        )
    else:
        deck_title, slides = parse_plain_text(text, max_bullets=max_bullets)

    theme: dict[str, Any] = {}
    if footer:
        theme["footerText"] = footer

    return {
        "canvas": {"preset": aspect},
        "meta": {
            "title": deck_title,
            "autoDeckPreset": preset,
            "source": str(path.resolve()),
        },
        "theme": theme,
        "slides": slides,
    }


def run_pipeline(
    source: Path,
    output_pptx: Path,
    *,
    preset: str = "research-progress",
    aspect: str = "4:3",
    footer: str = "",
    max_bullets: int = 5,
    strict: bool = True,
    render: bool = False,
    spec_only: bool = False,
) -> dict[str, str]:
    source = source.resolve()
    if not source.exists():
        raise FileNotFoundError(source)

    output_pptx = output_pptx.resolve()
    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    stem = output_pptx.with_suffix("")
    spec_path = stem.with_suffix(".spec.json")
    ir_path = stem.with_suffix(".ir.json")
    report_path = stem.with_suffix(".layout-report.json")
    manifest_path = stem.with_suffix(".auto-deck.json")
    render_dir = output_pptx.parent / f"{output_pptx.stem}_rendered"

    spec = load_source(
        source,
        aspect=aspect,
        footer=footer,
        preset=preset,
        max_bullets=max_bullets,
    )
    spec_path.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")

    result = {
        "source": str(source),
        "spec": str(spec_path),
        "ir": str(ir_path),
        "report": str(report_path),
        "pptx": str(output_pptx),
        "renderDir": str(render_dir),
        "manifest": str(manifest_path),
    }

    if not spec_only:
        repo_root = Path(__file__).resolve().parent.parent
        engine = repo_root / "src" / "pixel_layout_engine.py"
        cmd = [
            sys.executable, str(engine), "build", str(spec_path),
            "--out-ir", str(ir_path),
            "--pptx", str(output_pptx),
            "--report", str(report_path),
        ]
        if strict:
            cmd.append("--strict")
        if render:
            cmd.extend(["--render-dir", str(render_dir)])

        subprocess.run(cmd, cwd=repo_root, check=True)

    manifest = {
        **result,
        "preset": preset,
        "aspect": aspect,
        "strict": strict,
        "render": render,
        "specOnly": spec_only,
        "createdAtUtc": datetime.now(timezone.utc).isoformat(),
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return result


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Markdown/JSON/TXT -> compact spec -> lint -> PPTX -> optional render"
    )
    ap.add_argument("source", type=Path)
    ap.add_argument("--output", type=Path, help="output .pptx path")
    ap.add_argument(
        "--preset",
        choices=("research-progress", "scientific", "story-driven"),
        default="research-progress",
    )
    ap.add_argument("--aspect", choices=("4:3", "16:9"), default="4:3")
    ap.add_argument("--footer", default="")
    ap.add_argument("--max-bullets", type=int, default=5)
    ap.add_argument("--render", action="store_true", help="render final PNGs with desktop PowerPoint")
    ap.add_argument("--no-strict", action="store_true")
    ap.add_argument("--spec-only", action="store_true")
    args = ap.parse_args(argv)

    if args.max_bullets < 1:
        ap.error("--max-bullets must be >= 1")

    output = args.output or (Path("output") / f"{args.source.stem}.pptx")
    try:
        result = run_pipeline(
            args.source,
            output,
            preset=args.preset,
            aspect=args.aspect,
            footer=args.footer,
            max_bullets=args.max_bullets,
            strict=not args.no_strict,
            render=args.render,
            spec_only=args.spec_only,
        )
    except (OSError, ValueError, subprocess.CalledProcessError) as exc:
        print(f"[auto-deck] ERROR: {exc}", file=sys.stderr)
        return 1

    print("[auto-deck] complete")
    for key in ("spec", "ir", "report", "pptx", "renderDir", "manifest"):
        print(f"[auto-deck] {key}: {result[key]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
