---
name: auto-deck-pipeline
description: Turn a concise Markdown, JSON, or text brief into an editable PowerPoint through the repository's deterministic layout, lint, render, and review pipeline.
---

# Auto Deck Pipeline

## Use when

Use this skill when the user wants a PowerPoint produced with minimal manual layout work, especially for recurring technical reports, research progress, scientific explanations, and comparison decks.

## Core rule

The authoring surface should be small. Prefer one concise Markdown file over manually writing PowerPoint geometry.

Pipeline:

```text
brief.md / spec.json
        ↓
auto_deck.py
        ↓
compact slide spec
        ↓
pixel_layout_engine.py
        ↓
common IR + layout report
        ↓
lint
        ↓
editable PPTX
        ↓
optional desktop PowerPoint PNG render
        ↓
human correction → experience / preset update
```

## Recommended input

Markdown is the default.

```markdown
# Weekly research report

## What changed
- Added real-time processing
- Reduced manual operation
- Preserved RAW images

## Result
Short interpretation.

![measured result](images/result.png)

## Next action
- Validate with G-code
- Compare against reference sensor
```

Rules:

- `#` is the deck title.
- `##` starts one slide topic.
- Bullets are chunked automatically.
- Markdown image paths are resolved relative to the source file.
- A slide that contains text and an image uses an image-right layout by default.
- Use JSON only when exact layout control is required.

## Commands

Minimal:

```powershell
python src/auto_deck.py brief.md --output output/report.pptx
```

With desktop PowerPoint rendering:

```powershell
python src/auto_deck.py brief.md --output output/report.pptx --render
```

GUI:

```powershell
python src/auto_deck_ui.py
```

## Quality gates

Do not call the deck complete only because a PPTX file exists.

Required path:

1. compact spec generated
2. layout report has no blocking errors
3. common IR lint passes
4. PPTX writer succeeds
5. when PowerPoint is available, render PNGs and inspect visually
6. repeated human corrections are moved into presets/skills rather than patched ad hoc every week

## Design policy

- One main message per slide.
- Evidence before decoration.
- Avoid screenshots of whole slides when editable native elements are possible.
- Preserve image aspect ratio.
- Never hide unsupported elements behind silent approximations.
- Prefer deterministic layout rules over asking an LLM to invent coordinates repeatedly.

## Motion

PowerPoint slide animation is not the primary quality mechanism. Motion may be used for exported video or demo material, but the static slide must remain understandable without animation.
