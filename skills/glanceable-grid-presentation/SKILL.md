# Glanceable Grid Presentation Skill

## Use when

Creating or revising a PowerPoint where the audience should understand the structure before reading details.

Use this together with:
- `skills/presentation-structure/SKILL.md`
- `skills/scientific-presentation/SKILL.md`
- `docs/DESIGN_PREFERENCE_LOOP.md`
- `src/grid_deck.py`

## Core target

> 一目で構造、数秒で関係、読めば具体。

Design for three reading depths.

### 3 seconds
The audience should understand:
- what the slide is about,
- the dominant 1–7 blocks,
- the reading direction,
- the main result or question.

### 10 seconds
The audience should understand:
- input / process / output,
- human action vs system action,
- comparison or causal relationships,
- why arrows exist.

### 30 seconds
The audience may read:
- units,
- conditions,
- filenames,
- exact values,
- short explanatory text.

Do not move 30-second information into the 3-second visual hierarchy.

## Layout choice

Choose the smallest Grid that communicates the relationship.

Preferred forms:

```text
[A]

[A][B]

[A]
[B]

[A][B][C]

[A][B]
[C][D]

[A][B]
[A][C]

[A][A]
[B][C]
```

Use repeated IDs only for intentional rectangular spans.

Avoid building a 3×3 dashboard merely because space exists.

## Block design

A normal card should contain:
1. a short title,
2. at most 1–4 lines of detail,
3. no decorative paragraph.

Metric cards should prioritize the number over the label.

Media blocks should be visually dominant when they are evidence:
- measured image,
- SVG relationship diagram,
- GIF for short motion,
- video only when motion itself is evidence.

## Semantic roles

Do not choose arbitrary colors per slide.

Use:
- `input`
- `process`
- `output`
- `human`
- `warning`
- `neutral`

Color means role, not decoration.

## Human vs system

When a workflow contains both manual and automatic actions, make the distinction explicit.

Typical structure:

```text
Human action
    ↓
System capture
    ↓
Automatic calculation
    ↓
Human judgement
```

Use `human` role for manual focus adjustment, confirmation, machine-side checks, or approval.

## Connector rule

Add an arrow only when it communicates one of:
- sequence,
- dependency,
- data flow,
- causality,
- command direction.

Do not connect every neighboring box.

Prefer left-to-right or top-to-bottom flow. Do not mix reading directions on the same slide unless the meaning requires it.

## Typography

Default to the active glanceable preset.

Guideline:
- slide title: around 28–34 pt,
- block title: around 18–22 pt,
- body: around 15–20 pt,
- avoid body text below 14 pt.

If content does not fit, simplify, split, or move detail to another slide before shrinking aggressively.

## Design selection

Do not encode every human preference into permanent rules.

For presentation work:
1. read the short design preference summary when available,
2. prefer historically selected designs,
3. keep one exploration candidate,
4. promote a design change only after repeated human selection/correction.

Available baseline presets:
- `glanceable-scientific`
- `glanceable-soft-graphite`
- `glanceable-warm`

## Motion

Use motion only when it adds explanation.

Priority:
1. static SVG / image,
2. GIF for short repeated motion,
3. video for real temporal evidence,
4. native PowerPoint animation only when staged explanation is useful.

Do not animate decorative elements.

Native animation requests should be stored in the DSL and applied by `src/powerpoint_finalize.py` on Windows.

## Quality gate

Before accepting a slide, check:
- Can the structure be understood from thumbnail scale?
- Is one element visually dominant?
- Are there no unnecessary connectors?
- Are units and conditions visible where needed?
- Are human actions visually distinct from automated actions?
- Is measured evidence stronger than explanatory prose?
- Did the system avoid inventing missing values?
