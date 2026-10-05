# Round-trip First PowerPoint Architecture

## Goal

The presentation system should not claim broad PowerPoint support only because it can create many object types.

The stronger target is:

> A PowerPoint feature is considered supported only when the system can detect it in an existing PPTX, preserve or represent it during reverse conversion, reproduce or retain it during forward conversion, and verify the result by round-trip comparison.

This makes reverse conversion a core capability, not an auxiliary import feature.

---

## Why reverse conversion matters

### 1. Generation coverage becomes measurable

If a human can create a feature in PowerPoint, the reverse pipeline should detect it.

The system can then classify each feature as:

- semantic-roundtrip: represented in IR and regenerated intentionally
- preserved-roundtrip: not yet editable in IR, but retained losslessly in OOXML/package
- partial: some properties are represented but fidelity is incomplete
- unsupported: detected but neither reproduced nor safely preserved

The target is to reduce unsupported to zero and steadily move preserved-roundtrip toward semantic-roundtrip.

### 2. Human-created decks become design data

A human-edited PPTX contains more useful design evidence than a prompt describing a preferred style.

Reverse conversion can extract measurable properties such as:

- slide size and aspect ratio
- title/body font sizes
- font families and weights
- margins and whitespace
- object alignment
- repeated grid structures
- image-to-text ratios
- shape geometry
- fill/stroke colors
- connector use
- hierarchy between dominant and secondary objects
- relative spacing
- repeated slide families
- master/layout/theme usage

These observations should be stored as design evidence, not copied blindly into permanent rules.

### 3. Human corrections become training signals for the generator

Preferred loop:

generated deck -> human edit -> reverse both versions -> structural/style diff -> repeated corrections -> candidate preset/layout rule -> human confirmation -> promotion

One correction is evidence, not a universal rule.

---

## Canonical representations

### A. Source Bundle — preservation truth

`pptx_bundle.py` preserves the complete PPTX package including slide XML, relationships, theme, slide masters, slide layouts, media, charts, SmartArt/diagram parts, notes, embeddings, animation/timing XML, and unknown extension parts.

This is the loss-aware fallback.

### B. Slide IR — editable semantic representation

`pptx_to_ir.py` should progressively represent PowerPoint objects and properties in a normalized form.

A feature is not considered semantically supported merely because its raw XML is retained.

### C. Grid DSL — authoring representation

`grid_deck.py` is intentionally simpler than raw PowerPoint. It represents useful authoring intent such as 2D layout, semantic role, connector relationship, media block, animation intent, and design preset.

Reverse conversion does not silently invent Grid semantics that were not stored in the original PPTX.

### D. Design Evidence — learned observations

Human-created or human-edited decks should produce compact design statistics and reusable candidates.

Raw decks remain the evidence source; learned presets are derived artifacts.

---

## Coverage definition

| Feature | Detect | Preserve | IR | Generate | Round-trip test |
|---|---:|---:|---:|---:|---:|
| Text | yes | yes | yes | yes | required |
| Shape geometry | partial | yes | partial | yes | required |
| Images | yes | yes | yes | yes | required |
| Tables | yes | yes | yes | yes | required |
| GIF | partial | yes | partial | yes | required |
| Video/audio | partial | yes | partial | yes | required |
| Connectors | partial | yes | partial | yes | required |
| Charts | raw only | yes | planned | planned | planned |
| Equations | raw only | yes | planned | planned | planned |
| Groups | detected | yes | planned | planned | planned |
| Theme/master/layout | raw | yes | planned semantic model | template apply | required |
| Animation/timing | raw | yes | partial intent | COM partial | required |
| Transition | raw | yes | partial | COM partial | required |
| SmartArt | raw | yes | planned | preserve first | required |
| OLE/embedded Excel | raw | yes | planned | preserve first | required |

The table should eventually be generated from tests rather than maintained manually.

---

## Design learning policy

Do not overwrite presets from one deck. Extract observable features and accumulate evidence.

Suggested per-slide evidence fields:

- slide_family
- title_pt
- body_pt_median
- primary_gap_px
- content_density
- dominant_visual_ratio
- palette
- alignment
- connector_count
- image_count

Suggested aggregate signals:

- selected_count
- build_count
- human_edit_count
- recurrence across decks
- recency
- context match
- visual QA success
- human review time

Only repeated patterns should become preset candidates.

---

## Acceptance principle

The long-term acceptance test is not:

> Can Python create this feature?

It is:

> Can a real human-created PPTX containing this feature be reversed, understood or preserved, regenerated, and compared without silent loss?

That is the definition of meaningful PowerPoint coverage.
