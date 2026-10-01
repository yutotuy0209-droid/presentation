# Native deck validation contract

Client projects own their text, photographs and measurement claims. This repository owns reusable layout and validation methods. Do not copy client decks or evidence into the shared tool repository.

1. Extract teacher geometry using PPTX-to-IR; keep teacher content separate from layout choices. Choose fixed, prefer or flow per element rather than replacing the entire deck with flattened screenshots.
2. Compile a saved layout specification with Pixel Layout Engine. Preserve human-edited input on regeneration; refreshing seeded content must be explicit.
3. Run IR lint, canvas bounds and collision checks. Treat picture fit and signed connectors according to their geometry rather than assuming every bounding-box intersection is text overlap.
4. Export native text, shapes and charts where supported. Hiding inherited master graphics requires the slide's `showMasterSp="0"` XML attribute; merely assigning an unrecognized Python property is insufficient.
5. Render using the actual presentation application. Compare text bounds with shape bounds, and inspect every page in contact sheets. Native editable equation text is not equivalent to structured Office Math; describe the export format accurately.
6. Save client-side provenance with source specification hash, deck hash, tool revision and validation scope. If an existing deck hash differs from the last generated hash, preserve the deck and refuse automatic replacement.

Keep temporary renders and Python bytecode outside tracked changes. Tests, lint and renders establish software/layout behavior, not the validity of a client's research results. A successful render cannot establish camera timing, machine safety or measurement accuracy.
