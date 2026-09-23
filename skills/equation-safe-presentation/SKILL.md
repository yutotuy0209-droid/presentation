# Equation-Safe Presentation Skill

## Use when

A presentation contains LaTeX, mathematical expressions, subscripts, superscripts, Greek letters, matrices, fractions, or equations that may break during conversion.

## Goal

Preserve mathematical meaning first, then visual fidelity, then editability.

## Decision order

1. Keep simple expressions as native text when Unicode/subscript formatting is reliable.
2. Use native PowerPoint equation support when the generation path supports it correctly.
3. For complex equations, render from a trusted math source to SVG and embed the SVG.
4. Use PNG only as a fallback when SVG rendering is not reliable.
5. Never silently replace an equation with plain text that changes its meaning.

## Checks

- fractions are not flattened incorrectly
- subscripts/superscripts belong to the correct symbol
- Greek letters are not substituted with Latin lookalikes
- minus signs and negative values remain visible
- multiplication signs and decimal points remain distinct
- line breaks do not split an equation incorrectly
- equation size remains readable on the final slide

## Reuse principle

Store equation rendering as a helper function or reusable conversion step. Do not fix the same conversion failure manually in every deck.
