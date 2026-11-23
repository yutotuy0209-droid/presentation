#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deterministic design checks for the high-level Grid Deck spec."""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any



def _baseline_quality_checks(slide: dict[str, Any], sid: str) -> list[dict[str, Any]]:
    """Catch objective AI-generated slide defects before rendering.

    Semantic coherence and actual text wrapping still need rendered review.
    """
    issues: list[dict[str, Any]] = []
    blocks = slide.get("blocks") or {}
    entries: list[tuple[str, str, str]] = [
        ("slide", "title", str(slide.get("title") or "")),
        ("slide", "footer", str(slide.get("footer") or "")),
    ]
    repeated_labels: dict[str, list[str]] = {}
    for bid, block in blocks.items():
        if not isinstance(block, dict):
            continue
        for field in ("title", "subtitle", "body", "text", "description", "note", "footer"):
            if block.get(field):
                entries.append((str(bid), field, str(block[field])))
        for field in ("title", "body", "text"):
            value = str(block.get(field) or "")
            match = re.match(r"^\\s*(リスク|ご判断いただきたい点|判断事項)\\s*[：:]", value)
            if match:
                repeated_labels.setdefault(match.group(1), []).append(str(bid))

    for bid, field, value in entries:
        stripped = value.strip()
        location = {"slide": sid}
        if bid != "slide":
            location["block"] = bid
        if re.match(r"^[|｜]\\s+", stripped):
            issues.append({
                **location, "severity": "error", "rule": "decorative-leading-bar",
                "message": "Remove meaningless leading vertical bars in headings.",
            })
        if field == "title" and re.match(r"^(?:STEP\\s*)?0[1-9](?=\\s|[.．、:：])", stripped, re.I):
            issues.append({
                **location, "severity": "warning", "rule": "unnecessary-zero-padding",
                "message": "Use 1, 2, 3 instead of 01, 02, 03 unless the number is meaningful.",
            })
        if "（別途ご相談）" in value or "（個別の状況による）" in value:
            issues.append({
                **location, "severity": "error", "rule": "unresolved-placeholder",
                "message": "Replace vague hand-off text with responsible owner, next action and criteria.",
            })
        if any(x in value for x in ("代替ではなく伴走者", "価値を最大化", "シームレスに展開")):
            issues.append({
                **location, "severity": "warning", "rule": "generic-buzzword",
                "message": "Replace generic slogans with actions, observable facts or decisions.",
            })
        if field == "footer" and len(stripped) > 120:
            issues.append({
                **location, "severity": "error", "rule": "overlong-footer",
                "value": len(stripped),
                "message": "Move long legal/source notes to a readable dedicated area or appendix.",
            })
        if re.search(r"[はがのをにへとでも]\\n", value):
            issues.append({
                **location, "severity": "warning", "rule": "suspicious-manual-linebreak",
                "message": "Check whether this forced break separates a Japanese particle from its phrase.",
            })

    for label, bids in repeated_labels.items():
        if len(bids) > 1:
            issues.append({
                "slide": sid, "severity": "warning", "rule": "repeated-column-label",
                "value": {"label": label, "blocks": bids},
                "message": "Move the shared label to a common axis or header.",
            })

    title = str(slide.get("title") or "")
    if any(token in title for token in ("ご判断事項", "判断事項", "意思決定")):
        all_text = " ".join(value for _, _, value in entries)
        if not (("A案" in all_text and "B案" in all_text)
                and any(token in all_text for token in ("推奨", "採用案", "選択理由"))):
            issues.append({
                "slide": sid, "severity": "warning", "rule": "decision-without-options",
                "message": "A decision slide needs concrete alternatives and a reasoned recommendation.",
            })

    grid = slide.get("grid") or []
    for row_number, row in enumerate(grid, start=1):
        if not isinstance(row, list):
            continue
        ids = list(dict.fromkeys(str(x) for x in row))
        heights = [
            blocks[bid].get("height")
            for bid in ids
            if isinstance(blocks.get(bid), dict) and blocks[bid].get("height") is not None
        ]
        if len(heights) >= 2 and len(heights) == len(ids) and len({str(x) for x in heights}) > 1:
            issues.append({
                "slide": sid, "severity": "warning", "rule": "uneven-specified-card-heights",
                "value": row_number,
                "message": "Cards in the same row should have consistent visible height.",
            })

    return issues


def lint_grid_spec(spec: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []

    for index, slide in enumerate(spec.get("slides") or [], start=1):
        sid = str(slide.get("id") or f"slide-{index}")
        title = str(slide.get("title") or "")
        if len(title) > 44:
            issues.append({
                "slide": sid,
                "severity": "warning",
                "rule": "long-title",
                "value": len(title),
                "message": "Slide title is long; shorten the claim before reducing font size.",
            })

        blocks = slide.get("blocks") or {}
        block_count = len(blocks)
        if block_count > 7:
            issues.append({
                "slide": sid,
                "severity": "warning",
                "rule": "too-many-primary-blocks",
                "value": block_count,
                "message": "More than 7 primary blocks weakens glanceable hierarchy.",
            })

        connectors = slide.get("connectors") or []
        if block_count and len(connectors) > max(3, block_count):
            issues.append({
                "slide": sid,
                "severity": "warning",
                "rule": "connector-density",
                "value": len(connectors),
                "message": "Use arrows only for sequence, dependency, data flow, causality, or command direction.",
            })

        for bid, block in blocks.items():
            btype = str(block.get("type") or "card").lower()
            if btype in {"card", "text", "callout"}:
                body = str(block.get("body") or block.get("text") or "")
                lines = body.count("\n") + (1 if body else 0)
                if len(body) > 150 or lines > 4:
                    issues.append({
                        "slide": sid,
                        "block": str(bid),
                        "severity": "warning",
                        "rule": "dense-card",
                        "value": {"chars": len(body), "lines": lines},
                        "message": "Keep card detail subordinate; split or move detail instead of shrinking.",
                    })
            if btype == "metric" and not str(block.get("value") or "").strip():
                issues.append({
                    "slide": sid,
                    "block": str(bid),
                    "severity": "warning",
                    "rule": "empty-metric",
                    "message": "Metric blocks should visually prioritize a concrete value.",
                })

        issues.extend(_baseline_quality_checks(slide, sid))

        roles = [
            str(block.get("role") or "neutral")
            for block in blocks.values()
            if isinstance(block, dict)
        ]
        if block_count >= 4 and roles and all(role == "neutral" for role in roles):
            issues.append({
                "slide": sid,
                "severity": "warning",
                "rule": "no-semantic-role-hierarchy",
                "message": "Consider input/process/output/human/warning roles when they carry meaning.",
            })

    return issues


def main() -> int:
    ap = argparse.ArgumentParser(description="Lint high-level Grid Deck design")
    ap.add_argument("spec", type=Path)
    ap.add_argument("--out", type=Path)
    args = ap.parse_args()

    spec = json.loads(args.spec.read_text(encoding="utf-8"))
    issues = lint_grid_spec(spec)
    payload = json.dumps(issues, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(payload, encoding="utf-8")
        print(f"[design-lint] {args.out}")
    else:
        print(payload)
    return 1 if any(x.get("severity") == "error" for x in issues) else 0


if __name__ == "__main__":
    raise SystemExit(main())
