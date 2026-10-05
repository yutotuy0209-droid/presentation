#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deterministic design checks for the high-level Grid Deck spec."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


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
