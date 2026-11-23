#!/usr/bin/env python3
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from design_lint import lint_grid_spec  # noqa: E402


class DesignLintTests(unittest.TestCase):
    def test_dense_slide_warns(self):
        spec = {
            "slides": [{
                "id": "dense",
                "title": "A" * 50,
                "grid": [["A", "B"], ["C", "D"]],
                "blocks": {
                    "A": {"type": "card", "body": "detail " * 40},
                    "B": {"type": "card"},
                    "C": {"type": "card"},
                    "D": {"type": "card"},
                },
                "connectors": [
                    {"from": "A", "to": "B"},
                    {"from": "A", "to": "C"},
                    {"from": "A", "to": "D"},
                    {"from": "B", "to": "C"},
                    {"from": "C", "to": "D"},
                ],
            }]
        }
        rules = {x["rule"] for x in lint_grid_spec(spec)}
        self.assertIn("long-title", rules)
        self.assertIn("dense-card", rules)
        self.assertIn("connector-density", rules)
        self.assertIn("no-semantic-role-hierarchy", rules)

    def test_simple_semantic_slide_is_clean(self):
        spec = {
            "slides": [{
                "id": "clean",
                "title": "Input → Process → Output",
                "grid": [["A", "B", "C"]],
                "blocks": {
                    "A": {"type": "card", "role": "input", "title": "Input"},
                    "B": {"type": "card", "role": "process", "title": "Process"},
                    "C": {"type": "metric", "role": "output", "value": "10 µm"},
                },
                "connectors": [
                    {"from": "A", "to": "B"},
                    {"from": "B", "to": "C"},
                ],
            }]
        }
        self.assertEqual(lint_grid_spec(spec), [])

    def test_baseline_errors_are_reported(self):
        spec = {"slides": [{
            "title": "｜ 判断事項",
            "footer": "注意：" + "条件" * 70,
            "blocks": {
                "A": {"title": "01：リスク", "body": "（別途ご相談）"},
                "B": {"title": "リスク：対策", "body": "価値を最大化"},
                "C": {"title": "リスク：比較"},
            },
            "grid": [["A", "B", "C"]],
        }]}
        issues = lint_grid_spec(spec)
        rules = {x["rule"] for x in issues}
        self.assertIn("decorative-leading-bar", rules)
        self.assertIn("unresolved-placeholder", rules)
        self.assertIn("overlong-footer", rules)
        self.assertIn("generic-buzzword", rules)
        self.assertIn("unnecessary-zero-padding", rules)
        self.assertIn("repeated-column-label", rules)
        self.assertTrue(any(x["severity"] == "error" for x in issues))

    def test_decision_without_comparison_warns(self):
        spec = {"slides": [{
            "title": "ご判断事項",
            "blocks": {"A": {"type": "card", "body": "議論を進める"}},
        }]}
        rules = {x["rule"] for x in lint_grid_spec(spec)}
        self.assertIn("decision-without-options", rules)


if __name__ == "__main__":
    unittest.main()
