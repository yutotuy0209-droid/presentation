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


if __name__ == "__main__":
    unittest.main()
