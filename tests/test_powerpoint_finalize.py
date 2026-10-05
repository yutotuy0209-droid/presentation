#!/usr/bin/env python3
import unittest

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from powerpoint_finalize import build_finalize_plan, FinalizeError  # noqa: E402


class PowerPointFinalizePlanTests(unittest.TestCase):
    def test_animation_transition_and_hyperlink_plan(self):
        deck = {
            "slides": [{
                "id": "s1",
                "transition": {"type": "fade", "duration": 0.5},
                "elements": [
                    {
                        "id": "title",
                        "type": "sp",
                        "animation": {
                            "entrance": "fade",
                            "order": 2,
                            "trigger": "click",
                        },
                    },
                    {
                        "id": "link",
                        "type": "sp",
                        "interaction": {
                            "hyperlink": "https://example.com"
                        },
                    },
                ],
            }]
        }
        report = {
            "shapeMap": {
                "s1": {
                    "title": 1,
                    "link": 2,
                }
            }
        }
        plan = build_finalize_plan(deck, report)
        self.assertEqual(plan["slides"][0]["transition"]["type"], "fade")
        self.assertEqual(plan["slides"][0]["animations"][0]["shapeIndex"], 1)
        self.assertEqual(plan["slides"][0]["animations"][0]["order"], 2)
        self.assertEqual(plan["slides"][0]["hyperlinks"][0]["shapeIndex"], 2)
        self.assertEqual(plan["slides"][0]["hyperlinks"][0]["address"], "https://example.com")

    def test_template_path_flows_to_finalize_plan(self):
        deck = {
            "meta": {
                "template": {
                    "path": "C:/templates/lab.potx",
                    "applyOnWindows": True,
                }
            },
            "slides": [{"id": "s1", "elements": []}],
        }
        plan = build_finalize_plan(deck, {"shapeMap": {"s1": {}}})
        self.assertEqual(plan["template"]["path"], "C:/templates/lab.potx")
        self.assertTrue(plan["template"]["applyOnWindows"])

    def test_missing_shape_mapping_is_reported_not_guessed(self):
        deck = {
            "slides": [{
                "id": "s1",
                "elements": [{
                    "id": "a",
                    "type": "sp",
                    "animation": {"entrance": "appear"},
                }],
            }]
        }
        plan = build_finalize_plan(deck, {"shapeMap": {"s1": {}}})
        self.assertEqual(plan["slides"][0]["animations"], [])
        self.assertEqual(len(plan["warnings"]), 1)
        self.assertIn("shape mapping", plan["warnings"][0]["reason"])

    def test_unsupported_animation_is_rejected(self):
        deck = {
            "slides": [{
                "id": "s1",
                "elements": [{
                    "id": "a",
                    "type": "sp",
                    "animation": {"entrance": "magic-spin"},
                }],
            }]
        }
        with self.assertRaises(FinalizeError):
            build_finalize_plan(deck, {"shapeMap": {"s1": {"a": 1}}})


if __name__ == "__main__":
    unittest.main()
