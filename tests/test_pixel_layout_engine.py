#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from pixel_layout_engine import (  # noqa: E402
    DeckCompiler,
    LayoutError,
    Rect,
    px_to_emu,
    required_text_height_px,
)


class PixelLayoutEngineTests(unittest.TestCase):
    def test_px_to_emu_is_exact_at_96dpi(self):
        self.assertEqual(px_to_emu(96, 96), 914400)
        self.assertEqual(px_to_emu(1, 96), 9525)

    def test_rect_collision(self):
        self.assertTrue(Rect(0, 0, 100, 100).intersects(Rect(99, 99, 10, 10)))
        self.assertFalse(Rect(0, 0, 100, 100).intersects(Rect(100, 0, 10, 10)))

    def test_text_height_grows_with_text(self):
        short = required_text_height_px("短文", 400, 20, 1.34, 96)
        long = required_text_height_px("長い文章" * 80, 400, 20, 1.34, 96)
        self.assertGreater(long, short)

    def test_compiler_outputs_4x3_and_no_errors(self):
        spec = {
            "canvas": {"preset": "4:3", "dpi": 96},
            "slides": [
                {
                    "id": "s1",
                    "layout": "title-body",
                    "title": "タイトル",
                    "blocks": [
                        {"type": "text", "text": "本文です。"},
                        {"type": "text", "text": "二つ目の本文です。"},
                    ],
                }
            ],
        }
        with tempfile.TemporaryDirectory() as td:
            spec_path = Path(td) / "spec.json"
            spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            deck, report = DeckCompiler(spec, spec_path).compile()

        self.assertEqual(deck["slideSize"]["width"], px_to_emu(960, 96))
        self.assertEqual(deck["slideSize"]["height"], px_to_emu(720, 96))
        self.assertTrue(report["ok"])
        self.assertEqual(report["errors"], [])


    def test_fixed_position_is_preserved_exactly(self):
        spec = {
            "canvas": {"preset": "4:3", "dpi": 96},
            "slides": [
                {
                    "id": "fixed",
                    "title": "座標指定",
                    "blocks": [
                        {
                            "type": "text",
                            "text": "固定配置",
                            "position": {
                                "x": 120,
                                "y": 180,
                                "w": 320,
                                "h": 90,
                                "mode": "fixed",
                            },
                        }
                    ],
                }
            ],
        }
        with tempfile.TemporaryDirectory() as td:
            spec_path = Path(td) / "spec.json"
            spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            deck, report = DeckCompiler(spec, spec_path).compile()

        el = next(e for e in deck["slides"][0]["elements"] if e["id"] == "fixed-b0")
        self.assertEqual(
            el["source"]["pixel"],
            {"x": 120, "y": 180, "width": 320, "height": 90},
        )
        self.assertTrue(report["ok"])

    def test_fixed_collision_raises_instead_of_moving(self):
        spec = {
            "canvas": {"preset": "4:3", "dpi": 96},
            "slides": [
                {
                    "id": "collision",
                    "title": "固定衝突",
                    "blocks": [
                        {
                            "type": "text",
                            "text": "A",
                            "position": {"x": 120, "y": 180, "w": 300, "h": 80, "mode": "fixed"},
                        },
                        {
                            "type": "text",
                            "text": "B",
                            "position": {"x": 200, "y": 200, "w": 300, "h": 80, "mode": "fixed"},
                        },
                    ],
                }
            ],
        }
        with tempfile.TemporaryDirectory() as td:
            spec_path = Path(td) / "spec.json"
            spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            with self.assertRaises(LayoutError):
                DeckCompiler(spec, spec_path).compile()

    def test_preferred_position_nudges_only_when_needed(self):
        spec = {
            "canvas": {"preset": "4:3", "dpi": 96},
            "slides": [
                {
                    "id": "prefer",
                    "title": "優先座標",
                    "blocks": [
                        {
                            "type": "text",
                            "text": "A",
                            "position": {"x": 120, "y": 180, "w": 300, "h": 80, "mode": "fixed"},
                        },
                        {
                            "type": "text",
                            "text": "B",
                            "position": {
                                "x": 120,
                                "y": 180,
                                "w": 300,
                                "h": 80,
                                "mode": "prefer",
                                "axis": "y",
                                "maxNudgePx": 120,
                            },
                        },
                    ],
                }
            ],
        }
        with tempfile.TemporaryDirectory() as td:
            spec_path = Path(td) / "spec.json"
            spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            deck, report = DeckCompiler(spec, spec_path).compile()

        b = next(e for e in deck["slides"][0]["elements"] if e["id"] == "prefer-b1")
        self.assertEqual(b["source"]["pixel"]["x"], 120)
        self.assertGreater(b["source"]["pixel"]["y"], 180)
        self.assertLessEqual(b["source"]["pixel"]["y"] - 180, 120)
        self.assertTrue(report["ok"])

    def test_long_text_auto_paginates(self):
        spec = {
            "canvas": {"preset": "4:3", "dpi": 96},
            "theme": {"bodySizePt": 22, "minBodySizePt": 16},
            "slides": [
                {
                    "id": "long",
                    "layout": "title-body",
                    "title": "長文",
                    "blocks": [
                        {"type": "text", "text": "これは長文です。" * 1200}
                    ],
                }
            ],
        }
        with tempfile.TemporaryDirectory() as td:
            spec_path = Path(td) / "spec.json"
            spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            deck, report = DeckCompiler(spec, spec_path).compile()

        self.assertGreater(len(deck["slides"]), 1)
        self.assertTrue(report["ok"])


if __name__ == "__main__":
    unittest.main()
