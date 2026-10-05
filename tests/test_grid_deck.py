#!/usr/bin/env python3
import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from grid_deck import compile_grid_spec, load_design_preset  # noqa: E402
from pixel_layout_engine import DeckCompiler, image_size  # noqa: E402


class GridDeckTests(unittest.TestCase):
    def _compile(self, spec):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "grid.json"
            path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
            pixel_spec = compile_grid_spec(spec, path)
            deck, report = DeckCompiler(pixel_spec, path).compile()
            return pixel_spec, deck, report

    def test_single_cell_becomes_one_region(self):
        spec = {
            "canvas": {"preset": "16:9"},
            "slides": [{
                "id": "one",
                "title": "One",
                "grid": [["A"]],
                "blocks": {"A": {"type": "card", "title": "A", "body": "detail"}},
            }],
        }
        pixel_spec, deck, report = self._compile(spec)
        regions = pixel_spec["slides"][0]["regions"]
        self.assertIn("grid_A", regions)
        self.assertTrue(report["ok"])
        self.assertEqual(len(deck["slides"]), 1)

    def test_horizontal_and_vertical_layouts_are_distinct(self):
        horizontal = {
            "slides": [{
                "id": "h",
                "grid": [["A", "B"]],
                "blocks": {
                    "A": {"type": "card", "title": "A"},
                    "B": {"type": "card", "title": "B"},
                },
            }]
        }
        vertical = {
            "slides": [{
                "id": "v",
                "grid": [["A"], ["B"]],
                "blocks": {
                    "A": {"type": "card", "title": "A"},
                    "B": {"type": "card", "title": "B"},
                },
            }]
        }
        hp, _, _ = self._compile(horizontal)
        vp, _, _ = self._compile(vertical)
        ha, hb = hp["slides"][0]["regions"]["grid_A"], hp["slides"][0]["regions"]["grid_B"]
        va, vb = vp["slides"][0]["regions"]["grid_A"], vp["slides"][0]["regions"]["grid_B"]
        self.assertLess(ha["x"], hb["x"])
        self.assertEqual(ha["y"], hb["y"])
        self.assertEqual(va["x"], vb["x"])
        self.assertLess(va["y"], vb["y"])

    def test_repeated_cell_id_creates_rectangular_span(self):
        spec = {
            "slides": [{
                "id": "span",
                "grid": [["A", "B"], ["A", "C"]],
                "blocks": {
                    "A": {"type": "image", "path": "missing.png"},
                    "B": {"type": "card", "title": "B"},
                    "C": {"type": "card", "title": "C"},
                },
            }]
        }
        pixel_spec, _, _ = self._compile(spec)
        regions = pixel_spec["slides"][0]["regions"]
        self.assertGreater(regions["grid_A"]["height"], regions["grid_B"]["height"])

    def test_connector_compiles_to_arrow_line(self):
        spec = {
            "slides": [{
                "id": "flow",
                "grid": [["A", "B", "C"]],
                "blocks": {
                    "A": {"type": "card", "title": "Input"},
                    "B": {"type": "card", "title": "Process"},
                    "C": {"type": "card", "title": "Output"},
                },
                "connectors": [
                    {"from": "A", "to": "B"},
                    {"from": "B", "to": "C"},
                ],
            }]
        }
        pixel_spec, _, _ = self._compile(spec)
        arrows = [
            b for b in pixel_spec["slides"][0]["blocks"]
            if b.get("type") == "line" and b.get("style", {}).get("endArrowType") == "triangle"
        ]
        self.assertEqual(len(arrows), 2)

    def test_media_types_survive_high_level_compile(self):
        spec = {
            "slides": [{
                "id": "media",
                "grid": [["G", "S", "V"]],
                "blocks": {
                    "G": {"type": "gif", "path": "motion.gif"},
                    "S": {"type": "svg", "path": "diagram.svg"},
                    "V": {"type": "video", "path": "clip.mp4"},
                },
            }]
        }
        pixel_spec = compile_grid_spec(spec, Path("/tmp/grid.json"))
        blocks = pixel_spec["slides"][0]["blocks"]
        by_slot = {b.get("slot"): b for b in blocks if b.get("slot")}
        self.assertEqual(by_slot["grid_G"]["type"], "image")
        self.assertEqual(by_slot["grid_G"]["mediaType"], "gif")
        self.assertEqual(by_slot["grid_S"]["type"], "image")
        self.assertEqual(by_slot["grid_S"]["mediaType"], "svg")
        self.assertEqual(by_slot["grid_V"]["type"], "video")

    def test_design_preset_has_glanceable_hierarchy(self):
        preset = load_design_preset("glanceable-scientific")
        theme = preset["theme"]
        self.assertGreaterEqual(theme["titleSizePt"], 28)
        self.assertGreaterEqual(theme["bodySizePt"], 18)
        self.assertGreater(theme["blockGapPx"], 0)
        self.assertIn("roles", preset)
        self.assertIn("process", preset["roles"])

    def test_transition_hyperlink_and_audio_flow_to_pixel_spec(self):
        spec = {
            "slides": [{
                "id": "native",
                "title": "Native",
                "transition": {"type": "fade", "duration": 0.5},
                "grid": [["A", "B"]],
                "blocks": {
                    "A": {
                        "type": "card",
                        "title": "Link",
                        "hyperlink": "https://example.com",
                    },
                    "B": {
                        "type": "audio",
                        "path": "voice.mp3",
                    },
                },
            }]
        }
        pixel_spec, deck, _ = self._compile(spec)
        slide = pixel_spec["slides"][0]
        self.assertEqual(slide["transition"]["type"], "fade")
        link_blocks = [b for b in slide["blocks"] if b.get("hyperlink")]
        self.assertTrue(link_blocks)
        audio = next(b for b in slide["blocks"] if b.get("type") == "audio")
        self.assertEqual(audio["mediaType"], "audio")
        self.assertEqual(deck["slides"][0]["transition"]["type"], "fade")
        interactions = [
            e.get("interaction") for e in deck["slides"][0]["elements"]
            if e.get("interaction")
        ]
        self.assertTrue(any(x.get("hyperlink") == "https://example.com" for x in interactions))
        media = next(e for e in deck["slides"][0]["elements"] if e.get("type") == "media")
        self.assertEqual(media["source"]["mediaType"], "audio")

    def test_gif_dimensions_are_read(self):
        with tempfile.TemporaryDirectory() as td:
            p = Path(td) / "tiny.gif"
            # GIF89a + logical screen width=2, height=3.
            p.write_bytes(b"GIF89a\x02\x00\x03\x00\x00\x00\x00")
            self.assertEqual(image_size(p), (2, 3))


if __name__ == "__main__":
    unittest.main()
