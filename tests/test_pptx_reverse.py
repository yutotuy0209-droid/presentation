#!/usr/bin/env python3
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from pptx_reverse import ir_to_markdown  # noqa: E402


class PptxReverseTests(unittest.TestCase):
    def test_markdown_contains_slide_text_and_geometry(self):
        deck = {
            "meta": {"filename": "old.pptx"},
            "slides": [{
                "number": 1,
                "elements": [{
                    "id": "2",
                    "name": "Title 1",
                    "type": "sp",
                    "x": 914400,
                    "y": 457200,
                    "width": 3657600,
                    "height": 914400,
                    "text": "研究目的",
                    "runs": [{"text": "研究目的", "fontSizePt": 28}],
                }],
            }],
        }
        md = ir_to_markdown(deck)
        self.assertIn("Slide 1｜研究目的", md)
        self.assertIn("Title 1", md)
        self.assertIn("x=1.000in", md)
        self.assertIn("研究目的", md)


if __name__ == "__main__":
    unittest.main()
