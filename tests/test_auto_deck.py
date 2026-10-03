import tempfile
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from auto_deck import load_source, parse_markdown, parse_plain_text


class AutoDeckTest(unittest.TestCase):
    def test_markdown_becomes_slides_and_splits_long_bullets(self):
        text = """# Weekly Report

## Result
- one
- two
- three
- four
- five
- six
"""
        title, slides = parse_markdown(text, source_dir=Path("."), max_bullets=5)
        self.assertEqual(title, "Weekly Report")
        self.assertEqual(len(slides), 2)
        self.assertEqual(slides[0]["title"], "Result")
        self.assertEqual(slides[0]["blocks"][0]["items"], ["one", "two", "three", "four", "five"])
        self.assertEqual(slides[1]["blocks"][0]["items"], ["six"])

    def test_markdown_image_is_resolved_from_source_directory(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td)
            title, slides = parse_markdown(
                "# Demo\n\n## Method\nText\n\n![camera](images/frame.png)\n",
                source_dir=base,
            )
            self.assertEqual(title, "Demo")
            image = next(b for b in slides[0]["blocks"] if b["type"] == "image")
            self.assertEqual(Path(image["path"]), (base / "images/frame.png").resolve())
            self.assertEqual(slides[0]["layout"], "image-right")

    def test_json_compact_spec_is_preserved(self):
        with tempfile.TemporaryDirectory() as td:
            path = Path(td) / "deck.json"
            path.write_text('{"slides":[{"title":"A","bullets":["B"]}]}', encoding="utf-8")
            spec = load_source(
                path,
                aspect="16:9",
                footer="footer",
                preset="scientific",
                max_bullets=5,
            )
            self.assertEqual(spec["slides"][0]["title"], "A")
            self.assertEqual(spec["canvas"]["preset"], "16:9")
            self.assertEqual(spec["theme"]["footerText"], "footer")
            self.assertEqual(spec["meta"]["autoDeckPreset"], "scientific")

    def test_plain_text_uses_first_line_as_deck_title(self):
        title, slides = parse_plain_text(
            "Experiment Report\n\nMeasured displacement is stable.\n\nNext action is validation."
        )
        self.assertEqual(title, "Experiment Report")
        self.assertEqual(len(slides), 2)


if __name__ == "__main__":
    unittest.main()
