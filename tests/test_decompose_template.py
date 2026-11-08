"""Verify native component extraction preserves a group, media, and its effects."""

from __future__ import annotations

import hashlib
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from animate_deck import animate  # noqa: E402
from decompose_template import extract, inspect  # noqa: E402


class DecomposeTest(unittest.TestCase):
    def test_rejects_unknown_shape_before_creating_output(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source.pptx"
            deck = Presentation()
            deck.slides.add_slide(deck.slide_layouts[6])
            deck.save(source)
            with self.assertRaises(ValueError):
                extract(source, 1, [999], root / "components")
            self.assertFalse((root / "components").exists())

    @unittest.skipUnless(os.environ.get("PPT_NATIVE_TESTS") == "1", "Set PPT_NATIVE_TESTS=1 to run Windows PowerPoint integration tests")
    def test_extracts_editable_group_and_preserves_source(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            image = root / "picture.png"
            Image.new("RGB", (160, 100), "#147d92").save(image)
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            picture = slide.shapes.add_picture(str(image), Inches(1), Inches(1), Inches(3))
            caption = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(3), Inches(1))
            caption.text = "Editable evidence caption"
            group = slide.shapes.add_group_shape([picture, caption])
            group.name = "Evidence panel"
            group_id = group.shape_id
            other = slide.shapes.add_textbox(Inches(5), Inches(1), Inches(3), Inches(1))
            other.name = "Unselected title"
            other.text = "Remove this object"
            source, animated = root / "source.pptx", root / "animated.pptx"
            deck.save(source)
            plan = root / "animation.json"
            plan.write_text(json.dumps({"slides": [{"transition": "fade", "effects": [
                ["Evidence panel", "fade", "click", 0.5],
                ["Unselected title", "fade", "after", 0.4]
            ]}]}), encoding="utf-8")
            animate(source, plan, animated)
            before = hashlib.sha256(animated.read_bytes()).hexdigest()
            report = extract(animated, 1, [group_id], root / "components")
            self.assertEqual(before, hashlib.sha256(animated.read_bytes()).hexdigest())
            self.assertTrue(report["native_validated"])
            self.assertEqual(len(report["retained_motion"]["main_sequence"]), 1)
            self.assertEqual(report["media"][0]["sha256"], hashlib.sha256(image.read_bytes()).hexdigest())
            self.assertEqual(len(inspect(animated, 1)["shapes"]), 2)
            restored = Presentation(root / "components" / "components.pptx")
            shapes = restored.slides[0].shapes
            self.assertEqual(len(shapes), 1)
            self.assertEqual(shapes[0].shape_type, MSO_SHAPE_TYPE.GROUP)
            self.assertEqual(len(shapes[0].shapes), 2)
            self.assertIn("Editable evidence caption", [s.text for s in shapes[0].shapes if s.has_text_frame])


if __name__ == "__main__":
    unittest.main()
