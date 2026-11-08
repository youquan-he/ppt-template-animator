"""Check that planned motion settings survive a native PowerPoint save."""

from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from animate_deck import animate, normalize_effect  # noqa: E402


class AnimationOptionsTest(unittest.TestCase):
    def test_validates_motion_settings(self) -> None:
        effect = normalize_effect({"shape": "Figure", "effect": "fly", "trigger": "after",
                                   "duration": 0.7, "delay": 0.15, "direction": "up",
                                   "accelerate": 0.15, "decelerate": 0.2}, 1)
        self.assertEqual(effect["direction"], "up")
        with self.assertRaises(ValueError):
            normalize_effect({**effect, "effect": "fade"}, 1)

    @unittest.skipUnless(os.environ.get("PPT_NATIVE_TESTS") == "1", "Set PPT_NATIVE_TESTS=1 to run Windows PowerPoint integration tests")
    def test_native_save_preserves_motion_settings(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output, plan = root / "source.pptx", root / "result.pptx", root / "plan.json"
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1),
                                           Inches(2), Inches(1))
            shape.name = "Figure"
            deck.save(source)
            plan.write_text(json.dumps({"slides": [{"transition": "fade", "effects": [
                {"shape": "Figure", "effect": "fly", "trigger": "after",
                 "duration": 0.7, "delay": 0.15, "direction": "up",
                 "accelerate": 0.15, "decelerate": 0.2}
            ]}]}), encoding="utf-8")
            animate(source, plan, output)
            self.assertTrue(output.is_file())


if __name__ == "__main__":
    unittest.main()
