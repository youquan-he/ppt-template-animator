"""Check that planned motion settings survive a native PowerPoint save."""

from __future__ import annotations

import hashlib
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from animate_deck import animate, motion_path_tokens, normalize_effect, same_motion_path  # noqa: E402


class AnimationOptionsTest(unittest.TestCase):
    def transform(self, **values) -> dict:
        return {"shape": "Figure", "effect": "transform", "trigger": "with",
                "duration": 1.4, **values}

    def test_validates_motion_settings(self) -> None:
        effect = normalize_effect({"shape": "Figure", "effect": "fly", "trigger": "after",
                                   "duration": 0.7, "delay": 0.15, "direction": "up",
                                   "accelerate": 0.15, "decelerate": 0.2}, 1)
        self.assertEqual(effect["direction"], "up")
        with self.assertRaises(ValueError):
            normalize_effect({**effect, "effect": "fade"}, 1)

    def test_combines_native_transform_components(self) -> None:
        effect = normalize_effect(self.transform(
            path="M 0 0 C .1 -.1 .2 -.1 .3 0 L .4 .1 E",
            scale=[125, 80], rotate=360, duration=12,
            accelerate=0.1, decelerate=0.25), 3)
        self.assertEqual(effect["scale"], [125, 80])
        self.assertEqual(effect["rotate"], 360)
        self.assertEqual(effect["duration"], 12)
        self.assertEqual(effect["delay"], 0)
        self.assertTrue(effect["start_visible"])
        self.assertFalse(normalize_effect(self.transform(rotate=90, start_visible=False), 1)["start_visible"])
        for component in ({"path": "M0 0 L.2 .1 E"}, {"scale": [100, 100]},
                          {"rotate": 0}):
            with self.subTest(component=component):
                normalize_effect(self.transform(**component), 1)

    def test_motion_paths_compare_canonical_tokens(self) -> None:
        self.assertTrue(same_motion_path("M 0,0 L .25,-.125 E",
                                        "M 0.00000 0.0 L 2.5e-1 -0.125000 E"))
        self.assertTrue(same_motion_path("M0 0 L.2 .1 Z E", "M0 0 L.20001 .1 Z E"))
        self.assertFalse(same_motion_path("M0 0 L.2 .1 E", "M0 0 L.21 .1 E"))
        self.assertEqual(motion_path_tokens("M0 0 C0 .1 .2 .3 .4 0 Z E"),
                         ("M", 0.0, 0.0, "C", 0.0, 0.1, 0.2, 0.3, 0.4, 0.0, "Z", "E"))

    def test_rejects_unsupported_or_incomplete_motion_paths(self) -> None:
        cases = [None, "", "M0 0 L1 1", "L0 0 E", "M0 0 E", "M0 0 L.1 E",
                 "M0 0 C0 .1 .2 .3 .4 E", "M0 0 Q.1 .2 .3 .4 E",
                 "m0 0 l.2 .1 e", "M0 0 L4.1 0 E", "M0 0 LNaN 0 E",
                 "M0 0 L1e999 0 E", "M0 0 L.2 .1 E E", "M0 0 Z L.1 .1 E",
                 "M0 0 L.2 .1 Z0 E", "M0 0 L.2 .1 E trailing", "M0 0 L.2 .1 E" * 500]
        for path in cases:
            with self.subTest(path=str(path)[:80]), self.assertRaises(ValueError):
                normalize_effect(self.transform(path=path), 1)

    def test_rejects_transform_configuration_errors(self) -> None:
        cases = [self.transform(), self.transform(scale=[0, 100]),
                 self.transform(scale=[100]), self.transform(scale=(100, 100)),
                 self.transform(scale=[True, 100]), self.transform(scale=[100, float("nan")]),
                 self.transform(rotate=-20), self.transform(rotate=float("inf")),
                 self.transform(rotate=3601), self.transform(scale=[100, 100], duration=31),
                 self.transform(rotate=5, duration=float("nan")),
                 self.transform(rotate=5, delay=float("inf")),
                 self.transform(rotate=5, accelerate=float("nan")),
                 self.transform(rotate=5, accelerate=0.7, decelerate=0.5),
                 self.transform(rotate=5, direction="up"),
                 self.transform(rotate=5, effect="fade"),
                 self.transform(rotate=5, start_visible=1),
                 self.transform(rotate=5, start_visible="true"),
                 {"shape": "Figure", "effect": "fade", "trigger": "with",
                  "duration": 1, "start_visible": True}]
        for spec in cases:
            with self.subTest(spec=spec), self.assertRaises(ValueError):
                normalize_effect(spec, 1)

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

    @unittest.skipUnless(os.environ.get("PPT_NATIVE_TESTS") == "1", "Set PPT_NATIVE_TESTS=1 to run Windows PowerPoint integration tests")
    def test_native_transform_preserves_initial_visibility(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output, plan = root / "source.pptx", root / "result.pptx", root / "plan.json"
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            ids = {}
            for index, name in enumerate(("VisibleInitially", "AppearsOnStart")):
                shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1 + 3 * index), Inches(1),
                                               Inches(2), Inches(1))
                shape.name = name
                ids[name] = str(shape.shape_id)
            deck.save(source)
            plan.write_text(json.dumps({"slides": [{"transition": "fade", "effects": [
                self.transform(shape="VisibleInitially", path="M0 0 L.1 0 E", delay=0.44),
                self.transform(shape="AppearsOnStart", path="M0 0 L.1 0 E", delay=0.44,
                               start_visible=False)
            ]}]}), encoding="utf-8")
            animate(source, plan, output)
            # Native AddEffect(Custom)'s Set visibility behavior hides its target
            # before the .44s effect delay. It must be absent for the default
            # visible transform, and retained only for explicit entrance semantics.
            with ZipFile(output) as archive:
                xml = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
            visibility_targets = {
                node.find(f"./{p}cBhvr/{p}tgtEl/{p}spTgt").get("spid")
                for node in xml.iter(p + "set")
                if any(item.text == "style.visibility" for item in node.iter(p + "attrName"))
            }
            self.assertNotIn(ids["VisibleInitially"], visibility_targets)
            self.assertIn(ids["AppearsOnStart"], visibility_targets)

    @unittest.skipUnless(os.environ.get("PPT_NATIVE_TESTS") == "1", "Set PPT_NATIVE_TESTS=1 to run Windows PowerPoint integration tests")
    def test_native_save_preserves_composite_transform_and_morph(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, output, plan = root / "source.pptx", root / "result.pptx", root / "plan.json"
            deck = Presentation()
            for index in range(2):
                slide = deck.slides.add_slide(deck.slide_layouts[6])
                shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1 + index), Inches(1),
                                               Inches(2), Inches(1.5))
                shape.name = "!!Figure"
            deck.save(source)
            source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
            plan.write_text(json.dumps({"slides": [
                {"transition": "fade", "transition_duration": 0.8, "effects": [
                    self.transform(shape="!!Figure", trigger="after", duration=1.4, delay=0.15,
                                   path="M0 0 C.06 -.1 .15 -.1 .22 0 E",
                                   scale=[120, 85], rotate=45,
                                   accelerate=0.1, decelerate=0.2)]},
                {"transition": "morph", "transition_duration": 1.7, "effects": []}
            ]}), encoding="utf-8")
            # animate() reopens the saved file and checks each native behavior's concrete
            # path, percentages, angle, timing, target and slide-transition duration.
            animate(source, plan, output)
            self.assertTrue(output.is_file())
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), source_hash)


if __name__ == "__main__":
    unittest.main()
