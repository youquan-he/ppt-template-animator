"""Run the public demo without private assets, native Office, or a fixed cwd."""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "examples"))
sys.path.insert(0, str(ROOT / "scripts"))
from animate_deck import normalize_effect  # noqa: E402
from make_demo import create_demo  # noqa: E402


class PublicDemoTest(unittest.TestCase):
    def test_demo_builds_from_unrelated_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            target = root / "fixture"
            other_cwd = root / "other"
            other_cwd.mkdir()
            subprocess.run([sys.executable, "-X", "utf8", str(ROOT / "examples" / "make_demo.py"),
                            "--output-dir", str(target)], cwd=other_cwd, check=True,
                           capture_output=True, text=True, encoding="utf-8")
            plan_path = target / "build-plan.json"
            plan = json.loads(plan_path.read_text(encoding="utf-8"))
            self.assertEqual(plan["native_validation"], "off")
            self.assertFalse(Path(plan["template"]).is_absolute())
            self.assertFalse(Path(plan["output"]).is_absolute())
            source = target / plan["template"]
            source_hash = hashlib.sha256(source.read_bytes()).hexdigest()
            completed = subprocess.run([sys.executable, "-X", "utf8", str(ROOT / "scripts" / "build_deck.py"),
                                        str(plan_path)], cwd=other_cwd, check=True,
                                       capture_output=True, text=True, encoding="utf-8")
            result = json.loads(completed.stdout)
            self.assertFalse(result["native_validated"])
            self.assertEqual(source_hash, hashlib.sha256(source.read_bytes()).hexdigest())
            output = Presentation(str(target / plan["output"]))
            self.assertEqual(len(output.slides), 3)
            for number, choice in enumerate(plan["slides"]):
                shapes = {shape.shape_id: shape for shape in output.slides[number].shapes}
                for raw_id, expected in choice["text"].items():
                    self.assertEqual(shapes[int(raw_id)].text, expected)
                for raw_id, image_path in choice.get("pictures", {}).items():
                    self.assertFalse(Path(image_path).is_absolute())
                    self.assertEqual(hashlib.sha256(shapes[int(raw_id)].image.blob).digest(),
                                     hashlib.sha256((target / image_path).read_bytes()).digest())
            source_deck = Presentation(str(source))
            self.assertIn("Template cover title", [s.text for s in source_deck.slides[0].shapes if s.has_text_frame])
            animation = json.loads((target / "animation-plan.json").read_text(encoding="utf-8"))
            for number, slide_plan in enumerate(animation["slides"], 1):
                names = [shape.name for shape in output.slides[number - 1].shapes]
                for spec in slide_plan["effects"]:
                    effect = normalize_effect(spec, number)
                    self.assertEqual(names.count(effect["shape"]), 1)

    def test_generator_refuses_existing_directory(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(temporary)
            sentinel = output / "keep.txt"
            sentinel.write_text("Existing user content", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "new output directory"):
                create_demo(output)
            self.assertEqual(sentinel.read_text(encoding="utf-8"), "Existing user content")
            self.assertFalse((output / "templates").exists())


if __name__ == "__main__":
    unittest.main()
