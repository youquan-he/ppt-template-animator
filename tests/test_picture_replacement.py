"""Regression checks for replacing a cropped template picture."""

from __future__ import annotations

import hashlib
import tempfile
import unittest
from pathlib import Path
from sys import path as python_path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.util import Inches


python_path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from build_deck import A, A14, put_picture  # noqa: E402


class PictureReplacementTest(unittest.TestCase):
    def test_replaces_old_crop_and_image_effect(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old_image, new_image = root / "old.png", root / "new.png"
            Image.new("RGB", (300, 200), "red").save(old_image)
            Image.new("RGB", (200, 300), "blue").save(new_image)

            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            picture = slide.shapes.add_picture(str(old_image), 0, 0, Inches(4), Inches(4))
            picture.crop_left = 0.1
            picture.crop_right = 0.1
            blip = picture._element.blipFill.blip
            extensions = etree.SubElement(blip, f"{{{A}}}extLst")
            extension = etree.SubElement(extensions, f"{{{A}}}ext")
            etree.SubElement(extension, f"{{{A14}}}imgProps")

            put_picture(slide, picture, new_image)

            self.assertEqual(len(picture._element.blipFill.findall(f"{{{A}}}srcRect")), 1)
            self.assertEqual(len(blip.findall(f".//{{{A14}}}imgProps")), 0)
            self.assertEqual(hashlib.sha256(picture.image.blob).digest(),
                             hashlib.sha256(new_image.read_bytes()).digest())
            saved = root / "result.pptx"
            deck.save(saved)
            self.assertEqual(len(Presentation(saved).slides), 1)


if __name__ == "__main__":
    unittest.main()
