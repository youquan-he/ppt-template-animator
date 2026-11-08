"""Protect template animation semantics and previous outputs during rebuilds."""

from __future__ import annotations

import io
from copy import deepcopy
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from lxml import etree
from pptx import Presentation
from pptx.presentation import Presentation as PresentationModel
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from build_deck import MC, P, build, motion_state, verify_motion  # noqa: E402


def slide_xml(prefix: str = "p", compact: bool = False):
    xml = f'''<{prefix}:sld xmlns:{prefix}="{P}" xmlns:mc="{MC}" mc:Ignorable="{prefix}">
      <{prefix}:cSld><{prefix}:spTree><{prefix}:sp><{prefix}:nvSpPr>
        <{prefix}:cNvPr id="2" name="Figure"/>
      </{prefix}:nvSpPr></{prefix}:sp></{prefix}:spTree></{prefix}:cSld>
      <{prefix}:transition spd="med"><{prefix}:fade/></{prefix}:transition>
      <{prefix}:timing mc:Ignorable="{prefix}"><{prefix}:tnLst>
        <{prefix}:par><{prefix}:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <{prefix}:childTnLst><{prefix}:animMotion path="M 0 0 L .5 0 E">
            <{prefix}:cBhvr><{prefix}:cTn id="2" dur="700"><{prefix}:stCondLst>
              <{prefix}:cond delay="150"/>
            </{prefix}:stCondLst></{prefix}:cTn><{prefix}:tgtEl>
              <{prefix}:spTgt spid="2"/>
            </{prefix}:tgtEl></{prefix}:cBhvr>
          </{prefix}:animMotion></{prefix}:childTnLst>
        </{prefix}:cTn></{prefix}:par>
      </{prefix}:tnLst></{prefix}:timing>
    </{prefix}:sld>'''
    parser = etree.XMLParser(remove_blank_text=compact)
    return etree.fromstring(xml.encode(), parser)


class AnimationPreservationTest(unittest.TestCase):
    def test_prefixes_and_formatting_do_not_change_meaning(self) -> None:
        before = motion_state(slide_xml())
        after = slide_xml("x", compact=True)
        after.find(f"{{{P}}}timing").insert(0, etree.Comment("formatting-only comment"))
        verify_motion(before, after, 1)

    def test_changed_duration_delay_path_and_target_are_rejected(self) -> None:
        before = motion_state(slide_xml())
        for tag, attribute, value in (("cTn", "dur", "900"), ("cond", "delay", "200"),
                                      ("animMotion", "path", "M 0 0 L 0 .5 E"),
                                      ("spTgt", "spid", "3")):
            with self.subTest(attribute=attribute):
                after = slide_xml()
                list(after.iter(f"{{{P}}}{tag}"))[-1].set(attribute, value)
                with self.assertRaisesRegex(ValueError, "Animation structure changed"):
                    verify_motion(before, after, 1)

    def test_reordered_effects_are_rejected(self) -> None:
        source = slide_xml()
        first = next(source.iter(f"{{{P}}}animMotion"))
        second = deepcopy(first)
        second.set("path", "M 0 0 L 0 .5 E")
        parent = first.getparent()
        parent.append(second)
        before = motion_state(source)
        parent.remove(first)
        parent.append(first)
        with self.assertRaisesRegex(ValueError, "Animation structure changed"):
            verify_motion(before, source, 1)

    def test_dangling_target_is_rejected_even_with_identical_timing(self) -> None:
        source = slide_xml()
        before = motion_state(source)
        shape = next(source.iter(f"{{{P}}}sp"))
        shape.getparent().remove(shape)
        with self.assertRaisesRegex(ValueError, "missing shape targets"):
            verify_motion(before, source, 1)

    def test_transition_change_requires_explicit_expected_state(self) -> None:
        source = slide_xml()
        before = motion_state(source)
        transition = source.find(f"{{{P}}}transition")
        transition[0].tag = f"{{{P}}}push"
        with self.assertRaisesRegex(ValueError, "Transition structure changed"):
            verify_motion(before, source, 1)
        verify_motion(before, source, 1, compare_transition=False)
        verify_motion(motion_state(source), source, 1)

    def test_save_corruption_does_not_overwrite_previous_output(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            template, output = root / "source.pptx", root / "output.pptx"
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
            shape.text = "Source caption"
            timing = slide_xml().find(f"{{{P}}}timing")
            next(timing.iter(f"{{{P}}}spTgt")).set("spid", str(shape.shape_id))
            slide._element.append(timing)
            deck.save(template)
            output.write_bytes(b"previous-good-output")
            original_save = PresentationModel.save

            def tampered_save(presentation, path):
                original_save(presentation, path)
                with zipfile.ZipFile(path) as archive:
                    parts = {name: archive.read(name) for name in archive.namelist()}
                xml = etree.fromstring(parts["ppt/slides/slide1.xml"])
                list(xml.iter(f"{{{P}}}cTn"))[-1].set("dur", "999")
                parts["ppt/slides/slide1.xml"] = etree.tostring(xml)
                saved = io.BytesIO()
                with zipfile.ZipFile(saved, "w", zipfile.ZIP_DEFLATED) as archive:
                    for name, data in parts.items():
                        archive.writestr(name, data)
                Path(path).write_bytes(saved.getvalue())

            plan = {"template": str(template), "output": str(output),
                    "slides": [{"number": 1, "text": {str(shape.shape_id): "New caption"}}]}
            with patch.object(PresentationModel, "save", tampered_save):
                with self.assertRaisesRegex(ValueError, "Animation structure changed"):
                    build(plan, root)
            self.assertEqual(output.read_bytes(), b"previous-good-output")
            self.assertFalse(list(root.glob("ppt-build-*.pptx")))


if __name__ == "__main__":
    unittest.main()
