"""Source motion research must preserve raw semantics and adjacent-slide context."""
from __future__ import annotations

import hashlib
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from study_motion import A, P, R, REL, compatibility_tree, motion_info, resolve_part, select_slides, study  # noqa: E402

MORPH = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"


def fixture(path: Path, prefix: str = "p", escape: bool = False):
    target = "../../outside.xml" if escape else "slides/slide9.xml"
    presentation = f'''<{prefix}:presentation xmlns:{prefix}="{P}" xmlns:r="{R}">
      <{prefix}:sldIdLst><{prefix}:sldId id="300" r:id="last"/>
      <{prefix}:sldId id="301" r:id="first"/></{prefix}:sldIdLst></{prefix}:presentation>'''
    rels = f'''<Relationships xmlns="{REL}">
      <Relationship Id="first" Type="{R}/slide" Target="slides/slide2.xml"/>
      <Relationship Id="last" Type="{R}/slide" Target="{target}"/></Relationships>'''
    slide1 = f'''<{prefix}:sld xmlns:{prefix}="{P}" xmlns:a="{A}">
      <{prefix}:cSld><{prefix}:spTree><{prefix}:sp><{prefix}:nvSpPr>
      <{prefix}:cNvPr id="8" name="!!hero"/></{prefix}:nvSpPr><{prefix}:txBody><a:p>
      <a:r><a:t>First in presentation order</a:t></a:r></a:p></{prefix}:txBody>
      </{prefix}:sp></{prefix}:spTree></{prefix}:cSld></{prefix}:sld>'''
    slide2 = f'''<{prefix}:sld xmlns:{prefix}="{P}" xmlns:a="{A}" xmlns:m="{MORPH}">
      <{prefix}:cSld><{prefix}:spTree><{prefix}:grpSp><{prefix}:nvGrpSpPr>
      <{prefix}:cNvPr id="2" name="PictureGroup"/></{prefix}:nvGrpSpPr>
      <{prefix}:grpSpPr><a:xfrm rot="60000"><a:off x="10" y="20"/><a:ext cx="1000" cy="500"/>
      <a:chOff x="0" y="0"/><a:chExt cx="1000" cy="500"/></a:xfrm></{prefix}:grpSpPr>
      <{prefix}:sp><{prefix}:nvSpPr><{prefix}:cNvPr id="3" name="!!hero"/></{prefix}:nvSpPr>
      <{prefix}:spPr><a:xfrm><a:off x="30" y="40"/><a:ext cx="300" cy="400"/></a:xfrm></{prefix}:spPr>
      <{prefix}:txBody><a:p><a:r><a:t>Second content</a:t></a:r></a:p></{prefix}:txBody>
      </{prefix}:sp></{prefix}:grpSp><{prefix}:sp><{prefix}:nvSpPr>
      <{prefix}:cNvPr id="9" name="TriggerButton"/></{prefix}:nvSpPr></{prefix}:sp>
      </{prefix}:spTree></{prefix}:cSld>
      <m:transition dur="1250"><m:morph option="byObject"/></m:transition>
      <{prefix}:timing><{prefix}:tnLst><{prefix}:seq><{prefix}:cTn id="10" nodeType="interactiveSeq"
      presetClass="path" presetID="64" presetSubtype="1" dur="indefinite" repeatCount="3000" repeatDur="9000">
      <{prefix}:stCondLst><{prefix}:cond evt="onClick" delay="indefinite"><{prefix}:tgtEl>
      <{prefix}:spTgt spid="9"/></{prefix}:tgtEl></{prefix}:cond></{prefix}:stCondLst>
      <{prefix}:childTnLst><{prefix}:animMotion path="M 0 0 L .5 .1 E" origin="layout" pathEditMode="relative">
      <{prefix}:cBhvr><{prefix}:cTn id="11" dur="450" accel="12000"><{prefix}:stCondLst>
      <{prefix}:cond delay="-1"/></{prefix}:stCondLst></{prefix}:cTn><{prefix}:tgtEl>
      <{prefix}:spTgt spid="3"><{prefix}:txEl><{prefix}:pRg st="0" end="1"/></{prefix}:txEl>
      </{prefix}:spTgt></{prefix}:tgtEl></{prefix}:cBhvr></{prefix}:animMotion>
      <{prefix}:animRot by="21600000"><{prefix}:cBhvr><{prefix}:tgtEl><{prefix}:spTgt spid="9"/>
      </{prefix}:tgtEl></{prefix}:cBhvr></{prefix}:animRot>
      <{prefix}:animScale zoomContents="1"><{prefix}:cBhvr><{prefix}:tgtEl><{prefix}:spTgt spid="3"/>
      </{prefix}:tgtEl></{prefix}:cBhvr><{prefix}:by x="200000" y="150000"/></{prefix}:animScale>
      <{prefix}:animEffect filter="fade"><{prefix}:cBhvr><{prefix}:tgtEl><{prefix}:spTgt spid="99"/>
      </{prefix}:tgtEl></{prefix}:cBhvr></{prefix}:animEffect>
      </{prefix}:childTnLst></{prefix}:cTn></{prefix}:seq></{prefix}:tnLst></{prefix}:timing></{prefix}:sld>'''
    media_rels = f'''<Relationships xmlns="{REL}">
      <Relationship Id="aud" Type="{R}/audio" Target="../media/audio1.wav"/>
      <Relationship Id="vid" Type="{R}/video" Target="https://example.invalid/remote.mp4" TargetMode="External"/>
      </Relationships>'''
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in {"ppt/presentation.xml": presentation, "ppt/_rels/presentation.xml.rels": rels,
                           "ppt/slides/slide9.xml": slide1, "ppt/slides/slide2.xml": slide2,
                           "ppt/slides/_rels/slide2.xml.rels": media_rels,
                           "ppt/media/audio1.wav": "fixture-only"}.items():
            archive.writestr(name, data)


class StudyMotionTest(unittest.TestCase):
    def test_relationship_order_and_modern_morph_keep_neighbor_context(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            fixture(source)
            before = source.read_bytes()
            report = study(source, "2")
            self.assertEqual(report["source"]["sha256"], hashlib.sha256(before).hexdigest())
            self.assertEqual(source.read_bytes(), before)
            self.assertEqual(study(source)["slides"][0]["part"], "ppt/slides/slide9.xml")
            slide = report["slides"][0]
            self.assertEqual(slide["part"], "ppt/slides/slide2.xml")
            self.assertTrue(slide["transition"]["morph"])
            self.assertEqual(slide["transition"]["raw"][0]["attributes"]["dur"], "1250")
            self.assertEqual(slide["dependencies"][0]["forced_match_names"], ["!!hero"])
            self.assertFalse(slide["dependencies"][0]["both_selected"])
            self.assertTrue(any("1-2" in value for value in slide["warnings"]))
            both = study(source, "1-2")["slides"][0]["dependencies"][0]
            self.assertTrue(both["both_selected"])

    def test_raw_motion_timing_and_nested_shape_identity_are_preserved(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            fixture(source)
            slide = study(source, "2")["slides"][0]
            child = next(shape for shape in slide["shapes"] if shape["id"] == "3")
            self.assertEqual((child["parent_group_id"], child["top_level_id"]), ("2", "2"))
            self.assertEqual(child["geometry"][0]["children"][0]["attributes"], {"x": "30", "y": "40"})
            motion = slide["motion"]
            self.assertEqual(motion["interactive_sequence_count"], 1)
            preset = motion["preset_groups"][0]
            self.assertEqual(preset["attributes"]["repeatCount"], "3000")
            self.assertEqual(preset["attributes"]["presetSubtype"], "1")
            self.assertEqual(preset["conditions"][0]["children"][0]["attributes"],
                             {"evt": "onClick", "delay": "indefinite"})
            path, rotation, scale = motion["behaviors"][:3]
            self.assertEqual(path["parameters"]["path"], "M 0 0 L .5 .1 E")
            self.assertEqual(path["targets"][0]["shape"]["top_level_id"], "2")
            self.assertEqual(path["ancestor_timing"][0]["id"], "10")
            self.assertEqual(rotation["parameters"]["by"], "21600000")
            self.assertEqual(scale["raw"]["children"][-1]["attributes"], {"x": "200000", "y": "150000"})
            self.assertEqual(motion["timing_nodes"][1]["conditions"][0]["children"][0]["attributes"]["delay"], "-1")
            self.assertEqual(motion["unresolved_target_ids"], ["99"])

    def test_namespace_prefix_changes_do_not_change_motion_semantics(self):
        with tempfile.TemporaryDirectory() as directory:
            source, renamed = Path(directory) / "source.pptx", Path(directory) / "renamed.pptx"
            fixture(source)
            fixture(renamed, "q")
            first, second = study(source, "2")["slides"][0], study(renamed, "2")["slides"][0]
            for key in ("motion", "transition", "shapes", "dependencies"):
                self.assertEqual(first[key], second[key])

    def test_media_and_external_links_are_hints_not_downloads(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            fixture(source)
            hints = study(source, "2")["slides"][0]["relationship_hints"]
            self.assertEqual(hints[0]["package_part"], "ppt/media/audio1.wav")
            self.assertTrue(hints[0]["part_exists"])
            self.assertTrue(hints[1]["external"])
            self.assertIsNone(hints[1]["part_exists"])

    def test_relationships_cannot_escape_package(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.pptx"
            fixture(source, escape=True)
            with self.assertRaisesRegex(ValueError, "escapes"):
                study(source)
        for target in ("https://example.invalid/a.xml", "..\\other.xml", "%2e%2e/%2e%2e/outside.xml"):
            with self.subTest(target=target), self.assertRaises(ValueError):
                resolve_part("ppt/presentation.xml", target)

    def test_selection_is_bounded_and_follows_presentation_order(self):
        self.assertEqual(select_slides("7,1-4,3", 7), [1, 2, 3, 4, 7])
        for value in ("", "0", "2-1", "1,", "-1", "1-8", "1-", "1;2"):
            with self.subTest(value=value), self.assertRaises(ValueError):
                select_slides(value, 7)

    def test_modern_choice_does_not_double_count_its_fallback(self):
        xml = f'''<p:sld xmlns:p="{P}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          xmlns:modern="http://schemas.microsoft.com/office/powerpoint/2010/main">
          <mc:AlternateContent><mc:Choice Requires="modern"><p:animMotion path="M 0 0 L 1 0 E"/>
          <p:animScale/></mc:Choice><mc:Fallback><p:animMotion path="M 0 0 L 1 0 E"/>
          <p:animScale/></mc:Fallback></mc:AlternateContent></p:sld>'''.encode()
        raw, effective, branches = compatibility_tree(xml)
        self.assertEqual(len(list(raw.iter(f"{{{P}}}animMotion"))), 2)
        self.assertEqual([item["kind"] for item in motion_info(effective, [])["behaviors"]], ["animMotion", "animScale"])
        self.assertEqual(branches[0]["status"], "known_modern_choice")
        unknown = xml.replace(b"powerpoint/2010/main", b"powerpoint/2099/main")
        _, effective, branches = compatibility_tree(unknown)
        self.assertEqual(branches[0]["status"], "unresolved_host_support")
        self.assertEqual(branches[0]["selected"], "Fallback")
        self.assertEqual(len(motion_info(effective, [])["behaviors"]), 2)


if __name__ == "__main__":
    unittest.main()
