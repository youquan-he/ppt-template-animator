"""Check associated-theme discovery and distinguish explicit from inherited type."""

import sys
import tempfile
import unittest
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.opc.packuri import PackURI
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from decompose_template import A, inspect  # noqa: E402


class DesignTokensTest(unittest.TestCase):
    def test_reads_associated_theme_and_cjk_font_evidence(self):
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary) / "design.pptx"
            deck = Presentation()
            slide = deck.slides.add_slide(deck.slide_layouts[6])
            shape = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
            run = shape.text_frame.paragraphs[0].add_run()
            run.text = "中文图注"
            run.font.name = "Arial"
            run.font.size = Pt(23)
            run.font.color.rgb = RGBColor.from_string("2468AC")
            etree.SubElement(run._r.get_or_add_rPr(), f"{{{A}}}ea").set("typeface", "Microsoft YaHei")
            inherited = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
            inherited.text = "No directly declared font"
            master = slide.slide_layout.slide_master
            theme = next(rel.target_part for rel in master.part.rels.values() if rel.reltype.endswith("/theme"))
            theme.partname = PackURI("/ppt/theme/associated-theme.xml")
            xml = etree.fromstring(theme.blob)
            xml.set("name", "Test theme")
            scheme = xml.find(f"./{{{A}}}themeElements/{{{A}}}fontScheme/{{{A}}}minorFont")
            for font in scheme.findall(f"{{{A}}}font"):
                if font.get("script") == "Hans":
                    font.set("typeface", "Microsoft YaHei")
            theme._blob = etree.tostring(xml)
            deck.save(source)
            report = inspect(source, 1)
            self.assertEqual(report["design_system"]["theme"]["part"], "/ppt/theme/associated-theme.xml")
            self.assertEqual(report["design_system"]["theme"]["fonts"]["minor"]["script_fonts"]["Hans"], "Microsoft YaHei")
            style = report["shapes"][0]["direct_style"]
            self.assertEqual(style["font_names"], ["Arial", "Microsoft YaHei"])
            self.assertEqual(style["font_sizes_pt"], [23])
            self.assertIn("2468AC", style["rgb_colors"])
            self.assertEqual(report["shapes"][1]["direct_style"]["font_names"], [])


if __name__ == "__main__":
    unittest.main()
