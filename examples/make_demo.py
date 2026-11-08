"""Generate an original, self-contained three-slide fixture and relative plans.

This is a functional example, not a benchmark of presentation design quality.
All artwork and copy are synthetic. No network, private templates, or font files
are needed. The generated template is static; native animation is a separate,
optional Windows PowerPoint step.
"""

from __future__ import annotations

import argparse
import json
import os
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


INK = "17334A"
ACCENT = "187F83"
PAPER = "F5F7F9"


def artwork(path: Path, replacement: bool) -> None:
    """Draw geometric sample artwork; this is not a scientific illustration."""
    size = (900, 600) if replacement else (800, 600)
    image = Image.new("RGB", size, "#eaf2f4" if replacement else "#f0edf7")
    draw = ImageDraw.Draw(image)
    width, height = size
    for x in range(0, width, 60):
        draw.line((x, 0, x, height), fill="#d5e2e6", width=1)
    for y in range(0, height, 60):
        draw.line((0, y, width, y), fill="#d5e2e6", width=1)
    colors = ("#187f83", "#edae49", "#17334a") if replacement else ("#7564a0", "#b4a7d6", "#4b4369")
    draw.ellipse((width * 0.10, height * 0.18, width * 0.51, height * 0.81), fill=colors[0])
    draw.rectangle((width * 0.43, height * 0.28, width * 0.83, height * 0.72), fill=colors[1])
    draw.polygon([(width * 0.60, height * 0.12), (width * 0.91, height * 0.86),
                  (width * 0.31, height * 0.86)], outline=colors[2], width=10)
    image.save(path)


def text_box(slide, name: str, text: str, x: float, y: float, width: float,
             height: float, size: float, bold: bool = False, color: str = INK):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    shape.name = name
    frame = shape.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = "Arial"
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = RGBColor.from_string(color)
    return shape


def fill_shape(slide, name: str, kind, x: float, y: float, width: float,
               height: float, color: str):
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(width), Inches(height))
    shape.name = name
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color)
    shape.line.fill.background()
    return shape


def create_demo(output_dir: Path) -> dict:
    output_dir = output_dir.resolve()
    if output_dir.exists():
        raise ValueError("Choose a new output directory; the demo never overwrites existing files")
    output_dir.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="ppt-demo-", dir=output_dir.parent) as temporary:
        stage = Path(temporary)
        (stage / "templates").mkdir()
        (stage / "assets").mkdir()
        original_image = stage / "assets" / "original.png"
        replacement_image = stage / "assets" / "replacement.png"
        artwork(original_image, replacement=False)
        artwork(replacement_image, replacement=True)

        deck = Presentation()
        deck.slide_width, deck.slide_height = Inches(13.333333), Inches(7.5)
        deck.core_properties.title = "Synthetic reusable template"
        deck.core_properties.subject = "Self-contained functional demonstration"
        deck.core_properties.author = "ppt-template-animator contributors"
        deck.core_properties.keywords = "synthetic, demo, editable"
        choices = []

        cover = deck.slides.add_slide(deck.slide_layouts[6])
        fill_shape(cover, "CoverAccent", MSO_SHAPE.RECTANGLE, 0.65, 1.1, 0.12, 4.9, ACCENT)
        title = text_box(cover, "CoverTitle", "Template cover title", 1.1, 1.4, 11.1, 1.7, 48, True)
        subtitle = text_box(cover, "CoverSubtitle", "Template supporting text", 1.15, 3.6, 10.6, 1.4, 25)
        footer = text_box(cover, "CoverFooter", "Template footer", 1.15, 6.6, 10.8, 0.35, 13)
        choices.append({"number": 1, "text": {
            str(title.shape_id): "Reproducible PPT demo",
            str(subtitle.shape_id): "Generate, adapt, and verify editable slides using only repository content.",
            str(footer.shape_id): "Synthetic example. No private templates or external assets."
        }, "notes": "This fixture demonstrates the tools. All content and artwork are synthetic."})

        evidence = deck.slides.add_slide(deck.slide_layouts[6])
        title = text_box(evidence, "EvidenceTitle", "Template evidence title", 0.65, 0.5, 12, 0.9, 34, True)
        picture = evidence.shapes.add_picture(str(original_image), Inches(0.65), Inches(1.7),
                                             Inches(7.2), Inches(4.8))
        picture.name = "EvidenceImage"
        caption = text_box(evidence, "EvidenceCaption", "Template picture caption", 8.3, 2.0, 4.2, 2.7, 24)
        footer = text_box(evidence, "EvidenceFooter", "Template source note", 0.65, 6.9, 12, 0.3, 12)
        choices.append({"number": 2, "text": {
            str(title.shape_id): "Replace an image while keeping its frame",
            str(caption.shape_id): "The plan replaces this generated picture. The caption and frame remain editable.",
            str(footer.shape_id): "Synthetic geometric artwork generated locally; not research evidence."
        }, "pictures": {str(picture.shape_id): "assets/replacement.png"},
            "notes": "The picture is synthetic artwork generated by examples/make_demo.py. Verify its crop visually."})

        process = deck.slides.add_slide(deck.slide_layouts[6])
        title = text_box(process, "ProcessTitle", "Template process title", 0.65, 0.5, 12, 0.9, 34, True)
        steps = []
        for index, x in enumerate((0.8, 5.0, 9.2), 1):
            fill_shape(process, f"StepBackground{index}", MSO_SHAPE.ROUNDED_RECTANGLE,
                       x, 2.1, 3.3, 2.35, PAPER)
            steps.append(text_box(process, ("StepOne", "StepTwo", "StepThree")[index - 1],
                                  f"Template step {index}", x + 0.22, 2.6, 2.85, 1.3, 26, True))
        fill_shape(process, "ArrowOne", MSO_SHAPE.RIGHT_ARROW, 4.3, 2.93, 0.45, 0.5, ACCENT)
        fill_shape(process, "ArrowTwo", MSO_SHAPE.RIGHT_ARROW, 8.5, 2.93, 0.45, 0.5, ACCENT)
        footer = text_box(process, "ProcessFooter", "Template process footer", 0.8, 5.6, 11.8, 1.0, 21)
        choices.append({"number": 3, "text": {
            str(title.shape_id): "A small, testable production workflow",
            str(steps[0].shape_id): "1. Generate",
            str(steps[1].shape_id): "2. Adapt",
            str(steps[2].shape_id): "3. Verify",
            str(footer.shape_id): "Optional native animations are added in a separate Windows PowerPoint step."
        }, "notes": "This is a synthetic workflow illustration. The three steps and arrows are native editable objects."})
        deck.save(str(stage / "templates" / "demo-template.pptx"))

        build_plan = {"template": "templates/demo-template.pptx", "output": "adapted.pptx",
                      "require_animation": False, "strict_text": True, "native_validation": "off",
                      "slides": choices}
        animation_plan = {"slides": [
            {"transition": "fade", "effects": [
                ["CoverTitle", "fade", "after", 0.5],
                ["CoverSubtitle", "fade", "after", 0.4]]},
            {"transition": "fade", "effects": [
                ["EvidenceImage", "fade", "click", 0.6],
                ["EvidenceCaption", "fade", "with", 0.4]]},
            {"transition": "push", "effects": [
                ["StepOne", "fade", "click", 0.4],
                {"shape": "StepTwo", "effect": "fly", "trigger": "click", "duration": 0.55,
                 "direction": "up", "accelerate": 0.1, "decelerate": 0.2},
                ["ArrowOne", "fade", "with", 0.35],
                ["StepThree", "zoom", "click", 0.5],
                ["ArrowTwo", "fade", "with", 0.3]]}
        ]}
        ledger = {"example_only": True, "verification": "planned; inspect the generated final deck before marking verified",
                  "materials": [
                      {"material_id": "M001", "source": "build-plan.json", "locator": "slides[0].text",
                       "role": "content", "planned_use": "Cover title and supporting text",
                       "final_location": {"slide": 1, "shape_names": ["CoverTitle", "CoverSubtitle"]}},
                      {"material_id": "M002", "source": "assets/replacement.png", "role": "content",
                       "planned_use": "Visible synthetic example image",
                       "final_location": {"slide": 2, "shape_names": ["EvidenceImage"]}},
                      {"material_id": "M003", "source": "build-plan.json", "locator": "slides[2].text",
                       "role": "content", "planned_use": "Editable synthetic workflow",
                       "final_location": {"slide": 3, "shape_names": ["StepOne", "StepTwo", "StepThree"]}}
                  ]}
        for name, data in (("build-plan.json", build_plan), ("animation-plan.json", animation_plan),
                           ("material-ledger.json", ledger)):
            (stage / name).write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        os.replace(stage, output_dir)
    return {"output_dir": str(output_dir), "template": "templates/demo-template.pptx",
            "build_plan": "build-plan.json", "animation_plan": "animation-plan.json",
            "material_ledger": "material-ledger.json", "slides": 3,
            "synthetic": True, "native_animation_included": False}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True, help="New directory for generated fixture files")
    args = parser.parse_args()
    try:
        print(json.dumps(create_demo(args.output_dir), ensure_ascii=False, indent=2))
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
