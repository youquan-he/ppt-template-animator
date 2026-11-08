"""Render selected template pages and report their real PowerPoint animation state."""

from __future__ import annotations

import argparse
import json
import time
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps
from pptx import Presentation

from template_catalog import inspect_file
from native_powerpoint import NativePowerPointUnavailable, create_powerpoint


def extract(source: Path, numbers: list[int], output: Path) -> None:
    deck = Presentation(str(source))
    ids = deck.slides._sldIdLst
    original = list(ids)
    selected = [original[number - 1] for number in numbers]
    for item in original:
        if item not in selected:
            deck.part.drop_rel(item.rId)
            ids.remove(item)
    for item in selected:
        ids.remove(item)
        ids.append(item)
    deck.save(str(output))


def make_contact_sheet(paths: list[Path], labels: list[str], output: Path) -> None:
    thumb_w, thumb_h, pad, label_h, columns = 640, 360, 20, 34, 3
    rows = (len(paths) + columns - 1) // columns
    sheet = Image.new("RGB", (columns * (thumb_w + pad) + pad,
                              rows * (thumb_h + label_h + pad) + pad), "#e7e9e9")
    draw = ImageDraw.Draw(sheet)
    for index, (path, label) in enumerate(zip(paths, labels)):
        x = pad + index % columns * (thumb_w + pad)
        y = pad + index // columns * (thumb_h + label_h + pad)
        with Image.open(path) as im:
            thumb = ImageOps.contain(im.convert("RGB"), (thumb_w, thumb_h),
                                     Image.Resampling.LANCZOS)
            sheet.paste(thumb, (x + (thumb_w - thumb.width) // 2,
                                y + (thumb_h - thumb.height) // 2))
        draw.text((x + 5, y + thumb_h + 5), label, fill="#24313a", font=ImageFont.load_default())
    sheet.save(output)


def preview(source: Path, numbers: list[int], output_dir: Path, video: bool) -> dict:
    if not source.is_file():
        raise ValueError(f"Template not found: {source}")
    if not numbers or len(numbers) != len(set(numbers)):
        raise ValueError("Provide unique 1-based slide numbers")
    metadata = inspect_file(source, include_slides=True, selected=set(numbers))
    if min(numbers) < 1 or max(numbers) > metadata["slide_count"]:
        raise ValueError("Selected slide number is outside the template")
    output_dir.mkdir(parents=True, exist_ok=True)
    report = {"source": str(source.resolve()), "slides": []}
    images = []
    app = create_powerpoint()
    presentation = None
    try:
        presentation = app.Presentations.Open(str(source.resolve()), True, False, False)
        export_height = round(1600 * presentation.PageSetup.SlideHeight /
                              presentation.PageSetup.SlideWidth)
        details = {item["number"]: item for item in metadata["slides"]}
        for number in numbers:
            slide = presentation.Slides.Item(number)
            image_path = output_dir / f"slide-{number}.png"
            slide.Export(str(image_path.resolve()), "PNG", 1600, export_height)
            images.append(image_path)
            layout_warnings = []
            for shape_index in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes.Item(shape_index)
                if (shape.Left < -1 or shape.Top < -1 or
                        shape.Left + shape.Width > presentation.PageSetup.SlideWidth + 1 or
                        shape.Top + shape.Height > presentation.PageSetup.SlideHeight + 1):
                    layout_warnings.append({"shape": shape.Name, "issue": "outside-slide"})
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    frame = shape.TextFrame2
                    available_height = shape.Height - frame.MarginTop - frame.MarginBottom
                    available_width = shape.Width - frame.MarginLeft - frame.MarginRight
                    if frame.TextRange.BoundHeight > available_height + 2:
                        layout_warnings.append({"shape": shape.Name, "issue": "text-height-overflow"})
                    if frame.TextRange.BoundWidth > available_width + 2:
                        layout_warnings.append({"shape": shape.Name, "issue": "text-width-overflow"})
            report["slides"].append({
                "number": number,
                "image": str(image_path),
                "powerpoint_effects": slide.TimeLine.MainSequence.Count,
                "powerpoint_transition": slide.SlideShowTransition.EntryEffect,
                "layout_warnings": layout_warnings,
                "xml": {key: details[number][key] for key in (
                    "effect_count", "animation_primitives", "transition_type",
                    "text_chars", "text_shape_count", "picture_count",
                    "largest_picture_area_fraction")},
            })
    finally:
        if presentation is not None:
            presentation.Close()
        app.Quit()
    make_contact_sheet(images, [f"slide {number}" for number in numbers],
                       output_dir / "contact-sheet.png")
    if video:
        selected_path = output_dir / "selected.pptx"
        extract(source, numbers, selected_path)
        app = create_powerpoint()
        presentation = None
        try:
            presentation = app.Presentations.Open(str(selected_path.resolve()), True, False, False)
            video_path = output_dir / "animation-preview.mp4"
            presentation.CreateVideo(str(video_path.resolve()), False, 4, 720, 15, 65)
            deadline = time.monotonic() + 300
            while presentation.CreateVideoStatus not in (3, 4):
                if time.monotonic() > deadline:
                    raise TimeoutError("PowerPoint video export timed out")
                time.sleep(1)
            if presentation.CreateVideoStatus != 3 or not video_path.is_file():
                raise RuntimeError("PowerPoint video export failed")
            report["video"] = str(video_path)
        finally:
            if presentation is not None:
                presentation.Close()
            app.Quit()
    (output_dir / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                                             encoding="utf-8")
    return report


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("template", type=Path)
    parser.add_argument("--slides", required=True, help="Comma-separated 1-based slide numbers")
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--video", action="store_true", help="Export a short playback preview")
    args = parser.parse_args()
    try:
        result = preview(args.template, [int(x) for x in args.slides.split(",")],
                         args.output_dir, args.video)
    except NativePowerPointUnavailable as exc:
        parser.exit(2, f"Error: {exc}\n")
    print(json.dumps({"slides": len(result["slides"]),
                      "output_dir": str(args.output_dir.resolve()),
                      "video": result.get("video")}, ensure_ascii=False))
