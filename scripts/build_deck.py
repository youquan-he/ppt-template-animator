#!/usr/bin/env python
"""Build an animated PPTX from selected pages of one local template."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import sys
import tempfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from template_catalog import inspect_file
from native_powerpoint import NativePowerPointUnavailable, create_powerpoint


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XSI = "http://www.w3.org/2001/XMLSchema-instance"
TRANSITIONS = {"fade", "push-left", "push-right", "wipe-left", "wipe-right"}


def semantic_xml(element) -> tuple:
    """Compare XML meaning without depending on prefixes or indentation."""
    def text_value(value: str | None, owner) -> str | None:
        preserve = False
        while owner is not None:
            space = owner.get("{http://www.w3.org/XML/1998/namespace}space")
            if space is not None:
                preserve = space == "preserve"
                break
            owner = owner.getparent()
        return value if value and (value.strip() or preserve) else None

    def expanded_token(token: str, prefix_only: bool = False) -> str:
        prefix, separator, local = token.partition(":")
        namespace = element.nsmap.get(prefix)
        if namespace and (separator or prefix_only):
            return f"{{{namespace}}}{local if separator else ''}"
        return token

    attributes = []
    for name, value in element.attrib.items():
        if name == f"{{{MC}}}Ignorable" or (element.tag == f"{{{MC}}}Choice" and name == "Requires"):
            value = " ".join(sorted(expanded_token(token, True) for token in value.split()))
        elif name in {f"{{{MC}}}ProcessContent", f"{{{MC}}}PreserveElements", f"{{{MC}}}PreserveAttributes"}:
            value = " ".join(sorted(expanded_token(token) for token in value.split()))
        elif name == f"{{{XSI}}}type":
            value = expanded_token(value)
        attributes.append((name, value))
    children = tuple((semantic_xml(child), text_value(child.tail, element))
                     for child in element if isinstance(child.tag, str))
    return element.tag, tuple(sorted(attributes)), text_value(element.text, element), children


def motion_state(slide_element) -> dict:
    return {
        "timing": tuple(semantic_xml(node) for node in slide_element.iter(f"{{{P}}}timing")),
        "transition": tuple(semantic_xml(node) for node in slide_element.iter(f"{{{P}}}transition")),
    }


def verify_motion(expected: dict, slide_element, number: int, compare_transition: bool = True) -> None:
    actual = motion_state(slide_element)
    if expected["timing"] != actual["timing"]:
        raise ValueError(f"Animation structure changed on slide {number} (targets, order, triggers, timing, or effect parameters)")
    if compare_transition and expected["transition"] != actual["transition"]:
        raise ValueError(f"Transition structure changed on slide {number}")
    available_ids = {node.get("id") for node in slide_element.iter(f"{{{P}}}cNvPr")}
    targets = {node.get("spid") for timing in slide_element.iter(f"{{{P}}}timing")
               for node in timing.iter(f"{{{P}}}spTgt")}
    missing = targets - available_ids
    if missing:
        raise ValueError(f"Animation has missing shape targets on slide {number}: {sorted(str(value) for value in missing)}")


def resolve(base: Path, value: str) -> Path:
    path = Path(value).expanduser()
    return path if path.is_absolute() else base / path


def nested_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from nested_shapes(shape.shapes)


def shape_by_id(slide, shape_id: int):
    found = [shape for shape in nested_shapes(slide.shapes) if shape.shape_id == shape_id]
    if len(found) != 1:
        raise ValueError(f"Expected one shape ID {shape_id}, found {len(found)}")
    return found[0]


def put_text_frame(text_frame, label: str, new_text: str) -> None:
    paragraphs = text_frame.paragraphs
    lines = new_text.split("\n")
    if len(lines) > len(paragraphs):
        raise ValueError(
            f"{label} has {len(paragraphs)} paragraphs, but replacement has {len(lines)} lines"
        )
    for index, paragraph in enumerate(paragraphs):
        desired = lines[index] if index < len(lines) else ""
        text_nodes = paragraph._p.xpath(".//a:t")
        if not text_nodes and desired:
            paragraph.add_run().text = desired
            continue
        for run_index, node in enumerate(text_nodes):
            node.text = desired if run_index == 0 else ""


def put_text(shape, new_text: str) -> None:
    if not shape.has_text_frame:
        raise ValueError(f"Shape {shape.shape_id} has no text frame")
    put_text_frame(shape.text_frame, f"Shape {shape.shape_id}", new_text)


def put_table(shape, values: list[list[str]]) -> None:
    if not shape.has_table:
        raise ValueError(f"Shape {shape.shape_id} is not a table")
    table = shape.table
    if len(values) != len(table.rows) or any(len(row) != len(table.columns) for row in values):
        raise ValueError(f"Table {shape.shape_id} replacement dimensions differ from source")
    for row_index, row in enumerate(values):
        for col_index, value in enumerate(row):
            if not isinstance(value, str):
                raise ValueError(f"Table {shape.shape_id} cell {row_index},{col_index} must be text")
            put_text_frame(table.cell(row_index, col_index).text_frame, f"Table {shape.shape_id} cell {row_index},{col_index}", value)


def put_picture(slide, shape, image: Path) -> None:
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        raise ValueError(f"Shape {shape.shape_id} is not an ordinary picture")
    if not image.is_file():
        raise ValueError(f"Image not found: {image}")
    image_part, rid = slide.part.get_or_add_image_part(str(image))
    blip = shape._element.blipFill.blip
    blip.rEmbed = rid
    # The source image's a14:imgLayer references its original image. Keeping that
    # effect after replacing r:embed can make PowerPoint reject the saved file.
    extensions = blip.find(f"{{{A}}}extLst")
    if extensions is not None:
        for extension in list(extensions):
            if extension.find(f"{{{A14}}}imgProps") is not None:
                extensions.remove(extension)
        if not len(extensions):
            blip.remove(extensions)
    # crop_to_fit() appends a:srcRect. An existing template crop must be removed
    # first because two a:srcRect elements make the slide invalid in PowerPoint.
    blip_fill = shape._element.blipFill
    source_crop = blip_fill.find(f"{{{A}}}srcRect")
    if source_crop is not None:
        blip_fill.remove(source_crop)
    shape._element.crop_to_fit(image_part.image.size, (shape.width, shape.height))


def verify_native_powerpoint(path: Path, expected_slides: int, policy: str = "auto") -> bool:
    """Skip only unavailable hosts; never swallow an invalid presentation."""
    if policy not in {"auto", "required", "off"}:
        raise ValueError("native_validation must be auto, required, or off")
    if policy == "off":
        return False
    try:
        app = create_powerpoint()
    except NativePowerPointUnavailable:
        if policy == "auto":
            return False
        raise
    presentation = None
    try:
        presentation = app.Presentations.Open(str(path), True, False, False)
        if presentation.Slides.Count != expected_slides:
            raise ValueError("PowerPoint slide count differs from plan")
        return True
    finally:
        if presentation is not None:
            presentation.Close()
        app.Quit()


def set_transition(slide, effect: str) -> None:
    if effect not in TRANSITIONS:
        raise ValueError(f"Unsupported transition: {effect}")
    root = slide._element
    old = root.find(f"{{{P}}}transition")
    if old is not None:
        root.remove(old)
    transition = etree.Element(f"{{{P}}}transition", spd="med", advClick="1")
    kind, _, direction = effect.partition("-")
    child = etree.SubElement(transition, f"{{{P}}}{kind}")
    if direction:
        child.set("dir", "l" if direction == "left" else "r")
    timing = root.find(f"{{{P}}}timing")
    if timing is not None:
        root.insert(root.index(timing), transition)
    else:
        root.append(transition)


def strip_branding(deck, phrase: str) -> None:
    """Blank an exact source-template brand in slides, layouts, and masters."""
    containers = list(deck.slides)
    containers.extend(deck.slide_layouts)
    containers.extend(deck.slide_masters)
    for container in containers:
        for shape in nested_shapes(container.shapes):
            if shape.has_text_frame and phrase in shape.text:
                put_text(shape, shape.text.replace(phrase, ""))


def validate_plan(plan: dict, base: Path) -> tuple[Path, Path, list[dict]]:
    if plan.get("native_validation", "auto") not in {"auto", "required", "off"}:
        raise ValueError("native_validation must be auto, required, or off")
    template = resolve(base, plan["template"]).resolve()
    output = resolve(base, plan["output"]).resolve()
    if not template.is_file() or template.suffix.lower() != ".pptx":
        raise ValueError(f"PPTX template not found: {template}")
    if template == output:
        raise ValueError("Output must differ from source template")
    slides = plan["slides"]
    if not isinstance(slides, list) or not slides:
        raise ValueError("Plan requires at least one slide")
    numbers = [item["number"] for item in slides]
    if any(type(number) is not int or number < 1 for number in numbers) or len(numbers) != len(set(numbers)):
        raise ValueError("Slide numbers must be unique positive integers")
    return template, output, slides


def build(plan: dict, base: Path) -> dict:
    template, output, choices = validate_plan(plan, base)
    selected_numbers = {item["number"] for item in choices}
    source = inspect_file(template, include_slides=True, selected=selected_numbers)
    if max(item["number"] for item in choices) > source["slide_count"]:
        raise ValueError(f"Template has only {source['slide_count']} slides")
    source_slides = {slide["number"]: slide for slide in source["slides"]}
    require_animation = plan.get("require_animation", True)
    strict_text = plan.get("strict_text", True)
    brand_phrase = plan.get("strip_brand_text")
    if brand_phrase is not None and (not isinstance(brand_phrase, str) or not brand_phrase):
        raise ValueError("strip_brand_text must be a non-empty string")
    for choice in choices:
        original = source_slides[choice["number"]]
        if require_animation and not (original["animated"] or original["transition"]):
            raise ValueError(f"Source slide {choice['number']} has neither animation nor transition")
        replaced = {int(key) for key in choice.get("text", {})}
        kept = {int(key) for key in choice.get("keep_text", [])}
        original_ids = {shape["id"] for shape in original["shapes"]}
        unknown = (replaced | kept) - original_ids
        if unknown:
            raise ValueError(f"Slide {choice['number']} has no text shape IDs {sorted(unknown)}")
        if strict_text and original_ids != (replaced | kept):
            raise ValueError(
                f"Slide {choice['number']} has unreviewed text IDs {sorted(original_ids - replaced - kept)}"
            )
        table_ids = {table["id"] for table in original["tables"] if any(cell for row in table["cells"] for cell in row)}
        replaced_tables = {int(key) for key in choice.get("tables", {})}
        kept_tables = {int(key) for key in choice.get("keep_tables", [])}
        unknown_tables = (replaced_tables | kept_tables) - {table["id"] for table in original["tables"]}
        if unknown_tables:
            raise ValueError(f"Slide {choice['number']} has no table IDs {sorted(unknown_tables)}")
        if strict_text and table_ids != (replaced_tables | kept_tables):
            raise ValueError(f"Slide {choice['number']} has unreviewed table IDs {sorted(table_ids - replaced_tables - kept_tables)}")
        picture_ids = {picture["id"] for picture in original["pictures"]}
        unknown_pictures = {int(key) for key in choice.get("pictures", {})} - picture_ids
        if unknown_pictures:
            raise ValueError(f"Slide {choice['number']} has no picture IDs {sorted(unknown_pictures)}")
        if choice.get("transition") not in (None, *TRANSITIONS):
            raise ValueError(f"Unsupported transition on slide {choice['number']}")

    deck = Presentation(str(template))
    source_motion = {choice["number"]: motion_state(deck.slides[choice["number"] - 1]._element)
                     for choice in choices}
    slide_ids = deck.slides._sldIdLst
    original_ids = list(slide_ids)
    selected = [original_ids[choice["number"] - 1] for choice in choices]
    selected_set = set(selected)
    for slide_id in original_ids:
        if slide_id not in selected_set:
            deck.part.drop_rel(slide_id.rId)
            slide_ids.remove(slide_id)
    for slide_id in selected:
        slide_ids.remove(slide_id)
        slide_ids.append(slide_id)

    expected_motion = []
    for output_index, choice in enumerate(choices):
        slide = deck.slides[output_index]
        for raw_id, content in choice.get("text", {}).items():
            if not isinstance(content, str):
                raise ValueError(f"Slide {choice['number']} text {raw_id} must be a string")
            put_text(shape_by_id(slide, int(raw_id)), content)
        for raw_id, values in choice.get("tables", {}).items():
            put_table(shape_by_id(slide, int(raw_id)), values)
        for raw_id, image_name in choice.get("pictures", {}).items():
            put_picture(slide, shape_by_id(slide, int(raw_id)), resolve(base, image_name))
        if choice.get("transition"):
            set_transition(slide, choice["transition"])
        if "notes" in choice:
            if not isinstance(choice["notes"], str):
                raise ValueError(f"Slide {choice['number']} notes must be text")
            slide.notes_slide.notes_text_frame.text = choice["notes"]
        verify_motion(source_motion[choice["number"]], slide._element, output_index + 1,
                      compare_transition=choice.get("transition") is None)
        expected_motion.append(motion_state(slide._element))

    if brand_phrase:
        strip_branding(deck, brand_phrase)

    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(prefix="ppt-build-", suffix=".pptx", dir=output.parent, delete=False) as handle:
        temporary = Path(handle.name)
    try:
        deck.save(str(temporary))
        checked = inspect_file(temporary, include_slides=True)
        if checked["slide_count"] != len(choices):
            raise ValueError("Saved slide count differs from plan")
        for index, choice in enumerate(choices):
            before = source_slides[choice["number"]]
            after = checked["slides"][index]
            if before["animated"] != after["animated"] or (choice.get("transition") is None and before["transition"] != after["transition"]) or (choice.get("transition") is not None and not after["transition"]):
                raise ValueError(f"Animation or transition was lost on slide {index + 1}")
            actual_text = {shape["id"]: shape["text"] for shape in after["shapes"]}
            for raw_id, expected in choice.get("text", {}).items():
                if actual_text.get(int(raw_id)) != expected:
                    raise ValueError(f"Text verification failed for slide {index + 1}, shape {raw_id}")
            actual_tables = {table["id"]: table["cells"] for table in after["tables"]}
            for raw_id, expected in choice.get("tables", {}).items():
                if actual_tables.get(int(raw_id)) != expected:
                    raise ValueError(f"Table verification failed for slide {index + 1}, shape {raw_id}")
        reloaded = Presentation(str(temporary))
        if len(reloaded.slides) != len(choices):
            raise ValueError("PowerPoint package could not be reloaded")
        for index, choice in enumerate(choices):
            verify_motion(expected_motion[index], reloaded.slides[index]._element, index + 1)
            for raw_id, image_name in choice.get("pictures", {}).items():
                requested_image = resolve(base, image_name)
                actual_image = shape_by_id(reloaded.slides[index], int(raw_id)).image.blob
                if hashlib.sha256(actual_image).digest() != hashlib.sha256(requested_image.read_bytes()).digest():
                    raise ValueError(f"Picture verification failed for slide {index + 1}, shape {raw_id}")
        native_validated = verify_native_powerpoint(temporary, len(choices), plan.get("native_validation", "auto"))
        os.replace(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
    return {
        "output": str(output),
        "slides": len(choices),
        "animated_slides": checked["animated_count"],
        "transition_slides": checked["transition_count"],
        "native_validated": native_validated,
        "native_validation": plan.get("native_validation", "auto"),
        "bytes": output.stat().st_size,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("plan", type=Path, help="JSON plan with template, output, and slides")
    parser.add_argument("--native-validation", choices=("auto", "required", "off"),
                        help="Override the plan's optional native_validation policy")
    args = parser.parse_args()
    try:
        plan_path = args.plan.resolve()
        plan = json.loads(plan_path.read_text(encoding="utf-8"))
        if args.native_validation is not None:
            plan["native_validation"] = args.native_validation
        print(json.dumps(build(plan, plan_path.parent), ensure_ascii=False, indent=2))
        return 0
    except (OSError, KeyError, TypeError, ValueError, NativePowerPointUnavailable) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
