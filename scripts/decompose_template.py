"""Inspect template shapes and extract editable components with native animation."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import tempfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from native_powerpoint import NativePowerPointUnavailable, create_powerpoint


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def direct_style(element) -> dict:
    """Report declared styles, without pretending to resolve inherited values."""
    fonts, sizes, colors, theme_colors = set(), set(), set(), set()
    for node in element.iter():
        if node.tag in {f"{{{A}}}latin", f"{{{A}}}ea", f"{{{A}}}cs"} and node.get("typeface"):
            fonts.add(node.get("typeface"))
        if node.tag in {f"{{{A}}}rPr", f"{{{A}}}defRPr", f"{{{A}}}endParaRPr"} and node.get("sz"):
            sizes.add(int(node.get("sz")) / 100)
        if node.tag == f"{{{A}}}srgbClr" and node.get("val"):
            colors.add(node.get("val"))
        if node.tag == f"{{{A}}}schemeClr" and node.get("val"):
            theme_colors.add(node.get("val"))
    return {"font_names": sorted(fonts), "font_sizes_pt": sorted(sizes),
            "rgb_colors": sorted(colors), "theme_color_refs": sorted(theme_colors)}


def design_system(slide) -> dict:
    layout = slide.slide_layout
    master = layout.slide_master
    report = {"layout_part": str(layout.part.partname), "master_part": str(master.part.partname),
              "slide_direct_style": direct_style(slide._element), "theme": None,
              "theme_override_parts": [],
              "interpretation": "Declared styles and the associated master theme only; inherited text styling, color maps, overrides, and color transforms still require rendered verification."}
    for part in (slide.part, layout.part, master.part):
        for relationship in part.rels.values():
            if not relationship.is_external and relationship.reltype.endswith("/themeOverride"):
                report["theme_override_parts"].append(str(relationship.target_part.partname))
    for relationship in master.part.rels.values():
        if relationship.is_external or not relationship.reltype.endswith("/theme"):
            continue
        part = relationship.target_part
        theme = etree.fromstring(part.blob)
        palette = {}
        scheme = theme.find(f"./{{{A}}}themeElements/{{{A}}}clrScheme")
        if scheme is not None:
            for color in scheme:
                value = next(iter(color), None)
                if value is not None:
                    palette[etree.QName(color).localname] = {
                        "type": etree.QName(value).localname,
                        "value": value.get("val"), "last_color": value.get("lastClr")}
        fonts = {}
        for role in ("major", "minor"):
            node = theme.find(f"./{{{A}}}themeElements/{{{A}}}fontScheme/{{{A}}}{role}Font")
            if node is not None:
                fonts[role] = {"latin": "", "east_asian": "", "complex_script": "", "script_fonts": {}}
                for tag, label in (("latin", "latin"), ("ea", "east_asian"), ("cs", "complex_script")):
                    font = node.find(f"{{{A}}}{tag}")
                    if font is not None:
                        fonts[role][label] = font.get("typeface", "")
                fonts[role]["script_fonts"] = {font.get("script"): font.get("typeface", "")
                                                for font in node.findall(f"{{{A}}}font") if font.get("script")}
        report["theme"] = {"part": str(part.partname), "name": theme.get("name"),
                           "palette": palette, "fonts": fonts}
        break
    return report


def shape_record(shape) -> dict:
    result = {
        "id": shape.shape_id, "name": shape.name, "type": str(shape.shape_type),
        "bounds_emu": [shape.left, shape.top, shape.width, shape.height],
        "text": shape.text if shape.has_text_frame else None,
        "direct_style": direct_style(shape._element),
    }
    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        result["children"] = [shape_record(child) for child in shape.shapes]
    return result


def inspect(source: Path, number: int) -> dict:
    deck = Presentation(str(source))
    if not 1 <= number <= len(deck.slides):
        raise ValueError(f"Slide must be between 1 and {len(deck.slides)}")
    slide = deck.slides[number - 1]
    return {"source": str(source.resolve()), "slide": number,
            "slide_size_emu": [deck.slide_width, deck.slide_height],
            "design_system": design_system(slide),
            "shapes": [shape_record(shape) for shape in slide.shapes]}


def descendant_ids(records: list[dict]) -> set[int]:
    result = set()
    for record in records:
        result.add(record["id"])
        result.update(descendant_ids(record.get("children", [])))
    return result


def save_media(slide, shapes: list, directory: Path) -> list[dict]:
    directory.mkdir()
    assets = {}
    for shape in shapes:
        for element in shape._element.iter():
            rid = element.get(f"{{{R}}}embed")
            if not rid:
                continue
            relationship = slide.part.rels[rid]
            if relationship.is_external or not relationship.reltype.endswith("/image"):
                continue
            part = relationship.target_part
            digest = hashlib.sha256(part.blob).hexdigest()
            if digest not in assets:
                name = f"{digest}{Path(str(part.partname)).suffix}"
                (directory / name).write_bytes(part.blob)
                assets[digest] = {"file": f"media/{name}", "sha256": digest,
                                  "source_part": str(part.partname), "references": []}
            reference = {"top_level_shape_id": shape.shape_id, "relationship_id": rid}
            if reference not in assets[digest]["references"]:
                assets[digest]["references"].append(reference)
    return list(assets.values())


def native_state(slide) -> dict:
    import pywintypes

    effects = []
    sequence = slide.TimeLine.MainSequence
    for number in range(1, sequence.Count + 1):
        effect = sequence.Item(number)
        timing = effect.Timing
        record = {
            "order": number, "shape_id": int(effect.Shape.Id),
            "effect_type": int(effect.EffectType), "exit": bool(effect.Exit),
            "trigger_type": int(timing.TriggerType),
            "duration": round(float(timing.Duration), 4),
            "delay": round(float(timing.TriggerDelayTime), 4),
            "accelerate": round(float(timing.Accelerate), 4),
            "decelerate": round(float(timing.Decelerate), 4),
        }
        for label, read in (("effect_name", lambda: effect.DisplayName),
                            ("direction", lambda: int(effect.EffectParameters.Direction))):
            try:
                record[label] = read()
            except (AttributeError, pywintypes.com_error):
                record[label] = None
        effects.append(record)
    return {
        "shape_ids": sorted(int(slide.Shapes.Item(i).Id) for i in range(1, slide.Shapes.Count + 1)),
        "main_sequence": effects,
        "interactive_sequence_count": int(slide.TimeLine.InteractiveSequences.Count),
        "transition": int(slide.SlideShowTransition.EntryEffect),
    }


def extract(source: Path, number: int, shape_ids: list[int] | None, output: Path) -> dict:
    source, output = source.resolve(), output.resolve()
    if output.exists():
        raise ValueError("Use a new output directory; existing component assets are not overwritten")
    inventory = inspect(source, number)
    known = {item["id"] for item in inventory["shapes"]}
    chosen = known if shape_ids is None else set(shape_ids)
    if not chosen or chosen - known:
        raise ValueError(f"Select top-level IDs from {sorted(known)}; select a group to retain its children")
    if shape_ids is not None and len(shape_ids) != len(chosen):
        raise ValueError("Shape IDs must be unique")
    records = [item for item in inventory["shapes"] if item["id"] in chosen]
    retained_ids = descendant_ids(records)
    deck = Presentation(str(source))
    slide = deck.slides[number - 1]
    timing = slide._element.find(f"{{{P}}}timing")
    output.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="template-components-", dir=output.parent) as temporary:
        staging = Path(temporary)
        assets = save_media(slide, [s for s in slide.shapes if s.shape_id in chosen], staging / "media")
        if timing is not None:
            (staging / "source-timing.xml").write_bytes(etree.tostring(timing, encoding="utf-8", pretty_print=True))
        # Keep the native slide container and its relationships. PowerPoint then
        # removes unselected shapes and cleans their animation references.
        slide_ids = deck.slides._sldIdLst
        for index in reversed(range(len(slide_ids))):
            if index != number - 1:
                removed = slide_ids[index]
                deck.part.drop_rel(removed.rId)
                slide_ids.remove(removed)
        component_file = staging / "components.pptx"
        deck.save(str(component_file))

        app = create_powerpoint()
        presentation = None
        try:
            presentation = app.Presentations.Open(str(component_file), False, False, False)
            native_slide = presentation.Slides.Item(1)
            before = native_state(native_slide)
            if before["interactive_sequence_count"] and chosen != known:
                raise ValueError("Partial extraction of interactive triggers needs their full context; retain the whole slide or rebuild triggers")
            height = round(1600 * presentation.PageSetup.SlideHeight / presentation.PageSetup.SlideWidth)
            native_slide.Export(str(staging / "source-slide.png"), "PNG", 1600, height)
            for index in range(native_slide.Shapes.Count, 0, -1):
                shape = native_slide.Shapes.Item(index)
                if int(shape.Id) not in chosen:
                    shape.Delete()
            retained = native_state(native_slide)
            expected_effects = [(e["shape_id"], e["effect_type"]) for e in before["main_sequence"]
                                if e["shape_id"] in retained_ids]
            actual_effects = [(e["shape_id"], e["effect_type"]) for e in retained["main_sequence"]]
            if actual_effects != expected_effects:
                raise ValueError("Retained animation targets or effects changed during extraction")
            if retained["shape_ids"] != sorted(chosen):
                raise ValueError("Extracted shapes differ from selection")
            native_slide.Export(str(staging / "components.png"), "PNG", 1600, height)
            presentation.Save()
            presentation.Close()
            presentation = None
            presentation = app.Presentations.Open(str(component_file), True, False, False)
            if presentation.Slides.Count != 1 or native_state(presentation.Slides.Item(1)) != retained:
                raise ValueError("Component shapes or animation changed after native save")
        finally:
            if presentation is not None:
                presentation.Saved = True
                presentation.Close()
            app.Quit()

        reloaded = Presentation(str(component_file))
        shapes = [shape_record(s) for s in reloaded.slides[0].shapes]
        if descendant_ids(shapes) != retained_ids:
            raise ValueError("Group child objects changed after native save")
        report = {
            "source": str(source), "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "source_slide": number, "selected_shape_ids": sorted(chosen), "shapes": shapes,
            "source_design_system": inventory["design_system"],
            "media": assets, "source_motion": before, "retained_motion": retained,
            "native_validated": True,
            "notes": ["Source masters/background are retained; copy desired objects into the new composition.",
                      "Unselected background shapes may be removed; review contrast and recolor components for the target theme.",
                      "A single extracted slide cannot verify Morph continuity with its preceding slide.",
                      "Retained effects still require timing and trigger review in the target scene."],
        }
        (staging / "manifest.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        os.replace(staging, output)
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    read = sub.add_parser("inspect", help="List all top-level shapes and group children")
    read.add_argument("source", type=Path)
    read.add_argument("--slide", required=True, type=int)
    read.add_argument("--output", type=Path)
    split = sub.add_parser("extract", help="Export selected native components, media, and motion")
    split.add_argument("source", type=Path)
    split.add_argument("--slide", required=True, type=int)
    split.add_argument("--shape-ids", help="Comma-separated top-level IDs; omit to retain all")
    split.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    if args.command == "inspect":
        report = inspect(args.source, args.slide)
        content = json.dumps(report, ensure_ascii=False, indent=2)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(content + "\n", encoding="utf-8")
        else:
            print(content)
    else:
        ids = [int(value) for value in args.shape_ids.split(",")] if args.shape_ids else None
        try:
            report = extract(args.source, args.slide, ids, args.output_dir)
        except NativePowerPointUnavailable as exc:
            parser.exit(2, f"Error: {exc}\n")
        print(json.dumps({"output_dir": str(args.output_dir.resolve()),
                          "selected_shapes": len(report["selected_shape_ids"]),
                          "retained_effects": len(report["retained_motion"]["main_sequence"]),
                          "media_files": len(report["media"]), "native_validated": True}, ensure_ascii=False))


if __name__ == "__main__":
    main()
