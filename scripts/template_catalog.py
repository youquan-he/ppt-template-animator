#!/usr/bin/env python
"""Index local PPTX templates and inspect editable slide text."""

from __future__ import annotations

import argparse
from collections import Counter
import json
import os
import posixpath
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
EFFECT_TAGS = {"anim", "animEffect", "animMotion", "animRot", "animScale", "set", "cmd"}


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def slide_metrics(xml: bytes, slide_size: tuple[int, int]) -> dict:
    root = ET.fromstring(xml)
    primitives = Counter(local_name(node.tag) for node in root.iter()
                         if node.tag.startswith(f"{{{P}}}") and local_name(node.tag) in EFFECT_TAGS)
    effects = sum(1 for node in root.iter(f"{{{P}}}cTn") if "presetID" in node.attrib)
    transitions = list(root.iter(f"{{{P}}}transition"))
    transition_type = None
    if transitions:
        kinds = [local_name(node.tag) for transition in transitions
                 for node in transition.iter() if node is not transition]
        transition_type = next((kind for kind in kinds if "morph" in kind.lower()), None)
        if transition_type is None:
            transition_type = next((kind for kind in kinds if kind != "extLst"), "unspecified")
    text_values = slide_text(xml)
    pictures = root.findall(f".//{{{P}}}pic")
    slide_area = max(1, slide_size[0] * slide_size[1])
    picture_areas = []
    for picture in pictures:
        ext = picture.find(f"./{{{P}}}spPr/{{{A}}}xfrm/{{{A}}}ext")
        if ext is not None:
            picture_areas.append(int(ext.attrib.get("cx", 0)) * int(ext.attrib.get("cy", 0)) / slide_area)
    return {
        "effect_count": effects,
        "animation_primitives": dict(primitives),
        "transition_type": transition_type,
        "text_chars": sum(len(item["text"]) for item in text_values),
        "text_shape_count": len(text_values),
        "picture_count": len(pictures),
        "largest_picture_area_fraction": round(min(1.0, max(picture_areas, default=0.0)), 3),
    }


def slide_entries(archive: zipfile.ZipFile) -> list[tuple[int, str]]:
    presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
    relationships = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall(f"{{{PKG_REL}}}Relationship")
    }
    entries = []
    for number, slide in enumerate(presentation.findall(f"./{{{P}}}sldIdLst/{{{P}}}sldId"), 1):
        target = targets[slide.attrib[f"{{{R}}}id"]]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("ppt", target))
        entries.append((number, part))
    return entries


def slide_text(xml: bytes) -> list[dict]:
    root = ET.fromstring(xml)
    result = []
    for shape in root.iter(f"{{{P}}}sp"):
        props = shape.find(f"./{{{P}}}nvSpPr/{{{P}}}cNvPr")
        if props is None:
            continue
        paragraphs = ["".join(node.text or "" for node in para.iter(f"{{{A}}}t")) for para in shape.iter(f"{{{A}}}p")]
        content = "\n".join(paragraphs).strip()
        if content:
            result.append({"id": int(props.attrib["id"]), "name": props.attrib.get("name", ""), "text": content, "paragraphs": len(paragraphs)})
    return result


def slide_pictures(xml: bytes) -> list[dict]:
    root = ET.fromstring(xml)
    result = []
    for picture in root.iter(f"{{{P}}}pic"):
        props = picture.find(f"./{{{P}}}nvPicPr/{{{P}}}cNvPr")
        if props is not None:
            result.append({"id": int(props.attrib["id"]), "name": props.attrib.get("name", "")})
    return result


def slide_tables(xml: bytes) -> list[dict]:
    root = ET.fromstring(xml)
    result = []
    for frame in root.iter(f"{{{P}}}graphicFrame"):
        table = frame.find(f".//{{{A}}}tbl")
        props = frame.find(f"./{{{P}}}nvGraphicFramePr/{{{P}}}cNvPr")
        if table is None or props is None:
            continue
        rows = []
        for row in table.findall(f"{{{A}}}tr"):
            rows.append(["".join(node.text or "" for node in cell.iter(f"{{{A}}}t")) for cell in row.findall(f"{{{A}}}tc")])
        result.append({"id": int(props.attrib["id"]), "name": props.attrib.get("name", ""), "cells": rows})
    return result


def inspect_file(path: Path, include_slides: bool = False, selected: set[int] | None = None) -> dict:
    result = {"path": str(path), "bytes": path.stat().st_size}
    with zipfile.ZipFile(path) as archive:
        slides = slide_entries(archive)
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        size = presentation.find(f"{{{P}}}sldSz")
        slide_size = (int(size.attrib["cx"]), int(size.attrib["cy"])) if size is not None else (1, 1)
        animated = []
        transitions = []
        effect_counts = []
        transition_types = Counter()
        large_picture_slides = []
        details = []
        for number, name in slides:
            xml = archive.read(name)
            metrics = slide_metrics(xml, slide_size)
            has_animation = metrics["effect_count"] > 0 or bool(metrics["animation_primitives"])
            has_transition = metrics["transition_type"] is not None
            effect_counts.append(metrics["effect_count"])
            if has_animation:
                animated.append(number)
            if has_transition:
                transitions.append(number)
                transition_types[metrics["transition_type"]] += 1
            if metrics["largest_picture_area_fraction"] >= 0.30:
                large_picture_slides.append(number)
            if include_slides and (selected is None or number in selected):
                details.append({
                    "number": number,
                    "part": name,
                    "animated": has_animation,
                    "transition": has_transition,
                    **metrics,
                    "shapes": slide_text(xml),
                    "pictures": slide_pictures(xml),
                    "tables": slide_tables(xml),
                })
        result.update({
            "slide_count": len(slides),
            "animated_count": len(animated),
            "transition_count": len(transitions),
            "effect_count": sum(effect_counts),
            "transition_types": dict(transition_types),
            "large_picture_count": len(large_picture_slides),
            "animated_examples": animated[:12],
            "transition_examples": transitions[:12],
            "large_picture_examples": large_picture_slides[:12],
        })
        if include_slides:
            result["slides"] = details
        elif slides:
            result["first_slide_text"] = [item["text"][:80] for item in slide_text(archive.read(slides[0][1]))[:8]]
    return result


def roots_from_args(values: list[str] | None) -> list[Path]:
    if values:
        raw = values
    elif os.environ.get("PPT_TEMPLATE_ROOTS"):
        raw = [value.strip() for value in os.environ["PPT_TEMPLATE_ROOTS"].split(os.pathsep) if value.strip()]
    else:
        raw = ["templates"] if Path("templates").is_dir() else []
    if not raw:
        raise ValueError("No template directory configured. Pass --roots, set PPT_TEMPLATE_ROOTS, or create ./templates")
    roots = [Path(value).expanduser() for value in raw]
    missing = [str(root) for root in roots if not root.is_dir()]
    if missing:
        raise ValueError(f"Template directories not found: {missing}")
    return roots


def catalog(roots: list[Path], limit: int | None) -> dict:
    paths = sorted({path.resolve() for root in roots for path in root.rglob("*.pptx") if not path.name.startswith("~$")})
    if limit is not None:
        paths = paths[:limit]
    records, errors = [], []
    for index, path in enumerate(paths, 1):
        try:
            records.append(inspect_file(path))
        except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
            errors.append({"path": str(path), "error": str(exc)})
        if index % 100 == 0:
            print(f"Indexed {index}/{len(paths)}", file=sys.stderr)
    return {"roots": [str(root) for root in roots], "files": records, "errors": errors}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    scan = sub.add_parser("scan", help="Index template directories")
    scan.add_argument("--roots", nargs="+", help="Template directories; otherwise use PPT_TEMPLATE_ROOTS or ./templates")
    scan.add_argument("--output", required=True, type=Path, help="Write catalog JSON here")
    scan.add_argument("--limit", type=int, help="Index only this many files for a quick check")
    inspect = sub.add_parser("inspect", help="List slide text shape IDs and animation flags")
    inspect.add_argument("template", type=Path)
    inspect.add_argument("--output", type=Path, help="Write detailed JSON here")
    inspect.add_argument("--slides", help="Only print these 1-based slide numbers, comma separated")
    args = parser.parse_args()
    try:
        requested = {int(value) for value in args.slides.split(",")} if args.command == "inspect" and args.slides else None
        data = catalog(roots_from_args(args.roots), args.limit) if args.command == "scan" else inspect_file(args.template, True, requested)
        if requested:
            if any(number < 1 or number > data["slide_count"] for number in requested):
                raise ValueError(f"Slide selection must be between 1 and {data['slide_count']}")
        serialized = json.dumps(data, ensure_ascii=False, indent=2)
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(serialized + "\n", encoding="utf-8")
        else:
            print(serialized)
        if args.command == "scan":
            print(f"Indexed {len(data['files'])} PPTX files; {len(data['errors'])} errors", file=sys.stderr)
        return 0
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
