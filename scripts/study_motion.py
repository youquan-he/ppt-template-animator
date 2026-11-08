#!/usr/bin/env python
"""Inspect source motion and adjacent-slide dependencies without modifying a PPTX.

The report is research evidence, not an animation migration or visual-quality verdict.
Timing/behavior values remain raw OOXML values; no external relationship is fetched.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import json
import posixpath
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET
import zipfile

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
KNOWN_MODERN = {P, A, "http://schemas.microsoft.com/office/powerpoint/2010/main",
                "http://schemas.microsoft.com/office/powerpoint/2012/main",
                "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
                "http://schemas.microsoft.com/office/drawing/2010/main"}
BEHAVIORS = {"anim", "animEffect", "animMotion", "animRot", "animScale", "set", "cmd", "audio", "video"}
SHAPES = {"sp", "pic", "grpSp", "graphicFrame", "cxnSp", "contentPart"}


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def tree(node: ET.Element) -> dict:
    """Use expanded tags so prefix changes do not alter the reported semantics."""
    result = {"tag": node.tag, "attributes": dict(node.attrib)}
    if node.text and node.text.strip():
        result["text"] = node.text.strip()
    if len(node):
        result["children"] = [tree(child) for child in node]
    return result


def compatibility_tree(xml: bytes) -> tuple[ET.Element, ET.Element, list[dict]]:
    """Resolve known-modern Choice branches, without counting Fallback twice.

    This is a declared schema policy, not a probe of the actual playback host.
    Unknown Requires values remain explicitly unresolved even when a fallback exists.
    """
    contexts, stack, pending = {}, [], []
    parser = ET.iterparse(io.BytesIO(xml), events=("start-ns", "start", "end"))
    for event, value in parser:
        if event == "start-ns":
            pending.append(value)
        elif event == "start":
            context = dict(stack[-1]) if stack else {}
            context.update(pending)
            pending.clear()
            contexts[value] = context
            stack.append(context)
        else:
            stack.pop()
    raw, branches = parser.root, []

    def clone(node, location):
        if node.tag == f"{{{MC}}}AlternateContent":
            choices = [child for child in node if child.tag == f"{{{MC}}}Choice"]
            selected = None
            requirements = []
            for choice in choices:
                uris = [contexts[choice].get(prefix) for prefix in choice.get("Requires", "").split()]
                requirements.append(uris)
                if uris and all(uri in KNOWN_MODERN for uri in uris):
                    selected = choice
                    break
            resolved = selected is not None
            if selected is None:
                selected = next((child for child in node if child.tag == f"{{{MC}}}Fallback"), None)
            branches.append({"location": location, "selected": local(selected.tag) if selected is not None else None,
                             "status": "known_modern_choice" if resolved else "unresolved_host_support",
                             "choice_requirement_namespaces": requirements})
            return [copy for child in selected for copy in clone(child, location + "/" + local(selected.tag))] if selected is not None else []
        result = ET.Element(node.tag, dict(node.attrib))
        result.text, result.tail = node.text, node.tail
        for index, child in enumerate(node):
            result.extend(clone(child, location + "/" + local(child.tag) + f"[{index}]"))
        return [result]

    return raw, clone(raw, "/sld")[0], branches


def resolve_part(base: str, target: str) -> str:
    uri = urlsplit(target)
    if uri.scheme or uri.netloc or uri.query or "\\" in target:
        raise ValueError("An internal relationship must be a package-relative URI.")
    value = unquote(uri.path)
    part = posixpath.normpath(value.lstrip("/") if value.startswith("/")
                             else posixpath.join(posixpath.dirname(base), value))
    if part in (".", "..") or part.startswith("../") or ":" in part or "\x00" in part:
        raise ValueError("An internal relationship escapes the PPTX package.")
    return part


def relationships(archive: zipfile.ZipFile, part: str) -> dict[str, dict]:
    relpart = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if relpart not in archive.namelist():
        return {}
    result = {}
    for rel in ET.fromstring(archive.read(relpart)).findall(f"{{{REL}}}Relationship"):
        item = dict(rel.attrib)
        item["external"] = item.get("TargetMode") == "External"
        if not item["external"]:
            item["part"] = resolve_part(part, item["Target"])
        if item["Id"] in result:
            raise ValueError("Duplicate relationship ID in package.")
        result[item["Id"]] = item
    return result


def slide_parts(archive: zipfile.ZipFile) -> list[str]:
    rels = relationships(archive, "ppt/presentation.xml")
    root = ET.fromstring(archive.read("ppt/presentation.xml"))
    parts = []
    for entry in root.findall(f"./{{{P}}}sldIdLst/{{{P}}}sldId"):
        rel = rels[entry.attrib[f"{{{R}}}id"]]
        if rel["external"] or not rel.get("Type", "").endswith("/slide"):
            raise ValueError("A presentation slide must reference an internal slide part.")
        parts.append(rel["part"])
    return parts


def select_slides(value: str | None, count: int) -> list[int]:
    if count < 1:
        raise ValueError("The presentation contains no slides.")
    if value is None:
        return list(range(1, count + 1))
    selected = set()
    for token in value.split(","):
        match = re.fullmatch(r"\s*([1-9][0-9]*)(?:\s*-\s*([1-9][0-9]*))?\s*", token)
        if not match:
            raise ValueError("--slides must contain slide numbers or ascending ranges, such as 1-4,7.")
        first, last = int(match[1]), int(match[2] or match[1])
        if first > last or last > count:
            raise ValueError(f"Slide selection must be ascending and within 1-{count}.")
        selected.update(range(first, last + 1))
    return sorted(selected)


def shape_inventory(root: ET.Element) -> list[dict]:
    result = []

    def visit(container: ET.Element, parent: str | None = None, top: str | None = None):
        for shape in container:
            kind = local(shape.tag)
            if shape.tag != f"{{{P}}}{kind}" or kind not in SHAPES:
                continue
            props = next((node for child in shape if local(child.tag).startswith("nv")
                          for node in child if node.tag == f"{{{P}}}cNvPr"), None)
            sid = props.get("id") if props is not None else None
            transforms = [tree(x) for child in shape if local(child.tag) in ("spPr", "grpSpPr")
                          for x in child if local(x.tag) == "xfrm"]
            transforms += [tree(x) for x in shape if x.tag == f"{{{P}}}xfrm"]
            text = "\n".join("".join(t.text or "" for t in para.iter(f"{{{A}}}t"))
                             for body in shape.findall(f"{{{P}}}txBody") for para in body.findall(f"{{{A}}}p"))
            result.append({"id": sid, "name": props.get("name", "") if props is not None else "",
                           "kind": kind, "parent_group_id": parent, "top_level_id": top or sid,
                           "geometry": transforms, "geometry_space": "parent_group" if parent else "slide",
                           "text": text, "text_summary": text[:300]})
            if kind == "grpSp":
                visit(shape, sid, top or sid)

    container = root.find(f"./{{{P}}}cSld/{{{P}}}spTree")
    if container is not None:
        visit(container)
    return result


def transition_info(root: ET.Element) -> dict:
    # Modern Morph may live in extension/AlternateContent namespaces.
    transitions = [node for node in root.iter() if local(node.tag) == "transition"]
    morph = any("morph" in local(node.tag).lower() for node in root.iter())
    kinds = [local(child.tag) for node in transitions for child in node
             if local(child.tag) not in ("extLst", "AlternateContent")]
    return {"morph": morph, "types": ["morph"] if morph else kinds,
            "raw": [tree(node) for node in transitions],
            "morph_nodes": [tree(node) for node in root.iter() if "morph" in local(node.tag).lower()]}


def motion_info(root: ET.Element, shapes: list[dict]) -> dict:
    parents = {child: node for node in root.iter() for child in node}
    ids = {shape["id"]: shape for shape in shapes}
    timing_nodes, effects, interactive = [], [], []
    for node in root.iter(f"{{{P}}}cTn"):
        conditions = [tree(child) for child in node if local(child.tag) in
                      ("stCondLst", "endCondLst", "nextCondLst", "prevCondLst", "endSync")]
        record = {"attributes": dict(node.attrib), "conditions": conditions}
        timing_nodes.append(record)
        if node.get("nodeType") == "interactiveSeq":
            interactive.append(tree(node))
    for node in root.iter():
        if node.tag not in {f"{{{P}}}{kind}" for kind in BEHAVIORS}:
            continue
        ancestor, lineage = parents.get(node), []
        while ancestor is not None:
            if ancestor.tag == f"{{{P}}}cTn":
                lineage.append(dict(ancestor.attrib))
            ancestor = parents.get(ancestor)
        targets = []
        for target in node.iter(f"{{{P}}}spTgt"):
            item = {"attributes": dict(target.attrib), "raw": tree(target)}
            shape = ids.get(target.get("spid"))
            item["shape"] = {key: shape[key] for key in ("id", "name", "parent_group_id", "top_level_id")} if shape else None
            targets.append(item)
        effects.append({"kind": local(node.tag), "parameters": dict(node.attrib),
                        "ancestor_timing": list(reversed(lineage)), "targets": targets, "raw": tree(node)})
    preset_groups = [record for record in timing_nodes if any(k in record["attributes"] for k in
                     ("presetClass", "presetID", "presetSubtype"))]
    return {"timing_nodes": timing_nodes, "preset_groups": preset_groups, "behaviors": effects,
            "interactive_sequences": interactive, "interactive_sequence_count": len(interactive),
            "unresolved_target_ids": sorted({target["attributes"].get("spid", "") for effect in effects
                                             for target in effect["targets"] if target["shape"] is None})}


def study(path: Path, selection: str | None = None) -> dict:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while block := stream.read(4 * 1024**2):
            digest.update(block)
    with zipfile.ZipFile(path) as archive:
        parts = slide_parts(archive)
        selected = select_slides(selection, len(parts))
        parsed = [compatibility_tree(archive.read(part)) for part in parts]
        roots = [item[1] for item in parsed]
        shapes = [shape_inventory(root) for root in roots]
        transitions = [transition_info(root) for root in roots]
        reports = []
        for number in selected:
            index = number - 1
            neighbors = {"previous": number - 1 if index else None,
                         "next": number + 1 if number < len(parts) else None}
            dependencies, warnings = [], []
            for destination in (number, number + 1):
                if not 1 < destination <= len(parts) or not transitions[destination - 1]["morph"]:
                    continue
                before, after = destination - 1, destination
                left = {s["name"] for s in shapes[before - 1] if s["name"]}
                right = {s["name"] for s in shapes[after - 1] if s["name"]}
                dependencies.append({"kind": "adjacent_morph", "from_slide": before, "to_slide": after,
                                     "both_selected": before in selected and after in selected,
                                     "shared_shape_names": sorted(left & right),
                                     "forced_match_names": sorted(name for name in left & right if name.startswith("!!"))})
                if before not in selected or after not in selected:
                    warnings.append(f"Morph pair {before}-{after} is not fully selected; study and retain the neighboring state before reuse.")
            hints = []
            for rel in relationships(archive, parts[index]).values():
                kind = rel.get("Type", "").rsplit("/", 1)[-1]
                if rel["external"] or kind in ("audio", "video", "media"):
                    hints.append({"id": rel["Id"], "type": rel.get("Type"), "target": rel["Target"],
                                  "external": rel["external"], "package_part": rel.get("part"),
                                  "part_exists": rel.get("part") in archive.namelist() if not rel["external"] else None})
            motion = motion_info(roots[index], shapes[index])
            raw_root, _, branches = parsed[index]
            if any(branch["status"] == "unresolved_host_support" for branch in branches):
                warnings.append("Some compatibility choices require unknown host support; effective branches are unresolved, not a playback claim.")
            if motion["interactive_sequence_count"]:
                warnings.append("Interactive sequences exist; preserve trigger objects and the full timing context before rebinding.")
            if any(h["external"] for h in hints):
                warnings.append("External relationships are reported only and were not fetched.")
            texts = [shape["text_summary"] for shape in shapes[index] if shape["text_summary"]]
            reports.append({"number": number, "part": parts[index], "title_or_first_text": texts[0] if texts else "",
                            "text_summary": texts, "neighbors": neighbors, "shapes": shapes[index],
                            "transition": transitions[index], "motion": motion, "dependencies": dependencies,
                            "compatibility": {"policy": "known_modern_choice_not_host_probe", "branches": branches,
                                              "raw_xml_behavior_counts": {kind: sum(1 for _ in raw_root.iter(f"{{{P}}}{kind}"))
                                                                          for kind in sorted(BEHAVIORS)}},
                            "relationship_hints": hints, "warnings": warnings})
    return {"format_version": 1, "status": "source_study_only", "source": {"name": path.name,
            "sha256": digest.hexdigest(), "slide_count": len(parts)}, "selected_slides": selected,
            "limits": ["Raw OOXML values do not prove playback quality, appearance or successful migration.",
                       "Geometry is reported in its original coordinate space; group transforms are not flattened.",
                       "Preset IDs are recorded, not translated into assumed animation names."], "slides": reports}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--slides", help="1-based slide numbers/ranges, e.g. 1-4,7; default: all slides")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        if args.output.suffix.lower() != ".json" or args.output.resolve() == args.source.resolve():
            raise ValueError("--output must be a separate JSON file.")
        report = study(args.source, args.slides)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        print(json.dumps({"status": report["status"], "source": report["source"]["name"],
                          "selected_slides": report["selected_slides"], "output": str(args.output)}, ensure_ascii=False))
        return 0
    except (OSError, ValueError, KeyError, ET.ParseError, zipfile.BadZipFile) as exc:
        print(f"Motion study failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
