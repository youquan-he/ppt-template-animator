"""Apply native PowerPoint animations and transitions from a small JSON plan."""

from __future__ import annotations

import argparse
import json
import math
import os
from pathlib import Path
import re
import tempfile

from native_powerpoint import NativePowerPointUnavailable, create_powerpoint, powerpoint_constants


def finite_number(value, minimum: float, maximum: float) -> bool:
    return type(value) in (int, float) and minimum <= value <= maximum and math.isfinite(value)


def motion_path_tokens(path: str) -> tuple:
    """Validate a bounded normalized-slide VML M/L/C/Z/E path, preserving raw input."""
    if not isinstance(path, str) or not path or len(path) > 8192:
        raise ValueError("Motion path must be a nonempty VML string of at most 8192 characters")
    token = re.compile(r"[MLCZE]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?")
    tokens, end = [], 0
    for match in token.finditer(path):
        if path[end:match.start()].strip(" \t\r\n,"):
            raise ValueError("Motion path contains unsupported VML syntax")
        value = match.group()
        if value in {"M", "L", "C", "Z", "E"}:
            tokens.append(value)
        else:
            number = float(value)
            if not finite_number(number, -4, 4):
                raise ValueError("Motion path coordinates must be finite normalized-slide values between -4 and 4")
            tokens.append(number)
        end = match.end()
    if path[end:].strip(" \t\r\n,") or not tokens or tokens[0] != "M" or tokens[-1] != "E":
        raise ValueError("Motion path must start with M and finish with E")
    position, opened, drawn = 0, False, False
    while position < len(tokens):
        command = tokens[position]
        if not isinstance(command, str):
            raise ValueError("Motion path numbers must follow a supported command")
        finish = position + 1
        while finish < len(tokens) and not isinstance(tokens[finish], str):
            finish += 1
        count = finish - position - 1
        if command == "M":
            if count != 2:
                raise ValueError("VML M requires exactly two coordinates")
            opened = True
        elif command in {"L", "C"}:
            width = 2 if command == "L" else 6
            if not opened or count < width or count % width:
                raise ValueError("VML L/C requires an open subpath and complete coordinate groups")
            drawn = True
        elif command == "Z":
            if count or not opened:
                raise ValueError("VML Z closes an open subpath and takes no coordinates")
            opened = False
        elif count or finish != len(tokens) or not drawn:
            raise ValueError("VML E must finish a path containing at least one L/C segment")
        position = finish
    return tuple(tokens)


def same_motion_path(first: str, second: str) -> bool:
    left, right = motion_path_tokens(first), motion_path_tokens(second)
    return len(left) == len(right) and all(a == b if isinstance(a, str) or isinstance(b, str)
                                          else abs(a - b) <= 0.00002 for a, b in zip(left, right))


def normalize_effect(spec: list | dict, slide_number: int) -> dict:
    if isinstance(spec, list) and len(spec) == 4:
        spec = dict(zip(("shape", "effect", "trigger", "duration"), spec))
    if not isinstance(spec, dict):
        raise ValueError(f"Slide {slide_number} effect must be a four-item list or an object")
    allowed = {"shape", "effect", "trigger", "duration", "delay", "direction",
               "accelerate", "decelerate", "path", "scale", "rotate", "start_visible"}
    if set(spec) - allowed or not {"shape", "effect", "trigger", "duration"} <= set(spec):
        raise ValueError(f"Slide {slide_number} effect has missing or unsupported fields")
    if not isinstance(spec["shape"], str) or not spec["shape"]:
        raise ValueError(f"Slide {slide_number} effect shape must be a name")
    if spec["effect"] not in {"fade", "wipe", "zoom", "fly", "transform"} or spec["trigger"] not in {"click", "after", "with"}:
        raise ValueError(f"Slide {slide_number} has unsupported effect or trigger")
    if not finite_number(spec["duration"], 0.1, 30):
        raise ValueError(f"Slide {slide_number} effect duration must be 0.1–30 seconds")
    if "direction" in spec and (spec["effect"] not in {"fly", "wipe"} or
                                spec["direction"] not in {"left", "right", "up", "down"}):
        raise ValueError(f"Slide {slide_number} direction requires fly or wipe")
    transforms = {"path", "scale", "rotate"} & set(spec)
    if spec["effect"] == "transform":
        if not transforms:
            raise ValueError(f"Slide {slide_number} transform requires path, scale or rotate")
        if "start_visible" in spec and type(spec["start_visible"]) is not bool:
            raise ValueError(f"Slide {slide_number} start_visible must be a boolean")
        # Transforms normally move an already visible object. Explicit false
        # preserves the former native Custom entrance behavior when desired.
        spec = {**spec, "start_visible": spec.get("start_visible", True)}
        if "path" in spec:
            motion_path_tokens(spec["path"])
        if "scale" in spec and (not isinstance(spec["scale"], list) or len(spec["scale"]) != 2
                                or any(not finite_number(value, 1, 1000) for value in spec["scale"])):
            raise ValueError(f"Slide {slide_number} scale must be [xPercent, yPercent], each 1–1000")
        if "rotate" in spec and not finite_number(spec["rotate"], 0, 3600):
            raise ValueError(f"Slide {slide_number} rotate must be 0–3600 degrees; native By does not support negative values")
    elif transforms or "start_visible" in spec:
        raise ValueError(f"Slide {slide_number} path, scale, rotate and start_visible require effect transform")
    delay = spec.get("delay", 0.08 if spec["trigger"] == "after" else 0.0)
    accelerate = spec.get("accelerate", 0.0)
    decelerate = spec.get("decelerate", 0.0)
    if not finite_number(delay, 0, 10):
        raise ValueError(f"Slide {slide_number} effect delay must be 0–10 seconds")
    if (not finite_number(accelerate, 0, 1) or not finite_number(decelerate, 0, 1)
            or accelerate + decelerate > 1):
        raise ValueError(f"Slide {slide_number} acceleration and deceleration must total at most 1")
    return {**spec, "delay": float(delay), "accelerate": float(accelerate),
            "decelerate": float(decelerate)}


def add_transform(effect, spec: dict, enum) -> None:
    """Combine native Motion/Scale/Rotation behaviors in one custom effect.

    Microsoft object model: AnimationBehaviors.Add, MotionEffect.Path,
    ScaleEffect.FromX/FromY/ToX/ToY and RotationEffect.By.
    """
    for key, kind in (("path", "msoAnimTypeMotion"), ("scale", "msoAnimTypeScale"),
                      ("rotate", "msoAnimTypeRotation")):
        if key not in spec:
            continue
        behavior = effect.Behaviors.Add(enum(kind))
        if key == "path":
            behavior.MotionEffect.Path = spec[key]
        elif key == "scale":
            behavior.ScaleEffect.FromX = 100.0
            behavior.ScaleEffect.FromY = 100.0
            behavior.ScaleEffect.ToX = float(spec[key][0])
            behavior.ScaleEffect.ToY = float(spec[key][1])
        else:
            behavior.RotationEffect.By = float(spec[key])
        behavior.Timing.Duration = float(spec["duration"])
        behavior.Timing.TriggerDelayTime = 0.0
        # PowerPoint exposes Timing for behaviors but rejects their TriggerType
        # setter. Behaviors share the enclosing effect's trigger and start at 0.
        behavior.Timing.Accelerate = spec["accelerate"]
        behavior.Timing.Decelerate = spec["decelerate"]
    if spec.get("start_visible", True):
        # AddEffect(Custom) inserts Set(visibility=visible). This makes the
        # target hidden until the effect starts, including its delay. Remove
        # it only after adding motion behaviors: deleting the only behavior
        # would make PowerPoint delete the enclosing effect as well.
        for index in range(effect.Behaviors.Count, 0, -1):
            behavior = effect.Behaviors.Item(index)
            if (behavior.Type == enum("msoAnimTypeSet") and
                    behavior.SetEffect.Property == enum("msoAnimVisibility") and
                    behavior.SetEffect.To == 1):
                behavior.Delete()


def verify_transform(effect, spec: dict, enum, slide_number: int) -> None:
    expected = [(key, kind) for key, kind in (("path", "msoAnimTypeMotion"),
                ("scale", "msoAnimTypeScale"), ("rotate", "msoAnimTypeRotation")) if key in spec]
    # start_visible defaults to an already-visible transform. Explicit false
    # retains native Custom's initial hidden state and its visibility Set.
    offset = 0
    if not spec.get("start_visible", True) and effect.Behaviors.Count == len(expected) + 1:
        first = effect.Behaviors.Item(1)
        if (first.Type == enum("msoAnimTypeSet") and
                first.SetEffect.Property == enum("msoAnimVisibility") and first.SetEffect.To == 1):
            offset = 1
    if (effect.Behaviors.Count != len(expected) + offset or
            (not spec.get("start_visible", True) and offset != 1)):
        raise ValueError(f"Slide {slide_number} transform behavior or initial visibility changed after save")
    for index, (key, kind) in enumerate(expected, 1 + offset):
        behavior = effect.Behaviors.Item(index)
        if (behavior.Type != enum(kind) or abs(behavior.Timing.Duration - spec["duration"]) > 0.02
                or abs(behavior.Timing.TriggerDelayTime) > 0.02
                or abs(behavior.Timing.Accelerate - spec["accelerate"]) > 0.02
                or abs(behavior.Timing.Decelerate - spec["decelerate"]) > 0.02):
            raise ValueError(f"Slide {slide_number} transform behavior type or duration changed after save")
        if key == "path":
            valid = same_motion_path(behavior.MotionEffect.Path, spec[key])
        elif key == "scale":
            scale = behavior.ScaleEffect
            valid = all(abs(actual - wanted) <= 0.02 for actual, wanted in
                        ((scale.FromX, 100.0), (scale.FromY, 100.0),
                         (scale.ToX, spec[key][0]), (scale.ToY, spec[key][1])))
        else:
            valid = abs(behavior.RotationEffect.By - spec[key]) <= 0.02
        if not valid:
            raise ValueError(f"Slide {slide_number} transform {key} parameters changed after save")


def animate(source: Path, plan_path: Path, output: Path) -> None:
    plan = json.loads(plan_path.read_text(encoding="utf-8"))
    source, output = source.resolve(), output.resolve()
    if not source.is_file() or source.suffix.lower() != ".pptx":
        raise ValueError(f"PPTX source not found: {source}")
    if source == output:
        raise ValueError("Output must differ from source")
    if not isinstance(plan.get("slides"), list) or not plan["slides"]:
        raise ValueError("Plan requires a nonempty slides list")
    effects = {"fade": "msoAnimEffectFade", "wipe": "msoAnimEffectWipe",
               "zoom": "msoAnimEffectZoom", "fly": "msoAnimEffectFly", "transform": "msoAnimEffectCustom"}
    triggers = {"click": "msoAnimTriggerOnPageClick",
                "after": "msoAnimTriggerAfterPrevious",
                "with": "msoAnimTriggerWithPrevious"}
    transitions = {"fade": "ppEffectFadeSmoothly", "push": "ppEffectPushLeft",
                   "morph": "ppEffectMorphByObject"}
    directions = {"left": "msoAnimDirectionLeft", "right": "msoAnimDirectionRight",
                  "up": "msoAnimDirectionUp", "down": "msoAnimDirectionDown"}
    # Launch a separate PowerPoint process so an existing user session stays open.
    app = create_powerpoint()
    presentation = None
    temporary = None
    try:
        presentation = app.Presentations.Open(str(source), True, False, False)
        if presentation.Slides.Count != len(plan["slides"]):
            raise ValueError("Animation plan slide count differs from presentation")
        const = powerpoint_constants()
        fallback = {
            "msoAnimEffectFade": 10, "msoAnimEffectWipe": 22,
            "msoAnimEffectZoom": 23, "msoAnimEffectFly": 2,
            "msoAnimEffectCustom": 0, "msoAnimTypeMotion": 1,
            "msoAnimTypeScale": 3, "msoAnimTypeRotation": 4,
            "msoAnimTypeSet": 8, "msoAnimVisibility": 8,
            "msoAnimTriggerOnPageClick": 1,
            "msoAnimTriggerWithPrevious": 2,
            "msoAnimTriggerAfterPrevious": 3,
            "ppEffectFadeSmoothly": 3849, "ppEffectPushLeft": 3853,
            "ppEffectMorphByObject": 3954, "ppTransitionSpeedMedium": 2,
            "msoAnimDirectionLeft": 4, "msoAnimDirectionRight": 2,
            "msoAnimDirectionUp": 1, "msoAnimDirectionDown": 3,
        }

        def enum(name: str) -> int:
            return getattr(const, name, fallback[name])

        # Validate the whole plan before changing the PowerPoint presentation.
        shape_maps = []
        effect_specs = []
        for index, item in enumerate(plan["slides"], 1):
            if not isinstance(item, dict) or item.get("transition") not in transitions:
                raise ValueError(f"Slide {index} has unsupported transition")
            if "transition_duration" in item and not finite_number(item["transition_duration"], 0.1, 30):
                raise ValueError(f"Slide {index} transition_duration must be 0.1–30 seconds")
            slide = presentation.Slides.Item(index)
            if slide.TimeLine.MainSequence.Count or slide.TimeLine.InteractiveSequences.Count:
                raise ValueError(f"Slide {index} already has animations; refusing to duplicate them")
            names = {}
            for shape_index in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes.Item(shape_index)
                names.setdefault(shape.Name, []).append(shape)
            specs = [normalize_effect(spec, index) for spec in item.get("effects", [])]
            for spec in specs:
                if len(names.get(spec["shape"], [])) != 1:
                    raise ValueError(f"Slide {index} needs exactly one shape named {spec['shape']!r}")
            if item["transition"] == "morph":
                if index == 1:
                    raise ValueError("Morph requires a preceding slide")
                matched = {name for name in names if name.startswith("!!")
                           and name in shape_maps[-1]
                           and names[name][0].Type == shape_maps[-1][name][0].Type}
                if not matched:
                    raise ValueError(f"Slide {index} Morph needs a matching !!shape on the previous slide")
            shape_maps.append(names)
            effect_specs.append(specs)

        for index, item in enumerate(plan["slides"], 1):
            slide = presentation.Slides.Item(index)
            transition = slide.SlideShowTransition
            transition.EntryEffect = enum(transitions[item["transition"]])
            transition.AdvanceOnClick = True
            transition.Speed = enum("ppTransitionSpeedMedium")
            if "transition_duration" in item:
                transition.Duration = float(item["transition_duration"])
            for spec in effect_specs[index - 1]:
                effect = slide.TimeLine.MainSequence.AddEffect(
                    shape_maps[index - 1][spec["shape"]][0], enum(effects[spec["effect"]]), 0,
                    enum(triggers[spec["trigger"]]))
                if spec["effect"] == "transform":
                    add_transform(effect, spec, enum)
                if spec.get("direction"):
                    effect.EffectParameters.Direction = enum(directions[spec["direction"]])
                effect.Timing.Accelerate = spec["accelerate"]
                effect.Timing.Decelerate = spec["decelerate"]
                effect.Timing.Duration = float(spec["duration"])
                effect.Timing.TriggerDelayTime = spec["delay"]
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix="animated-", suffix=".pptx",
                                         dir=output.parent, delete=False) as handle:
            temporary = Path(handle.name)
        presentation.SaveAs(str(temporary), 24)
        presentation.Close()
        presentation = None
        checked = app.Presentations.Open(str(temporary), True, False, False)
        try:
            for index, item in enumerate(plan["slides"], 1):
                slide = checked.Slides.Item(index)
                if slide.TimeLine.MainSequence.Count != len(effect_specs[index - 1]):
                    raise ValueError(f"Slide {index} animation count changed after save")
                if slide.SlideShowTransition.EntryEffect != enum(transitions[item["transition"]]):
                    raise ValueError(f"Slide {index} transition changed after save")
                if ("transition_duration" in item and
                        abs(slide.SlideShowTransition.Duration - item["transition_duration"]) > 0.02):
                    raise ValueError(f"Slide {index} transition duration changed after save")
                for effect_index, spec in enumerate(effect_specs[index - 1], 1):
                    effect = slide.TimeLine.MainSequence.Item(effect_index)
                    if (effect.Shape.Name != spec["shape"] or effect.EffectType != enum(effects[spec["effect"]])
                            or effect.Timing.TriggerType != enum(triggers[spec["trigger"]])):
                        raise ValueError(f"Slide {index} animation target, effect or trigger changed after save")
                    actual_duration = effect.Timing.Duration
                    actual_delay = effect.Timing.TriggerDelayTime
                    if abs(actual_duration - spec["duration"]) > 0.02 or abs(actual_delay - spec["delay"]) > 0.02:
                        raise ValueError(f"Slide {index} animation timing changed after save: "
                                         f"duration {actual_duration} vs {spec['duration']}, "
                                         f"delay {actual_delay} vs {spec['delay']}")
                    if (abs(effect.Timing.Accelerate - spec["accelerate"]) > 0.02 or
                            abs(effect.Timing.Decelerate - spec["decelerate"]) > 0.02):
                        raise ValueError(f"Slide {index} animation easing changed after save")
                    if spec.get("direction") and effect.EffectParameters.Direction != enum(directions[spec["direction"]]):
                        raise ValueError(f"Slide {index} animation direction changed after save: "
                                         f"{effect.EffectParameters.Direction} vs "
                                         f"{enum(directions[spec['direction']])}")
                    if spec["effect"] == "transform":
                        verify_transform(effect, spec, enum, index)
        finally:
            checked.Close()
        os.replace(temporary, output)
        temporary = None
        print(f"Animated {len(plan['slides'])} slides: {output}")
    finally:
        if presentation is not None:
            presentation.Close()
        app.Quit()
        if temporary is not None:
            temporary.unlink(missing_ok=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("plan", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    try:
        animate(args.source, args.plan, args.output)
    except NativePowerPointUnavailable as exc:
        parser.exit(2, f"Error: {exc}\n")
