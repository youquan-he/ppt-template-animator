"""Apply native PowerPoint animations and transitions from a small JSON plan."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import tempfile

from native_powerpoint import NativePowerPointUnavailable, create_powerpoint, powerpoint_constants

def normalize_effect(spec: list | dict, slide_number: int) -> dict:
    if isinstance(spec, list) and len(spec) == 4:
        spec = dict(zip(("shape", "effect", "trigger", "duration"), spec))
    if not isinstance(spec, dict):
        raise ValueError(f"Slide {slide_number} effect must be a four-item list or an object")
    allowed = {"shape", "effect", "trigger", "duration", "delay", "direction",
               "accelerate", "decelerate"}
    if set(spec) - allowed or not {"shape", "effect", "trigger", "duration"} <= set(spec):
        raise ValueError(f"Slide {slide_number} effect has missing or unsupported fields")
    if not isinstance(spec["shape"], str) or not spec["shape"]:
        raise ValueError(f"Slide {slide_number} effect shape must be a name")
    if spec["effect"] not in {"fade", "wipe", "zoom", "fly"} or spec["trigger"] not in {"click", "after", "with"}:
        raise ValueError(f"Slide {slide_number} has unsupported effect or trigger")
    if type(spec["duration"]) not in (int, float) or not 0.1 <= spec["duration"] <= 5:
        raise ValueError(f"Slide {slide_number} effect duration must be 0.1–5 seconds")
    if "direction" in spec and (spec["effect"] not in {"fly", "wipe"} or
                                spec["direction"] not in {"left", "right", "up", "down"}):
        raise ValueError(f"Slide {slide_number} direction requires fly or wipe")
    delay = spec.get("delay", 0.08 if spec["trigger"] == "after" else 0.0)
    accelerate = spec.get("accelerate", 0.0)
    decelerate = spec.get("decelerate", 0.0)
    if type(delay) not in (int, float) or not 0 <= delay <= 10:
        raise ValueError(f"Slide {slide_number} effect delay must be 0–10 seconds")
    if (type(accelerate) not in (int, float) or type(decelerate) not in (int, float)
            or not 0 <= accelerate <= 1 or not 0 <= decelerate <= 1
            or accelerate + decelerate > 1):
        raise ValueError(f"Slide {slide_number} acceleration and deceleration must total at most 1")
    return {**spec, "delay": float(delay), "accelerate": float(accelerate),
            "decelerate": float(decelerate)}


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
               "zoom": "msoAnimEffectZoom", "fly": "msoAnimEffectFly"}
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
            if item.get("transition") not in transitions:
                raise ValueError(f"Slide {index} has unsupported transition")
            slide = presentation.Slides.Item(index)
            if slide.TimeLine.MainSequence.Count:
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
            for spec in effect_specs[index - 1]:
                effect = slide.TimeLine.MainSequence.AddEffect(
                    shape_maps[index - 1][spec["shape"]][0], enum(effects[spec["effect"]]), 0,
                    enum(triggers[spec["trigger"]]))
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
                for effect_index, spec in enumerate(effect_specs[index - 1], 1):
                    effect = slide.TimeLine.MainSequence.Item(effect_index)
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
