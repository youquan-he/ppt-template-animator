"""Generate an original editable four-slide motion demonstration.

This creates a static base and a native animation plan; use animate_deck.py on
Windows PowerPoint to apply the plan. It contains no private template assets.
The scenes illustrate content-to-scene design, not business or research claims.
"""

from __future__ import annotations

import argparse
import json
import math
import os
from pathlib import Path
import tempfile

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


W, H = 13.333333, 7.5
INK, PAPER = "241C30", "F7F0E6"
VIOLET, CORAL, LIME = "7139E8", "FF684E", "DDFA49"
TEAL, YELLOW = "3AC7B4", "FFCF54"
FONT = "Microsoft YaHei"


def shape(slide, name, kind, x, y, width, height, color, rotation=0):
    item = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(width), Inches(height))
    item.name = name
    item.fill.solid()
    item.fill.fore_color.rgb = RGBColor.from_string(color)
    item.line.fill.background()
    item.rotation = rotation
    return item


def text(slide, name, copy, x, y, width, height, size, color=INK, bold=False,
         align=PP_ALIGN.LEFT):
    item = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    item.name = name
    frame = item.text_frame
    frame.word_wrap = True
    frame.margin_left = frame.margin_right = 0
    frame.margin_top = frame.margin_bottom = 0
    for index, line in enumerate(copy.split("\n")):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = line
        paragraph.font.name = FONT
        paragraph.font.size = Pt(size)
        paragraph.font.bold = bold
        paragraph.font.color.rgb = RGBColor.from_string(color)
        paragraph.alignment = align
        paragraph.space_after = Pt(0)
    return item


def background(slide, color):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor.from_string(color)


def foot(slide, index, color=INK):
    text(slide, "SceneFooter", "内容成景  /  MOTION STUDY", 0.6, 7.02, 8.8, 0.24, 10, color)
    text(slide, "SceneNumber", f"0{index}", 11.85, 6.94, 0.85, 0.35, 17, color,
         align=PP_ALIGN.RIGHT)


def bubble(slide, key, label, number, x, y, diameter, color, focused=False):
    item = shape(slide, "!!Bubble" + key, MSO_SHAPE.OVAL, x, y, diameter, diameter, color)
    frame = item.text_frame
    frame.clear()
    frame.margin_left = frame.margin_right = Inches(0.10)
    frame.margin_top = frame.margin_bottom = Inches(0.06)
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    foreground = PAPER if color == VIOLET else INK
    for i, (copy, size) in enumerate(((number, 14 if diameter < 1.5 else 19),
                                     (label, 17 if diameter < 1.5 else 30 if focused else 27))):
        para = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        para.text = copy
        para.alignment = PP_ALIGN.CENTER
        para.font.name = FONT
        para.font.bold = i == 1
        para.font.size = Pt(size)
        para.font.color.rgb = RGBColor.from_string(foreground)
        para.space_after = Pt(4 if i == 0 else 0)
    return item


def cover(deck):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    background(slide, PAPER)
    shape(slide, "CoverField", MSO_SHAPE.RECTANGLE, 7.14, 0, 6.2, 7.5, CORAL)
    text(slide, "CoverKicker", "从素材出发 · 让动作讲清楚", 0.62, 0.58, 6.3, 0.35, 15)
    text(slide, "CoverTitle", "让内容\n成为场景", 0.55, 1.38, 6.65, 2.55, 58, bold=True)
    shape(slide, "CoverRule", MSO_SHAPE.RECTANGLE, 0.64, 4.65, 0.6, 0.08, VIOLET)
    text(slide, "CoverStatement", "先决定观众要理解什么，\n再决定画面怎样运动。", 0.64, 5.02, 6.0, 0.95, 23)
    text(slide, "CoverRoute", "素材  →  叙事  →  视觉  →  动效  →  验收", 0.65, 6.36, 6.15, 0.4, 13)
    pieces = []
    cx, cy = 10.12, 3.4
    for i in range(12):
        angle = math.radians(i * 30)
        px, py = cx + 1.71 * math.cos(angle), cy + 1.71 * math.sin(angle)
        pieces.append(shape(slide, f"OrbitBlade{i:02d}", MSO_SHAPE.ROUNDED_RECTANGLE,
                            px - 0.30, py - 0.72, 0.60, 1.44,
                            (VIOLET, LIME, PAPER)[i % 3], i * 30 + 90))
    pieces.append(shape(slide, "OrbitCore", MSO_SHAPE.OVAL, cx - 1.11, cy - 1.11, 2.22, 2.22, INK))
    wheel = slide.shapes.add_group_shape(pieces)
    wheel.name = "OrbitAssembly"
    text(slide, "OrbitText", "想法\n成形", cx - 0.91, cy - 0.64, 1.82, 1.3, 30, PAPER,
         True, PP_ALIGN.CENTER)
    text(slide, "CoverMotionCaption", "路径 × 缩放 × 旋转", 8.11, 6.2, 4.4, 0.45, 17, INK,
         align=PP_ALIGN.CENTER)
    slide.notes_slide.notes_text_frame.text = (
        "原创原生矢量场景。入场后 OrbitAssembly 自动沿曲线小幅移位、放大并旋转；"
        "文字中心单独保持静止，保证阅读。所有轮片均可取消组合编辑。"
        "此页用于验证原生路径、缩放和旋转同步，不以效果数量评价审美。")
    foot(slide, 1)


def overview(deck):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    background(slide, PAPER)
    text(slide, "OverviewKicker", "01—05  /  建立一条清晰主线", 0.65, 0.58, 10.3, 0.32, 14)
    text(slide, "OverviewTitle", "不是把素材摆满，\n而是让关系被看见。", 0.59, 1.15, 11.9, 1.48, 37, bold=True)
    entries = [
        ("Material", "素材", "01", 0.65, 3.6, 2.13, TEAL),
        ("Story", "叙事", "02", 2.91, 3.12, 2.78, CORAL),
        ("Visual", "视觉", "03", 5.72, 3.62, 2.18, LIME),
        ("Motion", "动效", "04", 8.0, 3.02, 2.6, VIOLET),
        ("Review", "验收", "05", 10.7, 3.68, 1.97, YELLOW),
    ]
    for entry in entries:
        bubble(slide, *entry)
    text(slide, "OverviewCaption", "先看全局，再选择一个值得放大的重点。", 0.68, 6.35, 11.9, 0.42, 19)
    slide.notes_slide.notes_text_frame.text = (
        "这是连续 Morph 场景的前态。下一页保持五个原生圆形的 !!Bubble 名称与形状类型；"
        "叙事气泡放大，其他环节缩小并让出阅读空间。请连续播放第 2—3 页，勿只抽单页。")
    foot(slide, 2)


def focus(deck):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    background(slide, PAPER)
    text(slide, "FocusKicker", "聚焦 02  /  叙事", 0.65, 0.58, 6.4, 0.32, 14)
    text(slide, "FocusTitle", "先讲清楚，\n为什么重要。", 0.59, 1.27, 6.4, 1.65, 44, bold=True)
    text(slide, "FocusBody", "把事实组织成问题、证据与结论。\n画面放大的，是当前要理解的关系。", 0.65, 3.39, 6.3, 1.0, 23)
    bubble(slide, "Story", "叙事", "02", 7.36, 1.45, 4.75, CORAL, True)
    for entry in [
        ("Material", "素材", "01", 0.67, 5.26, 1.22, TEAL),
        ("Visual", "视觉", "03", 2.09, 5.26, 1.22, LIME),
        ("Motion", "动效", "04", 3.51, 5.26, 1.22, VIOLET),
        ("Review", "验收", "05", 4.93, 5.26, 1.22, YELLOW),
    ]:
        bubble(slide, *entry)
    shape(slide, "FocusPointer", MSO_SHAPE.RIGHT_ARROW, 6.34, 3.46, 0.58, 0.33, VIOLET)
    slide.notes_slide.notes_text_frame.text = (
        "第 2 页到本页采用原生 Morph，时长 1.6 秒；五个同名原生圆形同时换位/缩放。"
        "这是从总览到重点的叙事聚焦，而不是给静态圆形追加五次淡入。"
        "检查其他气泡退让过程、圆形和文字连续性，以及焦点形成后正文可读性。")
    foot(slide, 3)


def reveal(deck):
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    background(slide, VIOLET)
    text(slide, "RevealKicker", "最后一步  /  让结果经得起播放", 0.65, 0.58, 11.6, 0.35, 14, PAPER)
    text(slide, "RevealTitle", "揭开画面，也揭开逻辑。", 0.59, 1.18, 12.0, 0.82, 38, PAPER, True)
    text(slide, "RevealSubtitle", "错峰揭示三个检查面，最后一起停留。", 0.65, 2.16, 11.7, 0.45, 20, PAPER)
    colors = (PAPER, LIME, CORAL)
    titles = ("内容完整", "画面有主角", "动作有意义")
    bodies = ("每份必用素材\n都有实际去向", "主题、色彩与层级\n共同指向重点", "移动、缩放与揭示\n服务于理解")
    for i in range(3):
        x = 0.65 + i * 4.16
        shape(slide, f"ReviewPanel{i}", MSO_SHAPE.RECTANGLE, x, 3.02, 3.71, 2.92, colors[i])
        text(slide, f"ReviewIndex{i}", f"0{i + 1}", x + 0.22, 3.23, 3.1, 0.45, 19)
        text(slide, f"ReviewTitle{i}", titles[i], x + 0.22, 3.88, 3.25, 0.55, 27, INK, True)
        text(slide, f"ReviewBody{i}", bodies[i], x + 0.22, 4.76, 3.22, 0.85, 20)
        # A real foreground group covers the content, then moves down and shrinks.
        # The irregular cap creates a moving diagonal edge instead of a title fly-in.
        pieces = [shape(slide, f"RevealMaskBody{i}", MSO_SHAPE.RECTANGLE,
                        x - 0.02, 3.0, 3.75, 2.98, "8A58EE"),
                  shape(slide, f"RevealMaskCap{i}", MSO_SHAPE.PARALLELOGRAM,
                        x - 0.02, 2.84, 3.75, 0.74, "8A58EE")]
        mask = slide.shapes.add_group_shape(pieces)
        mask.name = f"RevealMask{i}"
    text(slide, "RevealResult", "素材 → 叙事 → 视觉 → 动效 → 验收", 0.67, 6.35, 11.7, 0.43, 20, PAPER)
    foot(slide, 4, PAPER)
    slide.notes_slide.notes_text_frame.text = (
        "三个不规则前景遮罩与内容是不同原生对象。入场后遮罩按 0、0.22、0.44 秒错峰，"
        "沿竖向路径向下移走并缩小；下方卡片和文字保持稳定。最终遮罩移到页面外，"
        "形成三项并列验收观点。检查揭示中途边缘、最终挡字和播放停留；"
        "静态 base.pptx 有意保留遮罩初态，不是最终结束画面。")


def create_demo(output_dir: Path) -> dict:
    output_dir = output_dir.resolve()
    if output_dir.exists():
        raise ValueError("Output directory must be new; existing files are preserved")
    output_dir.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="advanced-demo-", dir=output_dir.parent) as temp:
        stage = Path(temp) / "demo"
        stage.mkdir()
        deck = Presentation()
        deck.slide_width, deck.slide_height = Inches(W), Inches(H)
        for scene in (cover, overview, focus, reveal):
            scene(deck)
        deck.core_properties.title = "让内容成为场景"
        deck.core_properties.subject = "Original editable motion-mechanism demonstration"
        deck.core_properties.author = ""
        deck.save(stage / "base.pptx")
        plan = {"slides": [
            {"transition": "fade", "transition_duration": 0.5, "effects": [
                {"shape": "OrbitAssembly", "effect": "transform", "trigger": "with",
                 "duration": 2.4, "delay": 0.15, "path": "M0 0 C.008 -.025 .028 -.025 .035 0 E",
                 "scale": [108, 108], "rotate": 60, "accelerate": 0.15, "decelerate": 0.3}]},
            {"transition": "fade", "transition_duration": 0.5, "effects": []},
            {"transition": "morph", "transition_duration": 1.6, "effects": []},
            {"transition": "fade", "transition_duration": 0.45, "effects": [
                {"shape": f"RevealMask{i}", "effect": "transform", "trigger": "with",
                 "duration": 1.15, "delay": i * 0.22,
                 "path": "M0 0 C.005 .16 -.005 .35 0 .58 E", "scale": [92, 70],
                 "accelerate": 0.1, "decelerate": 0.35} for i in range(3)]},
        ]}
        (stage / "animation-plan.json").write_text(
            json.dumps(plan, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        record = {
            "title": "让内容成为场景", "original_artwork": True,
            "source_assets": [], "status": "planned",
            "native_playback_verified": False,
            "scenes": [
                {"slides": [1], "mechanism": "compound path/scale/rotation", "target": "OrbitAssembly"},
                {"slides": [2, 3], "mechanism": "continuous native Morph with stable !! identities",
                 "target": "!!BubbleMaterial/Story/Visual/Motion/Review"},
                {"slides": [4], "mechanism": "staggered foreground masking with path and scale",
                 "target": "RevealMask0/1/2"},
            ],
            "verification_needed": ["native animation apply and reopen", "full-size layout inspection",
                                    "actual playback and endpoint inspection"],
        }
        (stage / "scene-record.json").write_text(
            json.dumps(record, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        os.replace(stage, output_dir)
    return {"output_dir": str(output_dir), "slides": 4, "base": "base.pptx",
            "plan": "animation-plan.json", "native_animation_included": False,
            "artwork": "original native editable vectors and text"}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = create_demo(args.output_dir)
    except (OSError, ValueError) as exc:
        parser.error(str(exc))
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
