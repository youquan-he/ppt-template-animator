# 制作工具与实现边界

Python 3.10–3.13。基础依赖固定在 `requirements.txt`；原生功能另装 `requirements-windows.txt`，并需要 Windows 桌面版 Microsoft PowerPoint。先运行 `python -X utf8 scripts/doctor.py --skip-native` 检查基础环境；原生功能用 `--require-native` 检查。没有 Office 仍可生成示例、索引、读取主题和改编内容；原生动画、组件提取、渲染和视频不可用。模板源文件保持原样，中间文件放在工作目录。

## 索引和预览

```powershell
python scripts/template_catalog.py scan --roots <template-directory> --output <work>/catalog.json
python scripts/template_catalog.py inspect "<template.pptx>" --slides 7,12 --output <work>/candidate.json
python scripts/preview_template.py "<template.pptx>" --slides 7,12 --output-dir <work>/preview --video
```

目录优先级：`--roots`、`PPT_TEMPLATE_ROOTS`、当前目录的 `./templates`。环境变量的多目录分隔符在 Windows 为分号，在 macOS/Linux 为冒号。未配置时明确报错，不扫描个人盘符或整个工作目录。模板中是否有使用说明页必须实际检查，不能固定跳过前几页。统计用于筛选，图像和播放用于判断。非连续页抽取不能证明跨页 Morph 正确；此类效果需检查原始相邻页。

当前预览脚本的溢出提示主要覆盖顶层文本形状，组内文字、表格、继承占位和视觉重叠仍需逐页检查。`--video` 采用每页默认4秒的自动导出设置，用于观察效果，不代表现场点击与讲述节奏；最终按分镜在 PowerPoint 中核查实际播放，不能仅凭导出视频判定时长合格。

## 研究连续场景与时间轴

```powershell
python -X utf8 scripts/study_motion.py "<template.pptx>" --slides 1-4 --output <work>/motion-study.json
```

用于高级动态请求；具体设计方法见 [advanced-motion.md](advanced-motion.md)。按演示稿的真实页序和关系读取形状、组、几何状态、原始路径/缩放/旋转参数及触发时间，记录源指纹、媒体线索和 Morph 邻页依赖。选择需要的完整连续范围，不将单页截图当作场景。

报告区分有效现代兼容分支与原始 XML 统计；`mc:AlternateContent` 的 Choice/Fallback 不能相加成实际播放数量。未识别扩展和分支支持会在报告中提示。动画节点数为零也可能是跨页 Morph；反之，节点多不说明动画精美。工具仅研究结构，不迁移动画、不证明实际播放，仍须对照源文件的渲染与视频。

## 拆解可编辑组件

```powershell
python scripts/decompose_template.py inspect "<template.pptx>" --slide 12 --output <work>/shapes.json
python scripts/decompose_template.py extract "<template.pptx>" --slide 12 --shape-ids 17,20 --output-dir <work>/component-01
```

必须使用检查到的 ID。省略 `--shape-ids` 保留该页全部顶层对象；选中组时连同子对象保留，不能直接以子对象 ID 代替顶层组。输出目录须为新目录。

`inspect` 还输出关联的版式、母版与主题部件，主题颜色、主/次字体、中文等脚本字体，以及各对象直接声明的字体、字号和颜色。它沿实际关系查找主题，不假定所有页面使用 theme1.xml。直接样式和主题声明不等于最终有效样式：颜色映射、主题覆盖、字体替代和未显式声明的继承仍需对照渲染图。

- `components.pptx`：所选原生形状、分组、图片或图表，由 PowerPoint 保存并重新打开校验。
- `media/`：所选对象通过图片关系引用的原始媒体，包括可读取的组内图片和 SVG/位图备用图。
- `source-slide.png`、`components.png`：原始上下文和拆解后画面。
- `manifest.json`：来源指纹、源页、形状结构、主题/直接样式报告、媒体来源、原始及保留的主动画序列、校验结果。
- `source-timing.xml`：原始动画结构，用于研究复杂效果和依赖，不能脱离关系文件直接导入。

组件文件保留源母版及幻灯片背景；以普通形状绘制的背景若未选中则会删除。对照两张预览检查依赖，例如原来深色背景上的白字需要在新主题中重新配色。用于新稿时复制所需对象、统一样式和重新布置，不要无意沿用样例文字、整页背景或备注。抽出的单页不能独立证明 Morph 连续性。存在交互触发序列时，脚本拒绝部分拆解，须保留整页上下文或在新稿重建触发关系。保留主动画目标和效果数量仍不代表点击节奏适合新场景；须重新设计并播放。

## 整页改编

```powershell
python scripts/build_deck.py <work>/plan.json
```

计划 `template` 与 `output` 必须不同。`slides` 按输出顺序列出源页 `number`；此构建器中同一源页只能使用一次。

- `text`：形状 ID 到新文字，沿用首个 run 样式，多段以 `\n` 分隔，段数不得超过源对象。
- `pictures`：普通图片 ID 到本地路径，保留对象并按原框重算裁切，需视觉核查。
- `tables`：表格 ID 到完整二维字符串数组，维度须与源表一致。
- `keep_text`、`keep_tables`：明确确认可保留内容。默认 `strict_text: true`，未处理文字/表格会阻止构建。
- `transition`：可选 fade、push-left、push-right、wipe-left、wipe-right。
- `notes`：讲者备注；根级 `strip_brand_text` 清空正文、版式及母版中的精确品牌短语。
- `require_animation` 默认 true；需要静态页时可明确设 false，避免为满足统计添加无意义动画。
- `native_validation`：auto（默认）在可用时原生打开，宿主不可用时明确报告 false；required 要求原生检查成功；off 只执行包结构/内容检查。CLI `--native-validation` 可覆盖计划。auto 不能吞掉文件损坏等实际开档错误。

构建器核查页数、文字、表格和图片，并比较保存前后的动画语义结构：目标、节点顺序、触发、时长、路径和效果参数。比较忽略 XML 前缀及格式化空白，拒绝悬空形状目标；转场与未改源结构或明确修改后的预期一致。该检查证明结构保留，不替代实际播放。在 Windows 原生打开成功后报告 `native_validated: true`；无原生宿主时报告 false，不能据此称 PowerPoint 验收通过。失败不覆盖已有输出。

## 新构图的原生动画

```powershell
python scripts/animate_deck.py <source.pptx> <animation-plan.json> <output.pptx>
```

计划 `slides` 与演示稿逐页对应。每页 `transition` 为 fade、push 或 morph。`effects` 可用 `[形状名, 效果, 触发方式, 时长秒数]` 或对象形式：

```json
{"shape":"Figure","effect":"fly","trigger":"after","duration":0.7,"delay":0.15,"direction":"up","accelerate":0.15,"decelerate":0.2}
```

效果支持 fade/wipe/zoom/fly/transform；触发支持 click/after/with；方向 left/right/up/down 用于 fly/wipe。时长为 0.1–30 秒，delay 为 0–10 秒；加减速为持续时间比例，和不超过 1。每页可用 `transition_duration` 设置转场时长。Morph 需要前一页同名同类的 `!!shape_name`。脚本验证唯一形状名、配对和存盘后参数，拒绝重复添加已有动画。

`transform` 在同一个原生效果中组合路径、缩放和旋转，至少提供一个参数：

```json
{"shape":"Figure","effect":"transform","trigger":"after","duration":1.4,"delay":0.15,"path":"M0 0 C.06 -.1 .15 -.1 .22 0 E","scale":[120,85],"rotate":45,"accelerate":0.1,"decelerate":0.2}
```

`path` 是归一化幻灯片坐标的 VML 路径，支持大写 M/L/C/Z/E；`scale` 是从原始 100% 到目标百分比的 `[x,y]`；`rotate` 是原生 By 顺时针角度。它们共用效果的触发与时长。`transform` 默认 `start_visible: true`，延迟期间对象已可见，适合遮罩或现有对象移动；显式 `false` 才在动作开始时出现。遮罩若在延迟期间隐藏，会使内容提前闪出，必须播放检查初态。数值范围、严格路径语法和保存后的具体行为都会验证。重新排版后应调整轨迹并检查最终状态，不能机械照搬源坐标。

辅助脚本的范围不是 Skill 的设计上限。迁移遮罩、文字分段、复杂触发或其他原生机制时，研究源时间轴，通过合适的原生接口制作并重新绑定对象，在代表页验证后使用。保留可编辑资产；不能用视频掩盖原生动画未实现。无 Office 时能生成原创样例的静态基稿和计划，不能将其声称为已完成高级动画。
