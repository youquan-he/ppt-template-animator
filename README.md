# ppt-template-animator

**内容完整、可编辑、可验证的 PPT 制作 Skill 与本地工具。**

A self-contained presentation skill and toolkit. Generate the demo without private templates or API keys; add native animations and previews with Windows desktop PowerPoint.

本项目包含素材覆盖、需求判断、模板学习与拆解、原创重组、中文排版和动态验收流程。它提供确定性的辅助工具；设计质量仍需要阅读素材、逐页看图和实际播放，不能由动画数量或测试总数保证。

高级动画走源序列研究、原生迁移/重建和播放验收流程，支持对象路径、缩放、旋转复合编排与连续 Morph。视觉方向依据主题和用户偏好选择，避免把静态稿加统一淡入当成高级动态。详细方法见 [高级动画制作](references/advanced-motion.md)。

## 运行条件

建议 Python 3.13；支持配置为 Python 3.10–3.13。首次安装依赖需要网络，示例生成和本地处理无需网络或 API key。

| 能力 | Windows 无 Office | macOS / Linux | Windows 桌面 PowerPoint |
| --- | --- | --- | --- |
| 生成原创示例、模板索引、形状/主题检查 | 支持 | 支持 | 支持 |
| 按计划替换文字、图片和表格，检查包结构 | 支持 | 支持 | 支持 |
| 原生开档验证、添加动画、可编辑组件拆解 | 不支持 | 不支持 | 另装可选依赖后支持 |
| 原生渲染 PNG / 导出 MP4 | 不支持 | 不支持 | 另装可选依赖后支持 |

PowerPoint 网页版、PowerPoint for Mac 和 LibreOffice 不是本项目 COM 工具的宿主。未通过 `doctor.py --require-native` 时，不运行依赖原生宿主的步骤。

## 安装

下载或克隆仓库后，在仓库根目录创建虚拟环境。不要在系统 Python 中混装依赖。

Windows PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

macOS / Linux：

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
```

下面命令中的 `python` 指这个虚拟环境的解释器。可激活环境，也可直接替换为 `.\.venv\Scripts\python.exe` 或 `.venv/bin/python`；无需修改 PowerShell 执行策略。版本固定在 `requirements.txt`，Windows 可选依赖在 `requirements-windows.txt`。

## 无需模板和 Office 的完整示例

在仓库根目录依次运行：

```bash
python -X utf8 scripts/doctor.py --skip-native
python -X utf8 examples/make_demo.py --output-dir work/demo
python -X utf8 scripts/template_catalog.py scan --roots work/demo/templates --output work/demo/catalog.json
python -X utf8 scripts/build_deck.py work/demo/build-plan.json --native-validation off
python -X utf8 scripts/decompose_template.py inspect work/demo/adapted.pptx --slide 2 --output work/demo/inspect.json
python -X utf8 -m unittest discover -s tests -v
```

预期产物：

```text
work/demo/
├── templates/demo-template.pptx   # 原创静态示例模板
├── assets/original.png           # 代码生成的几何素材
├── assets/replacement.png
├── build-plan.json               # 所有路径相对该计划文件
├── animation-plan.json
├── material-ledger.json
├── adapted.pptx                  # 替换文字和图片后的三页稿
├── catalog.json
└── inspect.json
```

构建结果应为3页，并报告 `native_validation: off`、`native_validated: false`。这表示结构与内容检查通过，**没有完成原生播放或视觉验收**。示例刻意采用静态模板；原生动画在下一节单独添加。素材台账初始状态是 planned，不能作为已完成视觉检查的证据。

示例素材由代码生成，不下载模板、不包含第三方图片，不需要安装本项目参考过的其他 Skill。它是功能演示，不是医学、科研或业务事实。再次运行请使用新目录，例如 `work/demo-2`，生成器不会覆盖已有目录。

## Windows 原生动画、拆解与预览

在可正常启动桌面 PowerPoint 的 Windows 用户会话中运行：

```powershell
python -m pip install -r requirements-windows.txt
python -X utf8 scripts/doctor.py --require-native
python -X utf8 scripts/build_deck.py work/demo/build-plan.json --native-validation required
python -X utf8 scripts/animate_deck.py work/demo/adapted.pptx work/demo/animation-plan.json work/demo/animated.pptx
python -X utf8 scripts/preview_template.py work/demo/animated.pptx --slides 1,2,3 --output-dir work/demo/preview
python -X utf8 scripts/decompose_template.py extract work/demo/animated.pptx --slide 2 --output-dir work/demo/components
```

可在预览命令末尾加 `--video` 导出 MP4。当前视频采用每页默认4秒的自动导出策略，不能代替现场点击与演讲时长测试。`extract` 省略 `--shape-ids` 时保留该页全部顶层对象；需要部分组件时先 `inspect`，再使用实际 ID。组件中的原生对象可继续编辑，原始图片和动画信息另存；详见 [工具说明](references/production-tools.md)。

基础示例与原生示例都是本地流程；原生路径需要安装并可用的 PowerPoint，本仓库不提供 Office。

## 原创复合动画与连续场景示例

这份四页原创样例演示组合图形的路径/缩放/旋转、连续气泡总览到聚焦、错峰遮罩揭示。全部为可编辑的原生图形和文字，不依赖私有模板或图片；它是机制与流程验证，不是所有主题的默认设计。

```powershell
python -X utf8 examples/make_advanced_demo.py --output-dir work/advanced-demo
python -X utf8 scripts/animate_deck.py work/advanced-demo/base.pptx work/advanced-demo/animation-plan.json work/advanced-demo/animated.pptx
python -X utf8 scripts/preview_template.py work/advanced-demo/animated.pptx --slides 1,2,3,4 --output-dir work/advanced-demo/preview --video
```

第一步可在没有 Office 的环境运行，生成静态 `base.pptx`、可执行 `animation-plan.json` 及初始为 `planned` 的 `scene-record.json`。后两步需要 Windows 桌面 PowerPoint；不能把第一步的遮罩初态或静态文件当作已完成动画。生成器要求新输出目录。实际播放后在工作记录中更新检查范围，不能自动将计划状态写成已验收。

## 使用自己的模板

模板库无需上传到仓库，也不要求使用作者的模板。三种目录来源按以下顺序生效：

1. `--roots` 显式传入，可指定多个目录。
2. `PPT_TEMPLATE_ROOTS` 环境变量：Windows 用分号分隔，macOS/Linux 用冒号分隔。
3. 当前工作目录下存在的 `./templates`。

```bash
python -X utf8 scripts/template_catalog.py scan --roots templates --output work/catalog.json
python -X utf8 scripts/template_catalog.py inspect templates/example.pptx --slides 1,3 --output work/candidate.json
```

这里的 `templates/example.pptx` 是你自己提供的文件，仓库中不包含它。无目录配置时脚本会报可读错误，不扫描个人盘符或整个计算机。改编计划里的 `template`、`output` 和图片路径均相对计划 JSON 所在目录解析；也可以明确传入绝对路径，但不要将私有计划提交到公开仓库。

`native_validation` 可在计划中设为 `auto`、`required` 或 `off`，CLI 参数优先：

- `auto`：宿主可用就原生验证；宿主不可用时报告 false。文件实际无法打开仍会失败。
- `required`：缺宿主或无法原生打开都失败。
- `off`：只执行内容和包结构检查，明确不声称原生验证。

## 作为 Skill 使用

保持整个目录结构，将 `SKILL.md` 及其 `scripts/`、`references/` 一起交给支持 Skill 的 AI 客户端。安装位置由客户端决定，本项目不假定某个个人主目录，也不会自动修改客户端配置。仅复制一份 `SKILL.md` 会丢失脚本和参考文件。

也可直接让代理读取仓库的 `SKILL.md`，说明素材目录、观众、用途、时间/页数和模板要求。无需代理即可单独运行上面的命令行工具。

需要记住本机模板库和个人视觉偏好时，将 `preferences.example.json` 复制为 `preferences.local.json` 后填写目录与偏好。它只供代理读取，不会自动变成命令行参数；脚本仍使用显式目录或环境变量。该本机文件及 `references/local-motion-library.json` 不进入公开包。公共版本没有对某个盘符、个人模板库或特定配色的依赖。

- [制作流程](SKILL.md)
- [素材台账](references/material-ledger.md)
- [设计、中文排版与复核](references/design-and-review.md)
- [脚本参数、能力与限制](references/production-tools.md)
- [高级动态：连续场景、复合运动与验收](references/advanced-motion.md)

## 测试与复现范围

默认测试只运行便携测试，原生测试会明确显示 skip。Windows 安装了 PowerPoint 和可选依赖后，可显式启用集成测试：

```powershell
$env:PPT_NATIVE_TESTS = '1'
python -X utf8 -m unittest discover -s tests -v
```

开启后宿主缺失会导致失败，不能通过跳过伪装成功。关闭可将该变量设为 `0`。

GitHub Actions 配置使用 Linux / Windows 和 Python 3.10 / 3.13 运行便携测试及示例流水线；托管 CI 不验证桌面 PowerPoint 动画。仓库附带工作流不等于已经在你的 GitHub 账户运行成功，首次推送后应查看实际结果。

复现目标是相同功能、页数、素材替换和检查结果，不是二进制逐字节相同；PowerPoint 元数据、字体和渲染环境会影响文件及画面。测试不能自动保证设计审美或资料事实正确。

常见问题：

- `No template directory configured`：提供 `--roots` 或生成上面的示例。
- 原生宿主不可用：先查看 `doctor.py --require-native` 的 code 和 message，确认运行在 Windows、可选依赖安装于当前环境且桌面 PowerPoint 可启动。
- 同一路径重复生成失败：示例/组件生成器要求新目录，改用新目录；构建器则只在验证成功后替换输出。
- 字体替换或缺字：检查目标机器实际字体并重新渲染；主题声明本身不能保证字体存在。
- 含中文路径的输出乱码：使用 `python -X utf8` 或设置 `PYTHONUTF8=1`。

## 公开发布与许可

仓库以 [MIT License](LICENSE) 发布；输入的第三方模板、字体和图片不随本项目获得新授权，见 [NOTICE.md](NOTICE.md)。

`.gitignore` 排除了虚拟环境、工作目录、私有模板、生成的 PPTX/视频和环境变量文件。Git 忽略规则不会自动移除已经被跟踪的文件；首次发布前检查将提交的文件列表。仓库不需要 API key、你的原始 PPT、整个模板库或本地软件缓存。
