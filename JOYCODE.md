# JOYCODE.md

本仓库的核心是一个 JoyCode Skill：**`docx-spec-html`** —— 把京东「商品信息运营规范」类 Word `.docx` 文档转成生产级单文件 HTML 长图页。`GeneratedProducts/` 存放用该 skill 生成的实际产物。

## 仓库结构

- `docx-spec-html/` —— skill 本体（`SKILL.md` 是权威说明，动手前必读）
  - `SKILL.md` —— 完整工作流、交付前必检清单、不可协商的输出规则
  - `references/design.md` —— 设计系统、CSS 模板、Word→HTML 映射、验收清单（体量大，100KB+）
  - `references/high-quality-workflow.md` —— 高质量工作流与常见失败模式
  - `scripts/` —— `extract_docx_manifest.py`（抽料）、`batch_generate.py`（草稿生成器）、`validate_output.py`（校验）
  - `assets/` —— `styles.css`、hero 字体、golden 示例、`vendor/`（html-editor.html + html2canvas.min.js）
  - `agents/openai.yaml` —— 外部工具集成的 agent 接口定义
- `GeneratedProducts/` —— 生成产物；`*-rebuild/` 目录里的 `probe_*.py`/`patch_*.py`/`inspect_*.py` 是一次性调试脚本，不是可复用工具
- `.joycode/memory/` —— 跨会话记忆（已记录「PDF 为主料」这条核心反馈）

## 最重要的规则（违反过会被用户挑错）

1. **PDF 才是结构标准答案，不是 docx。** `.docx` 几乎都是从 PDF 导出的，导出破坏了层级：无 heading 样式、合并单元格丢失、模块标题（`1、主图规范`）和正文要点（`1、视频画面需清晰`）都变成同样的编号段落。每次任务都要向用户索取配套 PDF。
2. **绝不把 `batch_generate.py` 的原始输出当成品交付。** 它只是草稿和抽料工具。正确流程是三段式：脚本抽料 → 模型读 PDF 重建层级 → 对照 PDF 逐屏审查修正。中间那步（模型重建）由模型做，永远不能跳过——跳过它曾导致 10 处结构错误。
3. **8 个一级模块固定**：主图规范 / 主图视频 / 长标题 / 短标题 / 通用卖点 / 主推标签 / 品质标签 / 属性。渲染时去掉标题里的 `N、` 前缀（外层已有 `01`–`08` 章节号）。禁止把正文编号或无编号短语升级成一级标题。

交付前逐条确认 `SKILL.md` 的「交付前必检清单」（浮动按钮、UTF-8 base64 解码、干净 hero 背景、内联 SVG 箭头、布局图 1:1 忠实重画、html2canvas 限制、视频播放卡）。

## 运行脚本前的环境准备

脚本依赖 **`python-docx`**，当前机器的 `python3`（`/usr/bin/python3`）**没有安装**它。运行前先装：

```bash
python3 -m pip install --user python-docx
```

常用命令（均在 `docx-spec-html/` 下相对 skill 根目录运行）：

```bash
python3 scripts/extract_docx_manifest.py /path/to/source.docx --out /path/to/source-manifest.json
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output   # 可加 --updated 2026.06 / --editable / --strict
python3 scripts/validate_output.py /path/to/source.docx /path/to/final.html --strict
```

`batch_generate.py` 顶部导入 `from validate_output import validate`，所以必须从 `scripts/` 目录运行（或让该目录在 `sys.path` 上）。

## 交付物形态

- 默认交付**单文件** `.html`（CSS、图片、字体、编辑器、html2canvas 全部内嵌），除非用户明确要求分离 CSS。
- 保留 docx 里全部可见文本、全部图片出现次数、Word 表格的行列关系。
- **不 reflow**：块的顺序、标题↔图片的上下关系必须与源文档一致；层级重建只改样式/分组，不改阅读顺序。
- 每页右下角固定两个浮动按钮（「编辑」+「下载整页图片」），都带 `data-html2canvas-ignore`，保持自包含、不走 CDN。
- `--editable` 只用于用户要的可编辑审阅版，锁定终稿不要开。

## 更新 skill 的时机

用户反馈揭示的是**通用规则**时，更新 `references/design.md`（必要时改脚本）；反馈是**单文档特例**时，只改那一份输出或项目专属生成器。校验脚本历史上有两类误报已收紧判定（见 `SKILL.md` §8 引用块），不要为迎合旧误报去改结构。
