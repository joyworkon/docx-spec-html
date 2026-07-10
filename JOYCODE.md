# JOYCODE.md

本仓库的核心是一个 JoyCode Skill：**`docx-spec-html`** —— 把京东「商品信息运营规范」类 Word `.docx` 文档转成生产级单文件 HTML 长图页。`GeneratedProducts/` 存放用该 skill 生成的实际产物。

## 仓库结构

- `docx-spec-html/` —— skill 本体（`SKILL.md` 是权威说明，动手前必读）
  - `SKILL.md` —— 核心工作流、固定八模块、交付必检与不可协商规则
  - `references/structure.md` —— PDF/DOCX 层级重建与内容顺序
  - `references/components.md` —— 复杂表格、图片布局、指标与视频组件
  - `references/visual-qa.md` —— 浏览器、导出与交付验收
  - `scripts/` —— `extract_docx_manifest.py`（抽料）、`batch_generate.py`（草稿生成器）、`validate_output.py`（校验）
  - `assets/` —— 唯一 CSS 源、WOFF2 hero 字体、轻量 golden 参考与 `vendor/` 运行时
  - `agents/openai.yaml` —— 外部工具集成的 agent 接口定义
- `GeneratedProducts/` —— 生成产物；`*-rebuild/` 目录里的 `probe_*.py`/`patch_*.py`/`inspect_*.py` 是一次性调试脚本，不是可复用工具
- `.joycode/memory/` —— 跨会话记忆（已记录「PDF 为主料」这条核心反馈）

## 最重要的规则（违反过会被用户挑错）

1. **PDF 才是结构标准答案，不是 docx。** `.docx` 几乎都是从 PDF 导出的，导出破坏了层级：无 heading 样式、合并单元格丢失、模块标题（`1、主图规范`）和正文要点（`1、视频画面需清晰`）都变成同样的编号段落。每次任务都要向用户索取配套 PDF。
2. **绝不把 `batch_generate.py` 的原始输出当成品交付。** 它只是草稿和抽料工具。正确流程是三段式：脚本抽料 → 模型读 PDF 重建层级 → 对照 PDF 逐屏审查修正。中间那步（模型重建）由模型做，永远不能跳过——跳过它曾导致 10 处结构错误。
3. **8 个一级模块固定**：主图规范 / 主图视频 / 长标题 / 短标题 / 通用卖点 / 主推标签 / 品质标签 / 属性。渲染时去掉标题里的 `N、` 前缀（外层已有 `01`–`08` 章节号）。禁止把正文编号或无编号短语升级成一级标题。

交付前逐条确认 `SKILL.md` 的「交付前必检清单」（浮动按钮、UTF-8 base64 解码、干净 hero 背景、内联 SVG 箭头、布局图 1:1 忠实重画、html2canvas 限制、视频播放卡）。

## 运行脚本前的环境准备

脚本依赖 **`python-docx`**。环境缺少依赖时安装：

```bash
python3 -m pip install --user python-docx
```

常用命令（均在 `docx-spec-html/` 下相对 skill 根目录运行）：

```bash
python3 scripts/extract_docx_manifest.py /path/to/source.docx --out /path/to/source-manifest.json
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output   # 可加 --style custom.css / --updated 2026.06 / --editable / --strict
python3 scripts/validate_output.py /path/to/source.docx /path/to/final.html --strict
```

`--design` 继续作为 `--style` 的兼容别名，并可读取旧版含 CSS 代码块的 Markdown 设计文件。

## 交付物形态

- 默认交付**单文件** `.html`（CSS、图片、字体、编辑器、html2canvas 全部内嵌），除非用户明确要求分离 CSS。
- 保留 docx 里全部可见文本、全部图片出现次数、Word 表格的行列关系。
- **不 reflow**：块的顺序、标题↔图片的上下关系必须与源文档一致；层级重建只改样式/分组，不改阅读顺序。
- 每页右下角固定两个浮动按钮（「编辑」+「下载整页图片」），都带 `data-html2canvas-ignore`，保持自包含、不走 CDN。
- `--editable` 只用于用户要的可编辑审阅版，锁定终稿不要开。

## 更新 skill 的时机

用户反馈揭示的是**通用规则**时，按内容更新 `references/structure.md`、`references/components.md` 或 `references/visual-qa.md`（必要时改脚本/CSS）；反馈是**单文档特例**时，只改那一份输出或项目专属生成器。不要把 CSS 复制回 Markdown，`assets/styles.css` 是唯一样式源。
