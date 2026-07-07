---
name: docx-spec-html
description: Generate production-quality single-file HTML spec pages from one or many Word .docx specification documents. Use when the user provides DOCX files or folders and asks to generate, batch generate, validate, package, refine, or match the high-quality golden HTML standard. The source .docx is usually exported FROM a PDF, so its hierarchy is degraded (no heading styles, merged cells lost, module titles and body items both numbered); therefore this skill treats the companion PDF as the authoritative structure and requires model-led hierarchy reconstruction against that PDF, not just running the local draft generator.
---

# DOCX Spec HTML

## Purpose

Turn Word specification documents into polished, single-file HTML pages that match the bundled visual system and the golden output quality. Preserve every visible text fragment, image occurrence, table relationship, and hierarchy.

## Resources

- `references/design.md`: full design system, CSS template, Word-to-HTML mapping rules, generator conventions (§16.1.2), and acceptance checklist (§18).
- `references/high-quality-workflow.md`: required high-quality workflow and common failure modes.
- `agents/openai.yaml`: agent interface definition (display name, description, default prompt) for external tooling integration.
- `assets/examples/auto-oil-golden-output.html`: final golden example; use for structure and visual quality comparison.
- `assets/fonts/JINGDONGLangZhengTi1-Bold.ttf`: embedded hero-title font.
- `assets/styles.css`: complete reusable CSS template; embed it into final HTML rather than linking it unless the user explicitly asks for separate CSS.
- `assets/vendor/html-editor.html`: the bundled visual HTML editor. The generator embeds it (base64) into every page so the fixed 「编辑」 button can open the page in it; it can also be opened standalone to upload and edit any HTML.
- `assets/vendor/html2canvas.min.js`: offline self-contained html2canvas for the 「下载整页图片」 button; must be embedded, never linked via CDN.
- `scripts/extract_docx_manifest.py`: export paragraph/table/image manifest for model-led hierarchy decisions.
- `scripts/batch_generate.py`: baseline draft generator for one DOCX or a folder; supports optional in-page text editing with `--editable`.
- `scripts/validate_output.py`: validation helper for text, image count, table structures, and CSS invariants.

## Critical Rule

Do not deliver raw `batch_generate.py` output as final production work. It is only a draft and extraction aid. For best results, use the model to reconstruct the document hierarchy and patch or rewrite the HTML until it meets the golden quality bar.

## PDF Is The Source Of Truth

**The `.docx` is almost always exported from a PDF, and that export corrupts structure**: heading styles are flattened, merged table cells are lost, and both module titles (`1、主图规范`) and body list items (`1、视频画面需清晰…`) end up as the same plain numbered paragraphs. A pure Python heuristic cannot tell these apart, and cannot recover a 10×4 rowspan comparison matrix or an alternating grey/red attribute header.

Therefore the companion PDF, not the docx, is the authoritative structure.

- **Always ask the user for the matching PDF** and use it as the standard answer for hierarchy and table semantics. Users of this skill have confirmed they can supply a PDF every time.
- If no PDF is available, fall back to `manifest + model reading of the docx`, but warn that hierarchy accuracy will be lower.
- The pipeline is three stages, and the middle stage is done by the model, never skipped:
  1. **Script extracts raw material** (deterministic dirty work): unzip docx, pull text, export images, dump table cell data → `manifest.json`. `batch_generate.py` may also emit a rough draft.
  2. **Model reconstructs structure by reading the PDF**: decide the 8 first-level module titles vs. body numbering, rebuild merged/alternating tables, place captions. This is the step that was skipped before and caused a 10-defect delivery.
  3. **Model reviews the HTML screen-by-screen against the PDF** and fixes every mismatch before delivery.

## ⚠️ 交付前必检清单（每次生成都必须逐条确认）

以下项目是高频遗漏项，**无论是脚本生成还是模型手写 HTML，交付前都必须确认**：

### 1. 「编辑」+「下载整页图片」浮动按钮

每个生成的页面**必须**包含右下角固定的两个浮动按钮，缺一不可：
- **「下载整页图片」**（红色，最右）— 内嵌 `assets/vendor/html2canvas.min.js`，点击把 `.poster` 光栅化为 PNG 下载。
- **「编辑」**（白色，在下载按钮左侧）— 内嵌 `assets/vendor/html-editor.html`（base64），点击在新窗口打开编辑器。
- 两个按钮都必须加 `data-html2canvas-ignore`，使其不出现在下载的 PNG 中。
- 参考实现见 `scripts/batch_generate.py` 中的 `editor_runtime()` 和 `download_runtime()` 函数。

### 2. 编辑器 base64 解码必须使用 UTF-8

编辑器 HTML 文件包含中文，`atob()` 只能处理 Latin-1 单字节字符，直接用会导致中文乱码。**必须**用以下方式解码：

```javascript
var bin = atob(base64String);
var bytes = new Uint8Array(bin.length);
for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
var html = new TextDecoder('utf-8').decode(bytes);
```

**禁止**直接 `atob()` 后当字符串使用。`batch_generate.py` 已内置此处理，手写 HTML 时也必须照做。

### 3. Hero 区干净背景

Hero 背景**只用**纯品牌红单色 `#FF2B22`（不要渐变）。**禁止**输出以下装饰元素：
- ❌ `.hero::before` 网格纹理 / 斜纹
- ❌ `.robot-deco` / `.ring` / `.ring-one` / `.ring-two` / `.ring-three` 圆环
- ❌ `.path-line` 路径线条
- ❌ `radial-gradient` 高光点

### 4. 箭头必须用内联 SVG

`.en-label span`（红、向左）与 `.hero-rule`（白、向右）的连体线+钩箭头**必须**用内联 `<svg>` 元素（`EN_LABEL_ARROW_SVG`/`HERO_RULE_ARROW_SVG`），`background:none`。html2canvas 不渲染 `data:image/svg` 背景图，用 CSS background 会导致下载 PNG 中箭头消失。**凡需进入下载图的图形，一律用真实元素（`<img>`/`<svg>`），别用 background。**

### 5. 首张主图模块化布局图必须 1:1 忠实重画

当出现 `…首张主图模块化布局图：` 标签且其后有参考图时，**绝不**直接嵌入带水印的原图，也**绝不**凭名字瞎编布局。必须：
1. 用多模态识图能力真正"看"那张参考图
2. 逐块抄下图上原文（一字不差）
3. 量出每块的相对位置和面积
4. 用 `.ml-grid` + `.ml-block` + 内联 `grid-template-columns`/`grid-column`/`grid-row` 精确还原
5. 每块一种明显区分的颜色

**⛔ 不存在"标准京东模板"。每张布局图结构都不同，必须以参考图为唯一依据。** 图上没有的词一个都不许加，图上有的词必须逐字照抄。校验器把每个 `.module-layout` 计为 1 张图片。

### 6. html2canvas 限制与 workaround

- **不支持 `object-fit`**：凡要进入下载图的内容图，不要靠 `object-fit:contain` + `max-height` 保比例。统一用"定一边、另一边 auto"。
- **不支持 `mix-blend-mode`**：Hero overlay 在截图前必须用 `flattenHeroOverlay(doc)` 预合成为扁平 PNG。
- **不处理 `<video>`**：截图前 `swapVideosForPosters` 把 `<video>` 临时换成 `<img>`（优先用 `poster`）。内嵌视频必须带 `poster`。

### 7. 主图视频播放卡

在 `{ 主图视频 }` 模块/章节里，出现任何"可点击观看视频"的信号（明确文案、视频链接、点击提示词），**必须**替换为播放卡（淡粉底 `#fbe2ec`、文案统一为 `点击播放`、深灰播放图标）。绝不可把视频链接当普通文字或 `<a>` 保留。同一章节多条提示只渲染一张卡。

## High-Quality Workflow

1. **Obtain the companion PDF (source of truth).** Ask the user for the matching PDF and read it — it carries the correct hierarchy, merged cells, and colour coding that the docx lost. Only fall back to docx-only reconstruction if no PDF exists (and warn about lower accuracy).
2. Read `references/high-quality-workflow.md` and `references/design.md`.
3. Extract a source manifest:

```bash
python3 scripts/extract_docx_manifest.py /path/to/source.docx --out /path/to/source-manifest.json
```

4. Generate a baseline draft (raw material / draft only — never the deliverable):

```bash
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output
```

Use Chinese hero update text. The generator defaults to the source file's last-modified date in `更新日期 YYYY年M月D日` format. To specify a date:

```bash
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output --updated 2026.06
# or with day precision:
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output --updated 2026.06.22
```

To add an in-page text editing toolbar for review copies only:

```bash
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output --editable
```

5. **Reconstruct the hierarchy against the PDF (model-led — this is the core step, never skip it).** Cross-read the PDF, manifest, and baseline HTML to decide:
   - the fixed first-level module sequence (see "Fixed First-Level Module Sequence" below) — strip the `N、` prefix, and demote any body numbering the draft wrongly promoted;
   - hero title and overview card;
   - grey-panel labels versus lower-level grey-dot items;
   - images and image captions;
   - true row/column correspondence, including merged cells and alternating headers the PDF shows but the docx flattened (see "Complex Table Reconstruction" below);
   - examples, before/after comparisons, and grids.
6. Patch or rewrite the generated HTML until the visual hierarchy matches both the PDF and `assets/examples/auto-oil-golden-output.html`.
7. **Review screen-by-screen against the PDF.** Walk the PDF top to bottom and confirm every module title, table, merge, colour, and caption matches; fix each mismatch before delivery.
8. Validate:

```bash
python3 scripts/validate_output.py /path/to/source.docx /path/to/final-output.html --strict
```

> **验证脚本的两类历史误报（已修复，勿再手动改结构去迎合）**：
> 早期 `validate_output.py` 会对结构正确的页面误报两条 warning，现已收紧判定，理解其原因可避免被误导：
> - **`image_cards_use_image_holders`**：`hero-overlay` 是生成器注入的可替换背景层，本就**不该**包在 `.image-holder` 里。脚本现按 `image_holder_count >= expected_images - hero_overlay_count` 判定，放行这 1 张例外。**不要**为了凑数把 hero 背景塞进 image-holder。
> - **`caption_images_keep_text_below_images`**：仅当出现 `[图A] + [图B] = 效果` 这类**带等号结果**的真正图注公式时才要求 `.caption-image-card`。长标题结构语法（如 `[品牌/系列] + 产品词 + 规格`）虽含方括号和加号，但**不是图注**，不得触发该规则。

9. If the browser or user reports a reusable issue, update `references/design.md` before packaging or delivering.

## Fixed First-Level Module Sequence

商品信息运营规范文档的一级模块是**固定且有限的**。文档正文里模块标题写成 `1、主图规范`…`8、属性`，而正文要点也用 `1、2、3、` 编号——脚本无法区分，必须由模型按 PDF 锚定。规则：

- **一级标题只能是这 8 个模块**（顺序固定，用 `01`–`08` 章节号）：`主图规范` / `主图视频` / `长标题` / `短标题` / `通用卖点` / `主推标签` / `品质标签` / `属性`。
- **去掉标题里的 `N、` 前缀**——外层已经有 `01`–`08` 章节号，`1、主图规范` 只渲染成 `{ 主图规范 }`。
- **禁止把正文编号升级为一级标题**。例如 `1、视频画面需清晰…`、`2、视频需要搭配字幕…`、`3、时长…` 在 PDF 里明明在「主图视频」模块下面，是正文 123 点，必须留�该模块卡片内，不能变成新章节。
- **禁止把无编号短语升级为一级标题**。例如 `卖点选词优先级` 属于「通用卖点」模块下的子项，不是第 9 个模块。
- 模块内的 `（1）（2）（3）（4）` 小节（如「首图模块化」「首图内容优化方向」）是**该模块卡片内的子标题**，连同其表格/图片一起放进同一张白卡，不得单独成章节。

## Complex Table Reconstruction

PDF 里合并单元格清晰可见，但 docx 导出后合并信息丢失、单元格被拍平成线性文本。`before_after / spec / generic` 三类不足以还原。模型看 PDF 后按语义选择渲染方式：

- **多列对比矩阵（优化内容 | 案例 | 优化前 | 优化后）**：如「首图内容优化方向」表。每个「优化内容」跨多行合并（`rowspan`），右侧对应「案例1 / 案例2 / 优化总结」多行；「优化前 / 优化后」在「优化总结」行内成对出现。按 PDF 的分组和合并结构重建，不得拆成独立竖列。
- **表头合并**：相邻表头文字相同时合并成一格（如「示例 | 示例」合并为一格，其下双列图片；「展现样式」跨多列）。
- **交替配色表头**：属性/卖点方向表（如「适用肤质/香型」「功效/成分」「适用人群/净含量」交替行）按 `design.md` 的配色规则上灰底/红底，不能全用同色。
- **内容居中**：纯对比/枚举型表格（如「信息不一致案例」）单元格内容水平+垂直居中。
- **图片宽度**：示例图偏小的表格（如「品质标签示例」）图片按 PDF 视觉比例放大（可到等宽/翻倍宽），不要保留脚本默认的小图。

## Batch Workflow

For many DOCX files:

1. Run `batch_generate.py` on the folder to create drafts and reports.
2. Sort by report warnings and document complexity.
3. Use model-led refinement on each page before delivery.
4. Re-run validation per final HTML.

Folder command:

```bash
python3 scripts/batch_generate.py /path/to/input-folder /path/to/output-folder
```

## Non-Negotiable Output Rules

- Deliver one self-contained `.html` file by default; no separate CSS unless explicitly requested.
- Embed CSS, images, and title font.
- Preserve all visible DOCX text, including repeated text.
- Preserve all DOCX image occurrences, including repeated images.
- **Keep the document's original block order and the title↔image relationship — do NOT reflow.** Emit blocks in the same order they appear in the DOCX. A "标题 → 图片 → 说明文字" sandwich must stay in that exact order. A title that sits **above** an image in the source must stay **above** that image in the HTML (as the image's `.label-line` caption) — never move it below the image; a caption that sits below an image stays below. Never pull later content earlier, never split a title from the image it introduces, and never swap their positions. Hierarchy reconstruction changes *styling/grouping*, not the reading order.
- Preserve true Word tables as table-like layouts with row/column relationships.
- Use exactly `{ 标题 }` for section titles, with one space inside both braces. Never use double braces such as `{{ 标题 }}`.
- Use `INTRODUCTION` for every card title right label.
- Force the hero title's `商品信息运营规范` category suffix onto the second line (insert `<br>` before it).
- Use `JINGDONGLangZhengTi1-Bold` for the hero title.
- Use `MiSans-Bold` for card section titles.
- Keep the hero upper-right bracket mark `.hero-mark` with `OPERATION` / `STANDARDS`.
- Use Chinese hero date text: `更新日期 YYYY年M月D日` (default to source file modification date; month/day no zero-padding).
- Keep `.gray-panel > * + * { margin-top: 18px; }` to prevent white modules touching or overlapping.
- Wrap grey-panel content in approved modules (`.text-block`, `.image-frame`, `.caption-image-card`, `.word-table-spec`, grids). Do not place bare labels, lists, or images directly under `.gray-panel`.
- Do not promote image captions or formula captions into higher-level titles unless the source hierarchy requires it.
- For five or more consecutive screen/detail examples, use `.detail-screen-grid`: two columns by default, `.span-full` only for wide or critical images. Do not use four-column `.screens-grid` for these examples.
- Keep the editing toolbar optional. Do not enable `--editable` for final locked deliverables unless the user asks for editable review output.
- Every generated page carries a fixed "下载整页图片" button (bottom-right) that rasterises the whole poster to one PNG via the embedded `assets/vendor/html2canvas.min.js`. Keep it self-contained — do not switch it to a CDN `<script src>`.
- Every generated page also carries a fixed "编辑" button (to the left of "下载整页图片"). The bundled `assets/vendor/html-editor.html` is embedded as inert base64; clicking 编辑 opens the editor in a new window pre-loaded with the current page (`window.__PRELOAD_HTML__`). Both floating buttons carry `data-html2canvas-ignore` so they never appear in the downloaded PNG. Keep this self-contained — do not link the editor as a separate file or CDN.
- **综述首段独立白底**：综述卡的第一段总述句单独用 `.lead`（白底圆角），与下方要点区分开。
- **`【…】` 要点转红色列表**：以 `【主图】`、`【标题】` 等方括号开头的并列要点统一进 `.red-list`，方括号词作为 `<b>`（红色方块 + 荧光条），其后描述紧随其后。
- **顶层「前缀：内容」转红色列表**：章节里顶层形如 `总结：…`、`字数范围：…` 的行须进 `.red-list`，冒号前缀作为 `<b>`（红色方块 + 荧光条）。以冒号**结尾**的纯标签（如 `结构：`）仍走 `.label-line`。
- **标签 + 示例图合并同一白底，示例图下级缩进**：标签（`.label-line`）、灰色说明 `.caption-line`（只用灰方块、不加红方块和荧光条）与图片须包在**同一个** `.text-block` 白底里；`.caption-line` 与图片右缩进 `28px`（`.indent`）。
- **优化前/优化后对比**：用 `.ba-compare` 两列，`优化前`表头底色为灰 `.ba-before`，`优化后`为品牌红 `.ba-after`。
- **三列规范表**：`主图 / 内容要求 / 示例` 三列表用 `.spec-table`（`grid-template-columns: 1fr 2fr 3fr`）；数据单元格内容**水平＋垂直都居中**；`示例`列图片 `width:100%` 保持等宽。
- **转化率绿色强调条**：`XX率+X%` 指标用 `.metric-emphasis`（白底 + 绿色描边 `#47b250`、圆角 14px）；`+X%` 用绿色 `.metric-value` 放大约两倍，其后紧跟绿色上扬箭头 `.metric-arrow`（内嵌 SVG）。独立成行时左右撑满灰面板；嵌在标签内时从标签中拆出来，置于 `.label-line` 标题下方，宽度撑满白底内容区。
- **图文详情半幅图**：`图文详情` 卡内示例图用 `.half-image`（`max-width:600px`）并居中。
- **同一标签下 ≥2 张并排示例图半宽等宽**：holder 加 `.sample-image`（`doc-image width:50%`）；单张示例图保持原大图、不缩半。
- **红方块标题后的内容统一字号 28px**：`.source-list`、`.caption-line`、`.example-line` 全部 28px；仅表格类保留各自更紧的字号体系。
- **连续纯图段落合并**：相邻只有图片没有文字的段落合并进同一个 `.text-block` 模块。
- **卡片正文整体放大约 1.5×**：`.lead`/`.red-list`/`.label-line`/`.source-list`/`.plain-block p`/`.example-line`/`.caption-line`≈28px；hero 区主标题 `h1` 68→102px、右上 `OPERATION STANDARDS` 14→21px（连同括号胶囊框一并放大）、右下更新日期 18→27px，`.hero` 增高至 600px。**卡片标题栏 `.section-head` 与转化率绿框 `.metric-emphasis` 不参与放大。** 实现：GENERIC_CSS 末尾用 `.poster.auto-doc .类名` 覆盖基础字号。
- **Hero overlay 可替换背景图层**：Hero 里固定输出 `<img class="hero-overlay">`（默认 1×1 透明 PNG），绝对定位铺满 hero，`mix-blend-mode:overlay` 与红底混合叠加，`z-index:0` 在文字之下。编辑器里双击 hero 空白处可替换图片，新图与红色叠加而非不透明盖住。下载前 `flattenHeroOverlay` 预合成。
- **粉色荧光条只盖标题、与冒号后内容分离**：红方块标题里带冒号的，第一个冒号之前是标题（`.label-text`，红方块＋粉条），冒号之后的内容拆到下一行 `.label-rest`（无方块无粉条）。没有冒号的标题只给红方块、不要粉条（`.label-plain`）。
- **章节直下的段落**：有项目符号→红方块项；无符号的独立大段→各自白底 `.plain-block` 容器。
- **更新日期与箭头底对齐**：GENERIC_CSS 用 `.poster.auto-doc .updated{bottom:165px}` 覆盖基础的 `bottom:116px`（按 hero 600px + 强制两行标题量得）。
- **绝不要在 GENERIC_CSS 里重写 hero 子元素的 `position`**：`.hero h1/.updated/.hero-mark/.hero-rule` 已有 `position:absolute` + `z-index:1`，覆盖 position 会导致排版错乱。

## Quality Bar

The final page should look closer to `assets/examples/auto-oil-golden-output.html` than to the raw batch draft. In particular:

- section titles use the correct red type, braces, chapter number, and `INTRODUCTION` label;
- the hero upper-right `OPERATION STANDARDS` bracket mark is present;
- the hero update date is Chinese, for example `更新日期 2026年6月22日`;
- the `INTRODUCTION` label sits above an inline SVG arrow (red, hook pointing left); the hero rule uses the same arrow in white, flipped (hook pointing right);
- grey panels contain white modules with stable spacing;
- `.label-line` uses red square plus the semi-transparent brand-red highlight bar `rgba(255,43,34,0.2)` only behind text;
- lower-level items use grey squares and consistent indentation;
- screen/detail examples are not cramped; long runs use the two-column `.detail-screen-grid`;
- images keep proportion and are centered;
- table columns align consistently across rows;
- no text overlaps, clipped content, or white-card collisions remain;
- card body text is ~1.5× enlarged (28px base) per the golden example's visual proportions;
- hero background is clean solid `#FF2B22` with no decorative overlays (overlay image via `.hero-overlay` is acceptable);
- all arrows in downloaded PNG are visible (inline SVG, not CSS background).

## When To Update The Skill

When user feedback identifies a general rule, update the design reference and, if needed, the scripts. When feedback is document-specific, patch only that output or a project-specific generator.

## Editable Review Mode

Editable mode adds a floating toolbar that toggles `contenteditable` on text nodes and downloads the modified HTML. It does not update the source DOCX, and browser security prevents silently overwriting the original local file. After a reviewer edits text and downloads a new HTML file, run `validate_output.py` again if source fidelity still matters.