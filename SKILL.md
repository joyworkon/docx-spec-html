---
name: docx-spec-html
description: Generate production-quality single-file HTML spec pages from one or many Word .docx specification documents. Use when the user provides DOCX files or folders and asks to generate, batch generate, validate, package, refine, or match the high-quality golden HTML standard; this skill requires model-led document hierarchy reconstruction, not just running the local draft generator.
---

# DOCX Spec HTML

## Purpose

Turn Word specification documents into polished, single-file HTML pages that match the bundled visual system and the golden output quality. Preserve every visible text fragment, image occurrence, table relationship, and hierarchy.

## Resources

- `references/design.md`: full design system, CSS template, Word-to-HTML mapping rules, and acceptance checklist.
- `references/high-quality-workflow.md`: required high-quality workflow and common failure modes.
- `assets/examples/auto-oil-golden-output.html`: final golden example; use for structure and visual quality comparison.
- `assets/fonts/JINGDONGLangZhengTi1-Bold.ttf`: embedded hero-title font.
- `assets/styles.css`: complete reusable CSS template; embed it into final HTML rather than linking it unless the user explicitly asks for separate CSS.
- `assets/vendor/html-editor.html`: the bundled visual HTML editor. The generator embeds it (base64) into every page so the fixed 「编辑」 button can open the page in it; it can also be opened standalone to upload and edit any HTML.
- `scripts/extract_docx_manifest.py`: export paragraph/table/image manifest for model-led hierarchy decisions.
- `scripts/batch_generate.py`: baseline draft generator for one DOCX or a folder; supports optional in-page text editing with `--editable`.
- `scripts/validate_output.py`: validation helper for text, image count, table structures, and CSS invariants.

## Critical Rule

Do not deliver raw `batch_generate.py` output as final production work. It is only a draft and extraction aid. For best results, use the model to reconstruct the document hierarchy and patch or rewrite the HTML until it meets the golden quality bar.

## ⚠️ 交付前必检清单（每次生成都必须逐条确认）

以下三项是高频遗漏项，**无论是脚本生成还是模型手写 HTML，交付前都必须确认**：

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

这些装饰在旧版模板中存在，但设计规范（`design.md` 的「头图干净背景」条目）已明确移除。生成 HTML 时 markup 和 CSS 中都不要包含它们。

## High-Quality Workflow

1. Read `references/high-quality-workflow.md` and `references/design.md`.
2. Extract a source manifest:

```bash
python3 scripts/extract_docx_manifest.py /path/to/source.docx --out /path/to/source-manifest.json
```

3. Generate a baseline draft:

```bash
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output
```

Use Chinese hero update text. The generator defaults to the current month in `更新日期 YYYY年MM月` format. To specify a month:

```bash
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output --updated 2026.06
```

To add an in-page text editing toolbar for review copies only:

```bash
python3 scripts/batch_generate.py /path/to/source.docx /path/to/output --editable
```

4. Use the manifest, DOCX structure, baseline HTML, and golden example to make model-led layout decisions:
   - hero title and overview card;
   - numbered chapter cards;
   - grey-panel labels versus lower-level grey-dot items;
   - images and image captions;
   - true Word table row/column correspondence;
   - examples, before/after comparisons, and grids.
5. Patch or rewrite the generated HTML until the visual hierarchy matches `assets/examples/auto-oil-golden-output.html`.
6. Validate:

```bash
python3 scripts/validate_output.py /path/to/source.docx /path/to/final-output.html --strict
```

7. If the browser or user reports a reusable issue, update `references/design.md` before packaging or delivering.

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
- Use Chinese hero date text: `更新日期 YYYY年MM月`.
- Keep `.gray-panel > * + * { margin-top: 18px; }` to prevent white modules touching or overlapping.
- Wrap grey-panel content in approved modules (`.text-block`, `.image-frame`, `.caption-image-card`, `.word-table-spec`, grids). Do not place bare labels, lists, or images directly under `.gray-panel`.
- Do not promote image captions or formula captions into higher-level titles unless the source hierarchy requires it.
- For five or more consecutive screen/detail examples, use `.detail-screen-grid`: two columns by default, `.span-full` only for wide or critical images. Do not use four-column `.screens-grid` for these examples.
- Keep the editing toolbar optional. Do not enable `--editable` for final locked deliverables unless the user asks for editable review output.
- Every generated page carries a fixed "下载整页图片" button (bottom-right) that rasterises the whole poster to one PNG via the embedded `assets/vendor/html2canvas.min.js`. Keep it self-contained — do not switch it to a CDN `<script src>`.
- Every generated page also carries a fixed "编辑" button (to the left of "下载整页图片"). The bundled `assets/vendor/html-editor.html` is embedded as inert base64; clicking 编辑 opens the editor in a new window pre-loaded with the current page (`window.__PRELOAD_HTML__`). Both floating buttons carry `data-html2canvas-ignore` so they never appear in the downloaded PNG. Keep this self-contained — do not link the editor as a separate file or CDN.
- **Redraw the「首张主图模块化布局图」schematic 1:1 — match the reference image, do NOT invent a layout.** When a `…首张主图模块化布局图：` label is followed by a reference image, NEVER ship the raw watermarked screenshot and NEVER guess a layout from the module names. The redraw must reproduce the reference image's **proportions, row/column split, block positions/spans, and verbatim on-image text** — only the watermark is dropped and each block gets a distinct colour. Follow the strict redraw protocol in `references/design.md` (open and actually *look* at the image with your multimodal/vision ability → copy each block's text character-for-character → measure each block's relative position and area → place blocks on the `.ml-grid` with inline `grid-template-columns` / `grid-column` / `grid-row` so the proportions match 1:1 → distinct colour per block, centered text). The container fixes only the light-yellow fill + red border + portrait `max-width`; everything else comes from the image. The generator's name-driven `.module-layout` is only a no-vision fallback — replace it with your faithful redraw before delivery. The validator counts each `.module-layout` as one image (`redrawn_image_count`).

## Quality Bar

The final page should look closer to `assets/examples/auto-oil-golden-output.html` than to the raw batch draft. In particular:

- section titles use the correct red type, braces, chapter number, and `INTRODUCTION` label;
- the hero upper-right `OPERATION STANDARDS` bracket mark is present;
- the hero update date is Chinese, for example `更新日期 2026年06月`;
- the `INTRODUCTION` label sits above a one-piece curved-hook SVG arrow (red, hook pointing left); the hero rule uses the same arrow in white, flipped (hook pointing right);
- grey panels contain white modules with stable spacing;
- `.label-line` uses red square plus the semi-transparent brand-red highlight bar `rgba(255,43,34,0.2)` only behind text;
- lower-level items use grey squares and consistent indentation;
- screen/detail examples are not cramped; long runs use the two-column `.detail-screen-grid`;
- images keep proportion and are centered;
- table columns align consistently across rows;
- no text overlaps, clipped content, or white-card collisions remain.

## When To Update The Skill

When user feedback identifies a general rule, update the design reference and, if needed, the scripts. When feedback is document-specific, patch only that output or a project-specific generator.

## Editable Review Mode

Editable mode adds a floating toolbar that toggles `contenteditable` on text nodes and downloads the modified HTML. It does not update the source DOCX, and browser security prevents silently overwriting the original local file. After a reviewer edits text and downloads a new HTML file, run `validate_output.py` again if source fidelity still matters.
