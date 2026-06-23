---
name: docx-spec-html
description: Generate production-quality MPDN50EU-style single-file HTML pages from one or many Word .docx specification documents. Use when the user provides DOCX files or folders and asks to generate, batch generate, validate, package, refine, or match the high-quality golden HTML standard; this skill requires model-led document hierarchy reconstruction, not just running the local draft generator.
---

# DOCX Spec HTML

## Purpose

Turn Word specification documents into polished, single-file HTML pages that match the bundled MPDN50EU visual system and the golden output quality. Preserve every visible text fragment, image occurrence, table relationship, and hierarchy.

## Resources

- `references/mpdn50eu-design.md`: full design system, CSS template, Word-to-HTML mapping rules, and acceptance checklist.
- `references/high-quality-workflow.md`: required high-quality workflow and common failure modes.
- `assets/examples/auto-oil-golden-output.html`: final golden example; use for structure and visual quality comparison.
- `assets/fonts/JINGDONGLangZhengTi1-Bold.ttf`: embedded hero-title font.
- `assets/mpdn50eu-styles.css`: complete reusable CSS template; embed it into final HTML rather than linking it unless the user explicitly asks for separate CSS.
- `scripts/extract_docx_manifest.py`: export paragraph/table/image manifest for model-led hierarchy decisions.
- `scripts/batch_generate.py`: baseline draft generator for one DOCX or a folder; supports optional in-page text editing with `--editable`.
- `scripts/validate_output.py`: validation helper for text, image count, table structures, and CSS invariants.

## Critical Rule

Do not deliver raw `batch_generate.py` output as final production work. It is only a draft and extraction aid. For best results, use the model to reconstruct the document hierarchy and patch or rewrite the HTML until it meets the golden quality bar.

## High-Quality Workflow

1. Read `references/high-quality-workflow.md` and `references/mpdn50eu-design.md`.
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

7. If the browser or user reports a reusable issue, update `references/mpdn50eu-design.md` before packaging or delivering.

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
- **Redraw the「首张主图模块化布局图」schematic with your own vision.** When a `…首张主图模块化布局图：` label is followed by a reference image, do NOT ship the raw watermarked screenshot. Use your multimodal image-reading ability (or an image-reading MCP) to read the image's ACTUAL on-image text and the real size/position of each module, then redraw it 1:1 into a `.module-layout` block (light-yellow fill, red border, portrait proportion) using the real on-image wording. The generator's name-driven `.module-layout` is only a no-vision fallback — replace it with your faithful redraw before delivery. The validator counts each `.module-layout` as one image (`redrawn_image_count`).

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
