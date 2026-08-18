# Structure and content reconstruction

Use this reference while mapping PDF/DOCX material into page hierarchy. The PDF-resolved numbered module sequence and core preservation contract remain authoritative in `SKILL.md`; only the body-care profile has a fixed eight-module sequence.

## Contents

1. Source precedence
2. Block-order preservation
3. Hierarchy mapping
4. Text and media grouping
5. Table reconstruction

## 1. Source precedence

Use the PDF for final content hierarchy, merged cells, row/column relationships, and module boundaries. Use the OfficeCLI-backed DOCX manifest for extractable text, raw block order, outline levels, list levels, run formatting, image anchors/dimensions, table spans, and occurrence counts. When they disagree because PDF-to-DOCX export flattened structure, preserve DOCX content but restore PDF semantics.

Read these manifest fields before classifying blocks:

- `outline_level` and `list_level` for author-declared hierarchy;
- paragraph `runs[].format` for localized bold or emphasis rather than whole-paragraph guesses;
- `images[].path`, `width`, `height`, and `wrap` for image placement and grouping;
- table-cell `grid_column`, `rowspan`, `colspan`, and `format.vmerge` for grid reconstruction.

Treat OfficeCLI structure as strong evidence, not absolute truth. Confirm every promoted chapter and complex merged table against the companion PDF.

Do not import visual styling from the PDF. PDF colours, square borders, cell spacing, image padding, and typography are non-authoritative; use the canonical HTML stylesheet and golden reference for all visual treatment.

Do not infer a heading from numbering alone. Plain numbered paragraphs may be body rules. Do not infer a new chapter from a short unnumbered phrase; phrases such as `卖点选词优先级` normally remain inside their PDF module.

## 2. Block-order preservation

Emit source blocks in original order. Keep these relationships intact:

- title above image → `.label-line` above the image;
- image followed by caption → caption below that image;
- consecutive image-only paragraphs → one white module without reordering images;
- table title → table → following explanation, in that sequence.

Never pull later material forward to improve visual balance. Split or combine wrappers only when reading order remains unchanged.

## 3. Hierarchy mapping

- Page title → Hero `h1`, with `商品信息运营规范` forced to line two.
- Leading `标题：XXX` duplicate → Hero metadata only when `XXX` equals the document title; omit it from overview body copy.
- Overview first sentence → independent white `.lead`.
- Overview bracketed items → `.red-list`.
- Fixed module title → numbered `.card.spec-card` using `{ 标题 }`.
- Module-level `前缀：内容` → `.red-list`; split the prefix into `.label-text` and the remainder into `.label-rest`.
- Pure labels ending with `：` → `.label-line`.
- Nested list items → `.source-list`, grey squares, one-level indentation.
- Independent prose → separate `.plain-block`.
- Module-local `（1）（2）（3）（4）` subtitles → child white modules inside the same card, never new chapters.
- Consecutive `子标题 → 表格` pairs under one module-local subtitle → one shared white subtitle container; child titles use grey squares and one nested indent.
- Numbered first-level modules → accept the source PDF's exact order and count; cross-check their short names against the overview's `【模块名】` labels before promoting them.
- Explanatory children under a bracket parent such as `【主图】` → `.sublevel` items with grey squares, no pink marker, and the same group/text indentation as `.source-list` children.
- Keep a bracket parent and its `.sublevel` children as consecutive direct `<li>` siblings in one `.red-list`. Never create a nested `.red-list`; it compounds indentation and removes the parent-to-first-child spacing.
- Consecutive numbered siblings (`1、` / `2、` / `3、`) → one shared parent module, identical weight and grey-square hierarchy even when only one item contains a colon.

Wrap a label, its grey caption, and its images in the same `.text-block`. Indent the caption and images as children of that label.

## 4. Text and media grouping

Preserve repeated source text, not only unique strings. Never add explanatory copy that is absent from the source, except fixed interface labels such as `INTRODUCTION`, `编辑`, `下载整页图片`, and `点击播放`.

Keep images proportional by setting one dimension and leaving the other `auto`. Center images in their holders. Do not use fixed width and height together. Group five or more screen/detail examples in a two-column `.detail-screen-grid`, using `.span-full` only for wide or critical images.

When a reference layout diagram must be redrawn, inspect the actual image and reproduce its text and geometry exactly. Count one `.module-layout` as one represented source image.

## 5. Table reconstruction

Start from the OfficeCLI manifest's `grid_column`, `rowspan`, `colspan`, and `vmerge` evidence, then use the PDF to confirm or restore merged cells and true relationships lost in DOCX export:

- preserve headers and row/column correspondence;
- use `rowspan`/`colspan` where PDF cells are merged;
- keep all table text vertically and horizontally centered;
- keep lists inside a cell grouped and vertically centered;
- enlarge example images to match PDF proportions rather than retaining tiny export defaults;
- never flatten a semantic table into unrelated cards or a free image grid.
- apply no PDF-derived visual style; every table uses the canonical rounded-card system.

Choose the detailed component pattern from `components.md` only after identifying the PDF table semantics.
