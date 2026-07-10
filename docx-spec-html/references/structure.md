# Structure and content reconstruction

Use this reference while mapping PDF/DOCX material into page hierarchy. The fixed eight-module sequence and core preservation contract remain authoritative in `SKILL.md`.

## Contents

1. Source precedence
2. Block-order preservation
3. Hierarchy mapping
4. Text and media grouping
5. Table reconstruction

## 1. Source precedence

Use the PDF for visual hierarchy, merged cells, colour coding, and module boundaries. Use the DOCX for extractable text, image blobs, occurrence counts, and raw block order. When they disagree because PDF-to-DOCX export flattened structure, preserve DOCX content but restore PDF semantics.

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
- Overview first sentence → independent white `.lead`.
- Overview bracketed items → `.red-list`.
- Fixed module title → numbered `.card.spec-card` using `{ 标题 }`.
- Module-level `前缀：内容` → `.red-list`; split the prefix into `.label-text` and the remainder into `.label-rest`.
- Pure labels ending with `：` → `.label-line`.
- Nested list items → `.source-list`, grey squares, one-level indentation.
- Independent prose → separate `.plain-block`.
- Module-local `（1）（2）（3）（4）` subtitles → child white modules inside the same card, never new chapters.

Wrap a label, its grey caption, and its images in the same `.text-block`. Indent the caption and images as children of that label.

## 4. Text and media grouping

Preserve repeated source text, not only unique strings. Never add explanatory copy that is absent from the source, except fixed interface labels such as `INTRODUCTION`, `编辑`, `下载整页图片`, and `点击播放`.

Keep images proportional by setting one dimension and leaving the other `auto`. Center images in their holders. Do not use fixed width and height together. Group five or more screen/detail examples in a two-column `.detail-screen-grid`, using `.span-full` only for wide or critical images.

When a reference layout diagram must be redrawn, inspect the actual image and reproduce its text and geometry exactly. Count one `.module-layout` as one represented source image.

## 5. Table reconstruction

Use the PDF to restore merged cells and true relationships lost in DOCX export:

- preserve headers and row/column correspondence;
- use `rowspan`/`colspan` where PDF cells are merged;
- keep pure comparison/enumeration cells horizontally and vertically centered;
- keep lists inside a cell grouped and vertically centered;
- enlarge example images to match PDF proportions rather than retaining tiny export defaults;
- never flatten a semantic table into unrelated cards or a free image grid.

Choose the detailed component pattern from `components.md` only after identifying the PDF table semantics.
