# Structure and content reconstruction

Use this reference while mapping PDF/DOCX material into page hierarchy. The PDF-resolved numbered module sequence and core preservation contract remain authoritative in `SKILL.md`; only the body-care profile has a fixed eight-module sequence.

## Contents

1. Source precedence
2. Block-order preservation
3. Hierarchy mapping
   - 3.1 Block-pattern decision table (classify before choosing components)
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
- A `【…】` bracket line that appears while a colon label is open is a child of that label: keep it at the grey-square `.source-list` level (its `【…】：` prefix may stay bold), never promote it to a top-level `.red-list` item.
- Fixed module title → numbered `.card.spec-card` using `{ 标题 }`.
- Module-level `前缀：内容` → `.red-list`; split the prefix into `.label-text` and the remainder into `.label-rest`.
- Pure labels ending with `：` → `.label-line`.
- Nested list items → `.source-list`, grey squares, one rung of indentation. Lettered sub-items (`a.`/`b.`/`c.`) under a numbered label **drop one further rung** — hollow-square `li.deep`, indented again — and must never sit flush with the numbers above them. (This replaces the earlier rule that kept lettered items at the grey-square level.)
- Independent prose → separate `.plain-block`.
- Module-local `（1）（2）（3）（4）` subtitles → child white modules inside the same card, never new chapters. Exception: when those subtitles sit under a numbered parent label (such as `2.各区域详细规范说明：`) and each carries its own sub-list (`a.`/`b.`/`c.`), keep the parent label and all subtitle groups in ONE white container and render each subtitle as a grey-square `.caption-line`; only leaf-level `（1）`–`（4）` subtitles keep the pink marker.
- Consecutive `子标题 → 表格` pairs under one module-local subtitle → one shared white subtitle container; child titles use grey squares and one nested indent.
- Numbered first-level modules → accept the source PDF's exact order and count; cross-check their short names against the overview's `【模块名】` labels before promoting them.
- Explanatory children under a bracket parent such as `【主图】` → `.sublevel` items with grey squares, no pink marker, and the same group/text indentation as `.source-list` children.
- Keep a bracket parent and its `.sublevel` children as consecutive direct `<li>` siblings in one `.red-list`. Never create a nested `.red-list`; it compounds indentation and removes the parent-to-first-child spacing.
- Consecutive numbered siblings (`1、` / `2、` / `3、`) → one shared parent module, identical weight and grey-square hierarchy even when only one item contains a colon. However, a numbered item whose body is itself a top-level `前缀：内容` pair (e.g. `1. 标题结构：…`) belongs to `.red-list` like any other top-level colon item — the leading number stays inside the pink prefix and does not demote the item to `.source-list`.
- Strip a leading callout arrow before classifying: `👉1. xxx：…` is numbered sibling `1.`, not a parent label. A run of numbered siblings is ONE list at ONE level for the whole run — never render the first sibling (the one carrying `👉`) as a pink/red parent while `2.` / `3.` sink to a deeper grey level. Sibling-level decisions apply uniformly: all module-internal rules → one `.source-list`; all top-level `前缀：内容` items → one `.red-list`.

### 3.0 Marker ladder (decide the rung before the component)

Four rungs, in this order, with no skipping. A child that carries text is always exactly one rung below its parent — never the same rung, never two rungs down, and never unmarked.

| Rung | Marker | Classes | Indent |
| --- | --- | --- | --- |
| L1 | solid red square | `.red-list > li`, `.label-line`, `.example-line` | none (card gutter only) |
| L2 | solid grey square | `.source-list > li`, `li.sublevel`, `.caption-line` | `margin-left: var(--level-indent)` (28px) |
| L3 | hollow square (transparent fill, `2px solid #c9c9c9`) | `li.deep` | one more `var(--level-indent)` |
| L4 | system round dot | `li.dot`, `ul.dot-list > li` | one more `var(--level-indent)` |

Every rung keeps `padding-left: var(--marker-gutter)` (38px) so the marker and its text never collide. 28px ≈ one character at production scale, which is the requested “indent by about one character per level”.

Consequences that used to be violated:

- Text under a red square is L2 grey — not a hollow square, not an unmarked `<p>`.
- Text under a grey square is L3 hollow — not an unmarked `<p>`.
- Text under a hollow square is L4 round dot.
- `1./2./3.` under a red-square heading are L2; `a./b./c.` under those numbers are L3; a further run is L4.
- `dom_contracts.marker_ladder_checks` fails a page that skips a rung, leaves a marked child flush with its parent, or puts a hollow square directly under a red square.

Wrap a label, its grey caption, and its images in the same `.text-block`. Indent the caption and images as children of that label.

### 3.1 Block-pattern decision table (classify before choosing components)

Before styling ANY block, classify it against this table and record the mapping per block. Never pick a list component by gut feel — the patterns below are the recurring JD-spec shapes, and misclassifying them is the most common hierarchy defect.

| Source pattern (after stripping `👉`) | Component | Notes |
| --- | --- | --- |
| Standalone `效果数据：…` / `数据：…` / `数据效果：…` line whose value contains `%` / `％` / `PP` / `↑` / `↓` | `.metric-emphasis` green box | Never render as a bold list item, `.label-line`, or plain paragraph — including inside module bodies (not only card headers). Ranges like `+1.42-1.80%` still qualify. |
| Top-level numbered `N. 前缀：内容` items (whole run) | ONE `.red-list`; number stays inline before the pink-highlighted prefix | The entire run shares one list at one level. Numbered items that hang under a red-square heading are themselves the grey-square rung (indent one step, never flush-left with the red square). Sub-bullets under one numbered item (e.g. under `2. 内容要求：`) drop to hollow-square `li.deep` nested inside that `<li>` — they do not demote the run. |
| `👉`-prefixed (sub-)items under an open colon label such as `基础要求：` | `.source-list li.deep` — hollow square, deeper indent | The arrow only marks "this was a callout"; it does not change the item into a parent. Apply uniformly to every arrowed item in the run. |
| One bordered / shaded source box (usually `👉 标签` + numbered items + `建议` + `平台规则`/`操作手册` together) | Exactly ONE white `.text-block` holding all of its children in source order | A source box boundary is a card boundary. Do NOT fragment one source box into several `.text-block`s or into a stack of consecutive one-row/one-cell `.doc-table` wrappers (a single-cell `.doc-table` is not a table — it is a `.text-block`), and do NOT merge separate source boxes into one. |
| A row of tinted/coloured headers (e.g. pink `筛选面板`/`搜索商卡`/…) with one image under each header | One `.doc-table`: header row + one media row | This is a table in the source, not a loose image grid — never render as `.title-image-grid` / separate `.image-frame`s. |
| Loose screen examples with no per-image header | `.detail-screen-grid` / `.sample-image` / `.half-image` per the usual image-grouping rules | Only use these when the source gives no header row. |

Self-check before publishing: for every numbered run, every metric line, and every source box, ask "which decision-table row did I map this to?" If a block cannot be named, re-examine it against the PDF render instead of guessing.

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
