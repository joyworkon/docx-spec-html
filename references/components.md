# Complex components

Read this file only when the source contains the corresponding component. Exact dimensions and visual styling live in `assets/styles.css`.

## Contents

1. Tables and comparisons
2. Image layouts
3. Metrics and labels
4. Video and export runtime

## 1. Tables and comparisons

### Multi-row optimization matrix

Use `.compare-matrix` for `优化内容 | 案例 | 优化前 | 优化后`. Merge each optimization group with `rowspan`; keep case rows and summary rows aligned. Put paired before/after images in the corresponding row. Image cells use `.cm-img` with the canonical equal inset and proportional full-width images.

### Material-type table

Use `.material-table` for `素材图类型 | 内容要求 | 示例`. Let the 示例 header span two image columns. Use four tracks at approximately `16% / 34% / 25% / 25%`; keep the two examples equal width with the same inset on all four sides.

### Attribute or keyword-direction table

Use `.attr-table` for paired directions. Default every header to the same brand red. Use an alternate header colour only when the source explicitly assigns a different semantic role; PDF colour variation alone is not evidence. In body-care documents, `适用肤质/香型 | 功效/成分 | 适用人群/净含量` belongs to 通用卖点, not 属性.

### Standard tables

- Use `.spec-table` for `主图 | 内容要求 | 示例`, with approximately `1fr 2fr 3fr` tracks; center body copy and top headers horizontally and vertically.
- Use `.word-table-spec` for other three-column semantic tables. Share one column track across all rows.
- Use `.ba-compare` for two-column 优化前/优化后 material; grey header before, red header after. Always wrap `.ba-compare` in a `.text-block` white container, even when no label precedes it — the `.ba-before` header's light-grey fill is only visible against the white card, never directly on the grey panel. Media cards stretch to the column height (backgrounds stay level; images keep source proportion and stay centred). When the source gives 优化前/优化后 a shared header spanning several image columns (colspan), keep ONE column per header and place its images side by side in `.ba-media-row`; never repeat the header per image.
- Merge repeated adjacent headers with `colspan` when the PDF shows one shared heading.
- Use `.tag-example-table` when a tag example must appear larger than the draft default.
- A tag example is always `.doc-table.tag-example-table` inside `.doc-table-wrap`; its first row contains non-empty `<th>` cells and every image body cell uses `.table-media-cell > .image-holder > img`. The variant class never replaces the required `.doc-table` base class.
- All semantic tables share one visual tier: separate rounded cells, 8px gaps, 10px corners, centred 24px body copy, uniform light-grey `#f7f7f7` body cells, centred red 24px/700 top headers, and 12px equal inset around every image. Cell text longer than 10 visible characters is the exception to centring: it carries `.justify-txt` and renders justified with a left-aligned last line.
- Use `.text-block.subtitle-table-group` when a local subtitle owns consecutive `child label → table` pairs. Each `.nested-table-group` stays inside that white container and uses `.caption-line` for its grey-square child heading.
- Use `.tag-example-table` for equal-column image showcases such as `展现样式` and `前台展示案例`; give every image in the showcase row one shared rendered height while keeping width automatic and source proportions intact. The generator derives that shared height from the real pixel aspects and the fixed column track (`--tag-example-media-height` inline on the table), so images never spill out of their cell cards; `max-width: 100%` remains as the hard clamp. Exception: when the row's image aspect ratios diverge strongly (widest/narrowest ratio above ~1.5), the generator adds the `.is-fit-width` variant class so the row switches to equal-width columns with fully proportional images (heights may differ, mirroring the source).
- Apply the media inset to `.cm-img`, `.mt-eg`, `.ba-col .image-holder`, attribute images, generic Word-table images, and specification-table images. Never use `padding: 0` to make table media touch a card edge.
- Add `.row-head` to a non-empty first-column body cell when its compact text is shorter than 10 characters; this makes short labels such as `首图`、`第二张`、`白底图` bold. Do not infer row-header weight for longer first-column prose.
- Independent prose always uses `.text-block.plain-block > p`; `.text-block` alone is only a white wrapper and does not carry the canonical 28px prose contract.

## 2. Image layouts

- Use `.detail-screen-grid` for five or more consecutive screens; default to two columns.
- Use `.sample-image` for two or more examples under one label, making them equal half-width.
- Use `.half-image` for centered 图文详情 examples up to about half the content width.
- Use `.image-holder` around every source image except the synthetic Hero overlay.
- Use `.caption-image-card` only for a real image formula such as `[图A] + [图B] = 效果`; title formulas such as `[品牌] + 产品词` are not image captions.
- Use `.module-layout` plus `.ml-grid`/`.ml-block` only after visually inspecting the source diagram. Reproduce exact on-image wording and relative areas, and set an inline `aspect-ratio` on `.module-layout` matching the source image's overall ratio (e.g. `aspect-ratio: 1 / 1` for a square schematic).

Never preserve aspect ratio through `object-fit: contain` plus a fixed box for export-critical images. Prefer `width:100%;height:auto` or `width:auto;max-width:100%`.

## 3. Metrics and labels

- Use `.metric-emphasis` for every standalone metric line shaped like `[效果数据/样本/日期标签：] XX率 +X%/％/PP`: white background, green border/value, enlarged number+unit, and inline green up-arrow SVG. Keep the optional source label (for example `10SKU（0522-0531）：`) inside the same component; do not leave it as an ordinary paragraph.
- Keep a metric inside the relevant white module but below its label; let it fill available width.
- For a colon label, highlight only `.label-text`; render `.label-rest` on the next aligned line without a red square or highlight.
- For a colon-less label, use `.label-plain`: red square, no pink highlight.
- Keep lower-level captions grey and do not give them a red square.
- A DOCX hyperlink （操作手册 / 平台规则 / 白皮书 …) keeps its anchor text verbatim in the body copy and gains one red `.link-btn`「点击查看」pill right after it; the URL lives only on the button. Never expose raw URLs as text.

## 4. Video and export runtime

In the 主图视频 module, collapse all click-to-watch signals into one `点击播放` card. Pair `视频案例` and `点击播放` as equal-height table-style headers and keep the inline play icon. Do not expose source URLs as text links.

For page export:

- embed html2canvas locally;
- replace `<video>` elements with poster `<img>` elements before capture;
- flatten `.hero-overlay` and the solid red Hero into one temporary canvas image because html2canvas does not reproduce `mix-blend-mode` reliably;
- render arrows and metric icons as real inline SVG elements;
- restore temporary replacements after capture;
- keep both floating controls outside the captured image through `data-html2canvas-ignore`.
