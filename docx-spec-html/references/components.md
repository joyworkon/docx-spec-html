# Complex components

Read this file only when the source contains the corresponding component. Exact dimensions and visual styling live in `assets/styles.css`.

## Contents

1. Tables and comparisons
2. Image layouts
3. Metrics and labels
4. Video and export runtime

## 1. Tables and comparisons

### Multi-row optimization matrix

Use `.compare-matrix` for `优化内容 | 案例 | 优化前 | 优化后`. Merge each optimization group with `rowspan`; keep case rows and summary rows aligned. Put paired before/after images in the corresponding row. Image cells use `.cm-img` with zero padding and proportional full-width images.

### Material-type table

Use `.material-table` for `素材图类型 | 内容要求 | 示例`. Let the 示例 header span two image columns. Use four tracks at approximately `16% / 34% / 25% / 25%`; keep the two examples equal width and flush with cell edges.

### Attribute or keyword-direction table

Use `.attr-table` for text-only paired directions. Apply grey–red–grey alternating groups when shown by the PDF. In body-care documents, `适用肤质/香型 | 功效/成分 | 适用人群/净含量` belongs to 通用卖点, not 属性.

### Standard tables

- Use `.spec-table` for `主图 | 内容要求 | 示例`, with approximately `1fr 2fr 3fr` tracks and centered cells.
- Use `.word-table-spec` for other three-column semantic tables. Share one column track across all rows.
- Use `.ba-compare` for two-column 优化前/优化后 material; grey header before, red header after.
- Merge repeated adjacent headers with `colspan` when the PDF shows one shared heading.
- Use `.tag-example-table` when a tag example must appear larger than the draft default.

## 2. Image layouts

- Use `.detail-screen-grid` for five or more consecutive screens; default to two columns.
- Use `.sample-image` for two or more examples under one label, making them equal half-width.
- Use `.half-image` for centered 图文详情 examples up to about half the content width.
- Use `.image-holder` around every source image except the synthetic Hero overlay.
- Use `.caption-image-card` only for a real image formula such as `[图A] + [图B] = 效果`; title formulas such as `[品牌] + 产品词` are not image captions.
- Use `.module-layout` plus `.ml-grid`/`.ml-block` only after visually inspecting the source diagram. Reproduce exact on-image wording and relative areas.

Never preserve aspect ratio through `object-fit: contain` plus a fixed box for export-critical images. Prefer `width:100%;height:auto` or `width:auto;max-width:100%`.

## 3. Metrics and labels

- Use `.metric-emphasis` for `XX率 +X%`: white background, green border/value, enlarged percentage, and inline green up-arrow SVG.
- Keep a metric inside the relevant white module but below its label; let it fill available width.
- For a colon label, highlight only `.label-text`; render `.label-rest` on the next aligned line without a red square or highlight.
- For a colon-less label, use `.label-plain`: red square, no pink highlight.
- Keep lower-level captions grey and do not give them a red square.

## 4. Video and export runtime

In the 主图视频 module, collapse all click-to-watch signals into one `.video-play-card` with `点击播放`. Do not expose source URLs as text links.

For page export:

- embed html2canvas locally;
- replace `<video>` elements with poster `<img>` elements before capture;
- flatten `.hero-overlay` and the solid red Hero into one temporary canvas image because html2canvas does not reproduce `mix-blend-mode` reliably;
- render arrows and metric icons as real inline SVG elements;
- restore temporary replacements after capture;
- keep both floating controls outside the captured image through `data-html2canvas-ignore`.
