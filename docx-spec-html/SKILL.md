---
name: docx-spec-html
description: Generate, batch-generate, validate, or refine production-quality single-file HTML specification pages from Word .docx documents, using the companion PDF as the authoritative hierarchy and table source. Use for JD-style 商品信息运营规范 documents, DOCX-to-HTML reconstruction, golden-quality matching, and editable or locked review deliverables.
---

# DOCX Spec HTML

Convert Word specification documents into polished 1280px single-file HTML pages while preserving text, images, tables, order, and hierarchy.

## Core contract

- Treat the companion PDF as the structural source of truth for content hierarchy, module boundaries, merged cells, and row/column relationships only. Never copy the PDF's colours, borders, corner treatment, spacing, or image-card styling; those always come from this Skill's HTML design system and golden reference.
- Never deliver raw `batch_generate.py` output as production work. It is extraction scaffolding only.
- Use exactly these eight first-level modules in order: `主图规范` / `主图视频` / `长标题` / `短标题` / `通用卖点` / `主推标签` / `品质标签` / `属性`.
- Strip the source `N、` prefix from module titles because the card already carries `01`–`08`. Never promote numbered body items or unnumbered phrases into extra chapters.
- Preserve every visible text occurrence, image occurrence, table relationship, and source block order. Hierarchy reconstruction may change grouping and styling, never reading order.

## Resource routing

- Read `references/structure.md` whenever reconstructing hierarchy or mapping DOCX/PDF content.
- Read `references/components.md` only when the document contains complex tables, comparisons, layout diagrams, long image runs, metrics, or video cards.
- Read `references/visual-qa.md` before final delivery and whenever browser/export rendering differs.
- Use `assets/styles.css` as the single canonical stylesheet. Do not copy CSS into Markdown references.
- Compare visual treatment with `assets/examples/auto-oil-golden-reference.html` and its three compressed WebP snapshots.
- Use `scripts/extract_docx_manifest.py` and `scripts/batch_generate.py` for deterministic extraction and draft generation. Use `scripts/dom_contracts.py` and `scripts/review_gate.py` for component/DOM validation. Publish production work only through `scripts/finalize_output.py`.

## Required workflow

1. Obtain and read the companion PDF.
2. Extract the DOCX manifest:

   ```bash
   python3 scripts/extract_docx_manifest.py source.docx --out source-manifest.json
   ```

3. Generate a baseline draft:

   ```bash
   python3 scripts/batch_generate.py source.docx output-dir
   ```

   Use `--style path.css` for custom CSS. Legacy `--design path.md` remains compatible. Add `--editable` only for an explicitly editable review copy.

4. Reconstruct hierarchy against the PDF. Resolve module boundaries, captions, merged cells, alternating headers, and image groupings with model judgment. Choose existing semantic components; never hand-author alternate wrappers or class combinations for them.
5. Review the HTML screen-by-screen against the PDF and golden reference. Fix every mismatch.
6. During iteration, validate the candidate page:

   ```bash
   python3 scripts/validate_output.py source.docx final-output.html --strict
   ```

7. During iteration, run the deterministic review gate. It auto-detects the body-care profile:

   ```bash
   python3 scripts/review_gate.py source.docx candidate.html --profile auto --out review-report.json
   ```

8. Complete the manual visual/runtime checks listed by the gate.
9. Publish through the only supported production exit. It re-runs source validation and DOM contracts, binds the final HTML SHA-256 to its report, and refuses to write a failed output:

   ```bash
   python3 scripts/finalize_output.py source.docx candidate.html final-output.html --profile auto --report review-report.json
   ```

Never deliver a candidate file or copy/rename it into place without this finalization step.

## Non-negotiable output rules

- Deliver one self-contained `.html` by default. Embed CSS, images, WOFF2 title font, editor, and html2canvas; use no CDN.
- Keep `标题 → 图片 → 说明` and every other source sequence unchanged. Titles above images remain above; captions below remain below.
- Preserve Word tables as row/column structures rather than unrelated cards.
- Render section titles exactly as `{ 标题 }`, with one inner space on both sides. Use `INTRODUCTION` on every card header.
- Force `商品信息运营规范` onto the second hero-title line. Use `JINGDONGLangZhengTi1-Bold` for `h1`, `MiSans-Bold` for card titles, and Chinese update text.
- Keep the Hero solid `#FF2B22`. Do not add gradients, textures, rings, paths, or radial highlights. A replaceable `.hero-overlay` is allowed only through the supplied runtime.
- Render both curved arrows as inline SVG elements, never CSS or data-URI backgrounds.
- Wrap grey-panel children in approved white modules; keep `.gray-panel > * + *` spacing. Do not leave bare labels, lists, or images directly under a grey panel.
- Use `.lead` for the overview lead; `.red-list` for bracketed or top-level `前缀：内容` items; `.label-line` and `.source-list` for nested labels/items; `.plain-block` for independent prose.
- Keep card body text near 28px and the enlarged 600px Hero proportions defined by the canonical stylesheet. Do not override Hero child positioning.
- Use `.detail-screen-grid` for five or more consecutive screen examples. Use `.sample-image` only for two or more same-label examples; use `.half-image` for 图文详情 examples.
- Use `.metric-emphasis` for conversion-rate metrics, `.ba-compare` for before/after comparisons, and `.spec-table` for three-column specifications.
- Treat canonical component DOM as immutable. Model judgment selects a component and maps source content into it; the generator owns its wrapper nesting and classes. Do not nest `.red-list` inside `.red-list`, omit `.plain-block` from independent prose, or detach a variant class such as `.tag-example-table` from its required base class.
- Render every semantic table with the one rounded-card system: 24px justified body text with the final line aligned left, 24px/700 centred red top headers, uniform light-grey `#f7f7f7` body cells, 10px cell corners, 8px gaps, and 12px equal inset around every table image. Use a non-red header only when the source explicitly gives it a distinct semantic role—not merely because the PDF uses another colour.
- In a table body, a non-empty first-column text cell shorter than 10 characters is a vertical row header by default and must be bold. Keep longer first-column prose at body weight unless its semantics explicitly require emphasis.
- Keep numbered siblings such as `1、` / `2、` / `3、` at identical weight and hierarchy. A colon inside one sibling, such as `3、时长：`, must not make only that sibling bold.
- Embed exactly one canonical `<style>` block and the `docx-spec-html/<release>` generator meta tag. Reject final HTML containing `body-care-review-fixes` or any other stacked review stylesheet; consolidate accepted changes into `assets/styles.css` and regenerate cleanly.

## Mandatory delivery checks

1. Include fixed `编辑` and `下载整页图片` buttons with `data-html2canvas-ignore`.
2. Decode the embedded Chinese editor through `Uint8Array` plus `TextDecoder('utf-8')`, never direct `atob()` text.
3. Confirm the Hero is clean solid red and the title/date/bracket positions are stable.
4. Confirm all arrows that enter downloaded PNGs are real inline SVG elements.
5. For any `首张主图模块化布局图`, inspect the source image and redraw its exact words, positions, proportions, and blocks with `.module-layout`; never invent a standard template.
6. Avoid `object-fit` for export-critical content images. Before html2canvas capture, flatten Hero blending and replace videos with poster images.
7. In the 主图视频 module, replace click-to-watch signals with one pink `点击播放` card; do not retain raw links.
8. Confirm no missing/underrepresented text, image-count mismatch, table loss, clipping, overlap, or white-card collision.
9. Under a bracket parent such as `【主图】`, render its explanatory children with grey squares, no pink highlight, and exactly the same nested indentation geometry as `.source-list` (`42px` group indent plus `25px` text offset at production scale); do not let the generic red-list rule override `.sublevel`.
10. Give every module-local `（1）`–`（4）` subtitle a pink marker. If a trailing parenthetical note contains a colon, keep that whole parenthesis together on the next line.
11. Align `视频案例` and `点击播放` as equal-height table-style headers; retain the inline play icon. Keep all table/card media proportional and use the canonical 12px equal padding on every side; never make an image flush with its card edge.
12. Confirm `review_gate.py` passes its DOM component contracts; then inspect screenshots and smoke-test both floating controls. Deliver only the file emitted by `finalize_output.py` with a matching SHA-256 review report.

## Editable review mode

Editable mode changes only the downloaded HTML, never the source DOCX. Re-run validation after reviewer edits whenever source fidelity still matters.

When feedback reveals a reusable rule, update the relevant reference and script. Keep document-specific exceptions in the generated output rather than the shared skill.
