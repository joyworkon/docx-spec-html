# Visual QA and delivery

Read this reference before every production delivery and after any CSS, editor, or export-runtime change.

## Contents

1. Visual reference
2. Browser inspection
3. Automated validation
4. Delivery

## 1. Visual reference

Open `assets/examples/auto-oil-golden-reference.html` for a lightweight component reference. Use these compressed snapshots when exact visual treatment matters:

- `assets/examples/golden-hero.webp`
- `assets/examples/golden-components.webp`
- `assets/examples/golden-table.webp`

Use the companion PDF—not the golden example—for document-specific content, hierarchy, image placement, and table semantics.

## 2. Browser inspection

Inspect at least the Hero, overview, one standard chapter, one complex table, and one image-heavy region.

Confirm:

- fixed 1280px poster width and no horizontal overflow;
- solid `#FF2B22` Hero, two-line title, bracket mark, Chinese date, and aligned white inline-SVG arrow;
- `{ 标题 }`, chapter number, `INTRODUCTION`, and red inline-SVG arrow on every card header;
- stable white-module spacing inside grey panels;
- red-square labels, title-only pink highlight, grey-square nested items, and consistent indentation;
- approximately 28px body text without clipping or overlap;
- proportional, centered images and unclipped captions;
- aligned table columns and merged headers/cells; body copy is justified with its final line aligned left while headers remain centered;
- one consistent rounded-card table system: 24px justified body copy, 24px/700 centered red top headers, light-grey body cells, bold short first-column row headers, 10px corners, 8px gaps, and 12px equal inset around every table image;
- grey-square bracket children aligned to the exact `.source-list` nested indent, identical typography across numbered siblings, and pink markers on all module-local numbered subtitles;
- equal-height `视频案例` / `点击播放` headers with the play icon retained;
- two-column long image runs rather than cramped four-column grids;
- no empty card, white-card collision, unexpected chapter, or source-order change.

Test both floating buttons. Confirm the downloaded PNG omits controls, includes arrows, uses video posters, and matches the blended Hero preview.

## 3. Automated validation

Run:

```bash
python3 scripts/validate_output.py source.docx final-output.html --strict
```

Then run:

```bash
python3 scripts/review_gate.py source.docx final-output.html [--profile body-care] --out review-report.json
```

The gate rejects missing release metadata, multiple/stacked stylesheets, known hierarchy and table regressions, external assets, missing controls, body-care profile mismatches, nested red lists, prose without `.plain-block`, malformed tag-example tables, and table images attached to the wrong media-cell component. The profile is auto-detected by default.

Resolve missing or underrepresented text, image-count mismatch, table loss, and CSS invariant failures before delivery.

Two intentional exceptions are already encoded in the validator:

- `.hero-overlay` is synthetic and does not require an `.image-holder`.
- `.caption-image-card` is required only for genuine image formulas containing an equals result, not title syntax containing brackets and plus signs.

Validation passing is necessary but not sufficient; browser inspection remains mandatory. Publish only through `finalize_output.py`, which binds the reviewed HTML hash to the report and does not emit failed candidates.

## 4. Delivery

Deliver one locked, self-contained HTML unless the user explicitly requests editable review mode or split assets. Report the output path and whether text, image count, tables, font embedding, and visual review passed.

Do not enable the in-page editable toolbar on a locked final page. The fixed `编辑` launcher and `下载整页图片` button remain required on every page.
