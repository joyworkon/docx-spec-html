# High-Quality Workflow

Use this workflow when the user expects final production quality, asks for "best effect", or compares output against the golden page.

## Core Rule

Do not deliver raw `batch_generate.py` output as the final page. The batch script is a draft generator and extraction aid. Final quality requires model-led semantic reconstruction, project-specific layout choices, validation, and visual review.

## Required Stages

1. Extract source structure
   - Run `scripts/extract_docx_manifest.py` to create a paragraph/table/image manifest.
   - Inspect heading candidates, label-like paragraphs, image positions, tables, and repeated text.
   - Use Word structure and visual evidence from the source document, not only paragraph order.

2. Generate a baseline
   - Run `scripts/batch_generate.py` for a first-pass HTML and report.
   - Treat this output as scaffolding, not the final answer.

3. Model-led restructuring
   - Build or patch a project-specific HTML from the source manifest.
   - Decide hierarchy explicitly: hero title, overview card, numbered chapter cards, grey panel labels, lower-level grey-dot items, image captions, tables, examples.
   - Preserve all visible text and every image occurrence.
   - Preserve Word tables as row/column relationships. Do not turn table rows into unrelated cards.
   - Keep image captions and formula captions as lower-level content under the corresponding image.

4. Golden comparison
   - Compare the output against `assets/examples/auto-oil-golden-output.html`.
   - Match the quality bar, not necessarily the exact content:
     - hero treatment and title font;
     - upper-right `OPERATION STANDARDS` bracket mark;
     - Chinese update date, such as `更新日期 2026年06月`;
     - section title format `{ 标题 }`;
     - right label `INTRODUCTION` and the one-piece curved-hook SVG arrow (red hook-left under the label; white hook-right for the hero rule);
     - hero title forces `商品信息运营规范` onto the second line;
     - grey panels and white modules with consistent spacing;
     - `.label-line` red square plus the semi-transparent brand-red highlight bar `rgba(255,43,34,0.2)` only behind text;
     - lower-level items as grey-square `.source-list`;
     - table columns aligned by row and vertically centered;
     - five or more screen/detail examples in a two-column `.detail-screen-grid`, with full-width exceptions only for wide or critical images;
     - images preserved, proportional, and centered.

5. Validate
   - Run `scripts/validate_output.py source.docx output.html --strict`.
   - Resolve missing text, underrepresented repeated text, image count mismatch, and CSS invariant failures.
   - If validation passes but the browser view still looks wrong, fix the visual issue and revalidate.

6. Update reusable rules
   - If a browser comment reveals a general rule, update `references/mpdn50eu-design.md`.
   - If the issue is document-specific, fix only the generated HTML or project-specific generator.

## Common Failure Modes

- Raw script output looks worse than the golden page because no model hierarchy decisions were made.
- Section titles are rendered as `{{ 标题 }}` or without the required inner spaces instead of exactly `{ 标题 }`.
- Overview section uses `.section-head` without `.spec-head` or loses the right `INTRODUCTION` label.
- Generic heading detection promotes lower-level labels into huge section titles.
- Grey panels contain bare labels, bare lists, or direct images instead of white modules such as `.text-block`, `.image-frame`, `.caption-image-card`, or a proper grid.
- Image cards touch previous white modules because new modules did not inherit `.gray-panel > * + *`.
- Formula text below images is incorrectly promoted into a top card title.
- Consecutive detail-screen examples are squeezed into four columns or listed as one long vertical pile instead of using `.detail-screen-grid`.
- Word three-column tables are split into free cards and lose row correspondence.
- Content requirement columns vary row by row instead of sharing one column track.
- The `INTRODUCTION` arrow is stretched/distorted instead of keeping the SVG aspect ratio, or points the wrong way (red must hook left, white hero arrow must hook right).
- The overview card loses the right `INTRODUCTION` label.
- Hero right-top `OPERATION STANDARDS` mark or bottom update date is missing.
- The update date remains English `UPDATED YYYY.MM` instead of Chinese `更新日期 YYYY年MM月`.
- Editable mode is enabled for a locked final deliverable without user request.

## Recommended Deliverable Shape

For each source DOCX, deliver:

- `name-output.html`: final single-file HTML, not draft output.
- `name-report.json`: validation report.
- A brief note if any source item could not be represented exactly.

Do not deliver separate CSS unless the user asks.

Use `--editable` only when the user wants a review copy that can be edited directly in the browser. After text edits, the downloaded HTML may no longer match the source DOCX; re-run validation if source fidelity is still required.
