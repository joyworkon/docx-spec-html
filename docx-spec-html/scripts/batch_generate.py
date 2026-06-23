#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import html
import json
import mimetypes
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from validate_output import validate


SKILL_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DESIGN = SKILL_ROOT / "references" / "mpdn50eu-design.md"
DEFAULT_FONT = SKILL_ROOT / "assets" / "fonts" / "JINGDONGLangZhengTi1-Bold.ttf"
DEFAULT_H2C = SKILL_ROOT / "assets" / "vendor" / "html2canvas.min.js"


@dataclass
class ParagraphBlock:
    text: str
    images: list[str]
    style: str


@dataclass
class TableBlock:
    table: Table


def clean_text(value: str) -> str:
    value = re.sub(r"[\u200b-\u200f\ufeff]", "", value or "")
    return re.sub(r"\s+", " ", value).strip()


def esc(value: str) -> str:
    return html.escape(value, quote=True)


def slugify(value: str) -> str:
    value = clean_text(value)
    value = re.sub(r"[\\/:*?\"<>|]+", "-", value)
    value = re.sub(r"\s+", "-", value).strip("-")
    return value[:80] or "docx"


def iter_blocks(doc: DocumentObject) -> list[ParagraphBlock | TableBlock]:
    blocks: list[ParagraphBlock | TableBlock] = []
    for child in doc.element.body.iterchildren():
        if isinstance(child, CT_P):
            paragraph = Paragraph(child, doc)
            text = clean_text(paragraph.text)
            images = paragraph_images(doc, paragraph)
            if text or images:
                blocks.append(ParagraphBlock(text=text, images=images, style=paragraph.style.name if paragraph.style else ""))
        elif isinstance(child, CT_Tbl):
            blocks.append(TableBlock(Table(child, doc)))
    return blocks


def paragraph_images(doc: DocumentObject, paragraph: Paragraph) -> list[str]:
    targets: list[str] = []
    for blip in paragraph._p.xpath(".//a:blip"):
        rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
        if rid and rid in doc.part.rels:
            targets.append(doc.part.rels[rid].target_ref)
    return targets


def image_target_to_blob(doc: DocumentObject) -> dict[str, bytes]:
    blobs: dict[str, bytes] = {}
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            blobs[rel.target_ref] = rel.target_part.blob
    return blobs


def extract_css(design_path: Path, font_path: Path | None) -> str:
    design = design_path.read_text(encoding="utf-8")
    if "## 十七、完整 CSS 模板" in design and "## 十八、验收与交付协议" in design:
        section = design.split("## 十七、完整 CSS 模板", 1)[1].split("## 十八、验收与交付协议", 1)[0]
    else:
        section = design
    blocks = re.findall(r"```css\n(.*?)```", section, flags=re.S)
    css = "\n\n".join(blocks)
    if not css.strip():
        raise ValueError(f"No CSS code block found in {design_path}")
    if font_path and font_path.exists():
        font_data = base64.b64encode(font_path.read_bytes()).decode("ascii")
        css = re.sub(
            r'src:\s*url\("JINGDONGLangZhengTi1-Bold\.ttf"\)\s*format\("truetype"\);',
            f'src: url("data:font/ttf;base64,{font_data}") format("truetype");',
            css,
        )
    return css + GENERIC_CSS


GENERIC_CSS = """

/* ===== Generic DOCX generator additions ===== */
body { background: #737373; }
.poster.auto-doc .hero { height: 500px; }
.poster.auto-doc .hero h1 { max-width: 760px; }
.poster.auto-doc .hero + .card { margin-top: -75px; }
.poster.auto-doc .gray-panel > * + * { margin-top: 18px; }
.poster.auto-doc .plain-block p {
  margin: 0;
  font-size: 19px;
  line-height: 1.55;
  font-weight: 400;
}
.poster.auto-doc .plain-block p + p { margin-top: 10px; }
.poster.auto-doc .image-holder {
  display: grid;
  place-items: center;
  min-height: 180px;
}
.poster.auto-doc .image-frame .image-holder { margin-top: 10px; }
.poster.auto-doc .wide-image .doc-image {
  width: auto;
  max-width: 100%;
  max-height: 560px;
  object-fit: contain;
}
.poster.auto-doc .doc-table-wrap {
  background: #fff;
  border-radius: 10px;
  padding: 18px;
}
.poster.auto-doc .doc-table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 8px;
  table-layout: fixed;
}
.poster.auto-doc .doc-table th,
.poster.auto-doc .doc-table td {
  vertical-align: middle;
  border-radius: 8px;
  padding: 12px;
  font-size: 16px;
  line-height: 1.45;
}
.poster.auto-doc .doc-table th {
  background: #ff2b22;
  color: #fff;
  font-weight: 600;
  text-align: center;
}
.poster.auto-doc .doc-table td { background: #f7f7f7; }
.poster.auto-doc .doc-table .doc-image {
  width: auto;
  max-width: 100%;
  max-height: 300px;
  object-fit: contain;
}

/* Grey-square caption for example images (no red square, no highlight) */
.poster.auto-doc .caption-line {
  position: relative;
  margin: 0;
  padding-left: 25px;
  color: #555;
  font-size: 17px;
  line-height: 1.4;
  font-weight: 600;
}
.poster.auto-doc .caption-line::before {
  content: "";
  position: absolute;
  left: 0;
  top: 7px;
  width: 11px;
  height: 11px;
  border-radius: 3px;
  background: #d8d8d8;
}
.poster.auto-doc .text-block .caption-line { margin-top: 14px; }
.poster.auto-doc .text-block .image-holder { margin-top: 10px; }
/* Sub-level indent: example caption + images sit one level under their label,
   aligned with .source-list (28px). */
.poster.auto-doc .caption-line.indent { margin-left: 28px; }
.poster.auto-doc .image-holder.indent { margin-left: 28px; }

/* Half-page-width images (used in 图文详情) */
.poster.auto-doc .half-image .doc-image {
  width: 100%;
  max-width: 600px;
  height: auto;
  object-fit: contain;
}

/* Example images under a label (e.g. 短标题「示例：」): half the container
   width and all equal width, regardless of each picture's aspect ratio. */
.poster.auto-doc .sample-image .doc-image {
  width: 50%;
  max-width: 50%;
  height: auto;
  object-fit: contain;
}

/* Module-layout schematic: a clean redraw of the "首张主图模块化布局图"
   reference image (品牌 / 主要功能卖点 / 主品 / 赠品 / 物流质保 / 材质 / 营销卖点),
   keeping the original spatial arrangement but dropping the watermark. */
.poster.auto-doc .module-layout {
  display: flex;
  flex-direction: column;
  gap: 12px;
  max-width: 620px;        /* keep the original schematic's portrait proportion */
  margin: 0 auto;
  background: #fff8e1;     /* light-yellow fill */
  border: 2px solid #ff2b22;  /* red border */
  border-radius: 12px;
  padding: 22px;
}
.poster.auto-doc .module-layout > * { font-weight: 700; }
.poster.auto-doc .ml-brand {
  align-self: flex-start;
  background: #f4a09c;
  color: #7a1f1a;
  padding: 14px 30px;
  border-radius: 8px;
  font-size: 22px;
}
.poster.auto-doc .ml-frame {
  position: relative;
  background: #b6e6bd;
  border-radius: 12px;
  padding: 24px 22px 96px;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 18px;
  min-height: 420px;
}
.poster.auto-doc .ml-band {
  width: 100%;
  background: #f7c98b;
  color: #5a3d12;
  border-radius: 8px;
  padding: 20px;
  text-align: center;
  font-size: 26px;
}
.poster.auto-doc .ml-sub {
  background: #efe6a6;
  color: #5a4d12;
  border-radius: 22px;
  padding: 12px 30px;
  font-size: 22px;
}
.poster.auto-doc .ml-center {
  flex: 1;
  display: grid;
  place-items: center;
  text-align: center;
  color: #1f5a2a;
  font-size: 26px;
}
.poster.auto-doc .ml-corner {
  position: absolute;
  bottom: 22px;
  max-width: 32%;
  padding: 14px 18px;
  border-radius: 8px;
  font-size: 18px;
  font-weight: 600;
  line-height: 1.4;
}
.poster.auto-doc .ml-gift { left: 22px; background: #f6c0d0; color: #7a2447; }
.poster.auto-doc .ml-logi { right: 22px; background: #f4b78a; color: #6e3b13; text-align: right; }
.poster.auto-doc .ml-bottom {
  display: grid;
  grid-template-columns: 1fr 3fr;
  gap: 12px;
}
.poster.auto-doc .ml-bottom > div {
  border-radius: 8px;
  padding: 24px;
  text-align: center;
  font-size: 24px;
}
.poster.auto-doc .ml-material { background: #bcdcf2; color: #1c4a73; }
.poster.auto-doc .ml-market { background: #cfc9ee; color: #3a2f78; }

/* Conversion-metric emphasis bar: white fill, green outline, centred.
   Black label + green "+X%" with an up-arrow. */
.poster.auto-doc .metric-emphasis {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 12px;
  width: 100%;
  margin: 0;
  padding: 18px 24px;
  border: 2px solid #47b250;
  border-radius: 14px;
  background: #fff;
  font-size: 24px;
  font-weight: 700;
}
.poster.auto-doc .metric-emphasis .metric-text {
  color: #111;
  line-height: 1;
}
.poster.auto-doc .metric-emphasis .metric-value {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  color: #47b250;
  font-size: 40px;
  font-weight: 800;
  line-height: 1;
}
.poster.auto-doc .metric-emphasis .metric-arrow {
  height: 34px;
  width: auto;
  display: block;
}
/* Inside a white module, leave room between the bar and the table below it. */
.poster.auto-doc .text-block .metric-emphasis { margin: 0 0 12px; }

/* Before / after compare (优化前 grey, 优化后 red) */
.poster.auto-doc .ba-compare {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px;
}
.poster.auto-doc .ba-col {
  display: flex;
  flex-direction: column;
  gap: 10px;
}
.poster.auto-doc .ba-head {
  border-radius: 8px;
  padding: 10px;
  text-align: center;
  font-size: 18px;
  font-weight: 600;
}
.poster.auto-doc .ba-before { background: #f1f1f1; color: #555; }
.poster.auto-doc .ba-after { background: #ff2b22; color: #fff; }
.poster.auto-doc .ba-col .image-holder { min-height: 0; }
.poster.auto-doc .ba-text { margin: 0; font-size: 15px; line-height: 1.45; color: #333; }

/* Three-column spec table: 主图 narrowest, 内容要求 ~2x, 示例 half width */
.poster.auto-doc .spec-table {
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.poster.auto-doc .spec-row {
  display: grid;
  grid-template-columns: 1fr 2fr 3fr;
  gap: 8px;
  align-items: stretch;
}
.poster.auto-doc .spec-cell {
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 6px;
  background: #f7f7f7;
  border-radius: 8px;
  padding: 12px;
  font-size: 16px;
  line-height: 1.45;
  color: #333;
}
.poster.auto-doc .spec-head .spec-cell {
  background: #ff2b22;
  color: #fff;
  font-weight: 600;
  text-align: center;
  align-items: center;
}
.poster.auto-doc .spec-cell .image-holder {
  min-height: 0;
  width: 100%;
}
.poster.auto-doc .spec-cell .doc-image {
  /* Equal width via 100%, with height:auto driving the true aspect ratio.
     No max-height / object-fit: those force a wrong-ratio box that html2canvas
     (which ignores object-fit) would squish in the downloaded PNG. */
  display: block;
  width: 100%;
  height: auto;
}
.edit-toolbar {
  position: fixed;
  right: 18px;
  bottom: 18px;
  z-index: 9999;
  display: flex;
  gap: 8px;
  padding: 10px;
  border-radius: 10px;
  background: rgba(17, 17, 17, 0.88);
  color: #fff;
  font-family: "MiSans", "Microsoft YaHei", "PingFang SC", sans-serif;
}
.edit-toolbar button {
  border: 0;
  border-radius: 8px;
  padding: 8px 12px;
  background: #ff2b22;
  color: #fff;
  font-size: 13px;
  line-height: 1;
  cursor: pointer;
}
.edit-toolbar button.secondary { background: #555; }
.dl-page-btn {
  position: fixed;
  right: 18px;
  bottom: 18px;
  z-index: 9999;
  border: 0;
  border-radius: 10px;
  padding: 12px 18px;
  background: #ff2b22;
  color: #fff;
  font-family: "MiSans", "Microsoft YaHei", "PingFang SC", sans-serif;
  font-size: 14px;
  font-weight: 600;
  line-height: 1;
  cursor: pointer;
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.25);
}
.dl-page-btn[disabled] { opacity: 0.6; cursor: default; }
body.editing [contenteditable="true"] {
  outline: 2px dashed rgba(255, 43, 34, 0.75);
  outline-offset: 3px;
  cursor: text;
}

/* ===== Card body copy scaled ~1.5x (to match the target screenshot) =====
   Only the cards' body content grows; the hero and the title bar
   (.section-head: chapter number, { 标题 }, INTRODUCTION) keep their sizes.
   List dots, indents and module spacing scale with the text so proportions
   stay balanced. Higher specificity (.poster.auto-doc .class) overrides the
   §17 / base sizes above. */

/* 1) Body font sizes */
.poster.auto-doc .lead { font-size: 28px; }
.poster.auto-doc .plain-block p { font-size: 28px; }
.poster.auto-doc .red-list li,
.poster.auto-doc .red-list b,
.poster.auto-doc .red-list p { font-size: 28px; }
.poster.auto-doc .label-line { font-size: 28px; }
.poster.auto-doc .source-list li { font-size: 28px; }
.poster.auto-doc .source-list b { font-size: 28px; }
/* All text that follows a red-square title shares ONE size (28px): list items,
   sub-captions ("示例图：") and example lines. Table cells keep their own
   tighter tier below. */
.poster.auto-doc .example-line { font-size: 28px; }
.poster.auto-doc .caption-line { font-size: 28px; }
.poster.auto-doc .doc-table th,
.poster.auto-doc .doc-table td { font-size: 24px; }
.poster.auto-doc .spec-cell { font-size: 24px; }
.poster.auto-doc .ba-head { font-size: 27px; }
.poster.auto-doc .ba-text { font-size: 23px; }
/* Metric bar is NOT scaled with the body copy — it keeps the original
   24/40px text and 34px arrow defined above. */

/* Hero text scaled ~1.5x as well: main title, OPERATION STANDARDS, date.
   The hero grows taller so the enlarged title, rule and date keep their
   spacing and the white rule is not covered by the first card. */
.poster.auto-doc .hero { height: 600px; }
.poster.auto-doc .hero h1 { font-size: 102px; max-width: 1140px; }
.poster.auto-doc .hero-mark {
  font-size: 21px;
  width: 232px;
  height: 93px;
  border-radius: 48px;
  border-width: 2px;
}
.poster.auto-doc .updated { font-size: 27px; }

/* 2) List dots scaled (~11 -> 16) */
.poster.auto-doc .red-list li::before,
.poster.auto-doc .label-line::before,
.poster.auto-doc .source-list li::before,
.poster.auto-doc .red-list li > p:not(.sub-dot)::before,
.poster.auto-doc .sub-dot::before {
  width: 16px;
  height: 16px;
  top: 13px;
  border-radius: 4px;
}
.poster.auto-doc .caption-line::before {
  width: 16px;
  height: 16px;
  top: 11px;
  border-radius: 4px;
}

/* 3) Indents scaled */
.poster.auto-doc .red-list li,
.poster.auto-doc .label-line,
.poster.auto-doc .caption-line { padding-left: 38px; }
.poster.auto-doc .source-list { margin-left: 42px; }
.poster.auto-doc .sub-dot { padding-left: 42px; }
.poster.auto-doc .caption-line.indent,
.poster.auto-doc .image-holder.indent { margin-left: 42px; }

/* 4) Module spacing scaled */
.poster.auto-doc .gray-panel > * + * { margin-top: 27px; }
.poster.auto-doc .text-block { padding: 33px 39px; }
.poster.auto-doc .lead { margin-bottom: 33px; padding: 36px 51px; }
.poster.auto-doc .red-list li { margin-bottom: 24px; }
.poster.auto-doc .spec-text .red-list li { margin-bottom: 27px; }
.poster.auto-doc .label-line { margin-bottom: 15px; }
.poster.auto-doc .source-list li { margin-top: 18px; }
"""


EDITABLE_RUNTIME = """
<div class="edit-toolbar" data-html-edit-toolbar data-html2canvas-ignore>
  <button type="button" data-edit-toggle>编辑文字</button>
  <button type="button" data-edit-save>下载HTML</button>
  <button type="button" class="secondary" data-edit-exit>退出编辑</button>
</div>
<script>
(() => {
  const selectors = [
    "main.poster h1",
    "main.poster h2",
    "main.poster p",
    "main.poster li",
    "main.poster th",
    "main.poster td",
    "main.poster .label-text",
    "main.poster .chapter",
    "main.poster .en-label strong",
    "main.poster .hero-mark",
    "main.poster .image-caption-line span"
  ].join(",");
  const editableNodes = () => Array.from(document.querySelectorAll(selectors));
  const setEditing = (on) => {
    document.body.classList.toggle("editing", on);
    editableNodes().forEach((node) => {
      if (on) node.setAttribute("contenteditable", "true");
      else node.removeAttribute("contenteditable");
    });
    const toggle = document.querySelector("[data-edit-toggle]");
    if (toggle) toggle.textContent = on ? "正在编辑" : "编辑文字";
  };
  document.querySelector("[data-edit-toggle]")?.addEventListener("click", () => {
    setEditing(!document.body.classList.contains("editing"));
  });
  document.querySelector("[data-edit-exit]")?.addEventListener("click", () => setEditing(false));
  document.querySelector("[data-edit-save]")?.addEventListener("click", () => {
    const clone = document.documentElement.cloneNode(true);
    clone.querySelector("body")?.classList.remove("editing");
    clone.querySelectorAll("[contenteditable]").forEach((node) => node.removeAttribute("contenteditable"));
    const html = "<!doctype html>\\n" + clone.outerHTML;
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = (document.title || "edited-output") + "-edited.html";
    link.click();
    URL.revokeObjectURL(link.href);
  });
})();
</script>
"""


def label_line(text: str) -> str:
    return f'<div class="label-line"><span class="label-text">{esc(text)}</span></div>'


def source_list(items: list[str]) -> str:
    body = "".join(f"<li>{emphasize_prefix(esc(item))}</li>" for item in items if clean_text(item))
    return f'<ul class="source-list">{body}</ul>' if body else ""


def emphasize_prefix(text: str) -> str:
    for mark in ("：", ":"):
        if mark in text and text.index(mark) <= 18:
            idx = text.index(mark) + 1
            return f"<b>{text[:idx]}</b>{text[idx:].strip()}"
    return text


def plain_block(items: list[str]) -> str:
    paras = "".join(f"<p>{esc(item)}</p>" for item in items if clean_text(item))
    return f'<div class="text-block plain-block">{paras}</div>' if paras else ""


def text_block(label: str, items: list[str]) -> str:
    return f'<div class="text-block">{label_line(label)}{source_list(items)}</div>'


def image_tag(target: str, blobs: dict[str, bytes], alt: str) -> str:
    data = blobs[target]
    mime = mimetypes.guess_type(target)[0] or "application/octet-stream"
    uri = f"data:{mime};base64,{base64.b64encode(data).decode('ascii')}"
    return f'<img class="doc-image" src="{uri}" alt="{esc(alt)}">'


def image_frame(label: str, target: str, blobs: dict[str, bytes]) -> str:
    safe_label = label or "示例图："
    return (
        '<div class="image-frame wide-image">'
        f"{label_line(safe_label)}"
        f'<div class="image-holder">{image_tag(target, blobs, safe_label)}</div>'
        "</div>"
    )


def render_table(table: Table, doc: DocumentObject, blobs: dict[str, bytes]) -> str:
    rows_html: list[str] = []
    for row_idx, row in enumerate(table.rows):
        cells_html: list[str] = []
        tag = "th" if row_idx == 0 else "td"
        for cell in row.cells:
            texts = [clean_text(p.text) for p in cell.paragraphs if clean_text(p.text)]
            image_html = []
            for paragraph in cell.paragraphs:
                for target in paragraph_images(doc, paragraph):
                    if target in blobs:
                        image_html.append(f'<div class="image-holder">{image_tag(target, blobs, texts[0] if texts else "表格图片")}</div>')
            text_html = "<br>".join(esc(text) for text in texts)
            cells_html.append(f"<{tag}>{text_html}{''.join(image_html)}</{tag}>")
        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
    return '<div class="doc-table-wrap"><table class="doc-table">' + "".join(rows_html) + "</table></div>"


def document_uses_heading_styles(blocks: list[ParagraphBlock | TableBlock]) -> bool:
    """True when the DOCX marks its chapters with real Word heading styles."""
    for block in blocks:
        if isinstance(block, ParagraphBlock) and block.text:
            style = block.style.lower()
            if style.startswith("heading") or "标题" in style:
                return True
    return False


def is_section_heading(block: ParagraphBlock, heading_styles_present: bool = False) -> bool:
    text = block.text
    if not text or block.images:
        return False
    style = block.style.lower()
    if style.startswith("heading") or "标题" in style:
        return True
    if len(text) > 48:
        return False
    # A paragraph ending in a colon is a label/lead-in, never a chapter title.
    if text.endswith(("：", ":")):
        return False
    # Explicit chapter numerals always win, including bracketed forms like （一）.
    if re.match(r"^[（(]?[一二三四五六七八九十]+[）)、.．]", text):
        return True
    if re.match(r"^第[一二三四五六七八九十0-9]+[章节部分]", text):
        return True
    if re.match(r"^【第[一二三四五六七八九十0-9-]+屏】", text):
        return True
    # Heuristic fallbacks only fire when the doc has NO real heading styles to
    # trust. Otherwise they promote in-chapter labels (主图首张, 卖点建议顺序…)
    # into spurious top-level cards.
    if heading_styles_present:
        return False
    if re.match(r"^[0-9]{1,2}[.、．]\s*[^：:]{1,24}$", text):
        return True
    top_level_prefixes = ("整体规范", "主图", "主图视频", "长标题", "短标题", "卖点", "参数楼层", "属性", "图文详情")
    if text.startswith(top_level_prefixes) and len(text) <= 12 and "：" not in text and ":" not in text:
        return True
    return False


def clean_section_title(text: str) -> str:
    """Strip leading chapter numerals and trailing colons from a card title."""
    cleaned = text.strip()
    cleaned = re.sub(r"^[（(][一二三四五六七八九十0-9]+[）)]\s*", "", cleaned)
    cleaned = re.sub(r"^[一二三四五六七八九十]+[、.．]\s*", "", cleaned)
    cleaned = re.sub(r"^第[一二三四五六七八九十0-9]+[章节部分][:：、.．\s]*", "", cleaned)
    cleaned = cleaned.rstrip("：:").strip()
    return cleaned or text.strip()


def strip_title_prefix(title: str) -> str:
    """Remove a leading label like 主标题：/ 标题: from the document main title."""
    return re.sub(r"^\s*(?:主标题|标题|文档标题|page\s*title)\s*[:：]\s*", "", title, flags=re.I)


SCREEN_LABEL_RE = re.compile(r"^第\s*[0-9一二三四五六七八九十]+(?:\s*[-–~至]\s*[0-9一二三四五六七八九十]+)?\s*屏")


def is_label(text: str) -> bool:
    if not text or len(text) > 40:
        return False
    if text.endswith(("：", ":")):
        return True
    if re.match(r"^[0-9]{1,2}[.、．]", text):
        return True
    if SCREEN_LABEL_RE.match(text):
        return True
    return False


def split_sections(blocks: list[ParagraphBlock | TableBlock]) -> tuple[str, list[tuple[str, list[ParagraphBlock | TableBlock]]]]:
    first_text_index = next((i for i, block in enumerate(blocks) if isinstance(block, ParagraphBlock) and block.text), None)
    if first_text_index is None:
        return "未命名规范", [("整体规范综述", blocks)]
    title = strip_title_prefix(blocks[first_text_index].text)  # type: ignore[union-attr]
    body = blocks[first_text_index + 1 :]
    heading_styles_present = document_uses_heading_styles(blocks)
    sections: list[tuple[str, list[ParagraphBlock | TableBlock]]] = []
    current_title = "整体规范综述"
    current_blocks: list[ParagraphBlock | TableBlock] = []

    for block in body:
        if isinstance(block, ParagraphBlock) and is_section_heading(block, heading_styles_present):
            if current_blocks:
                sections.append((current_title, current_blocks))
            current_title = clean_section_title(block.text)
            current_blocks = []
        elif (
            isinstance(block, ParagraphBlock)
            and block.text
            and not block.images
            and block.text.rstrip("：:").strip() == current_title.rstrip("：:").strip()
        ):
            # Drop a paragraph that merely restates the current card title
            # (e.g. an "整体规范综述：" lead-in under the 整体规范综述 card).
            continue
        else:
            current_blocks.append(block)
    if current_blocks:
        sections.append((current_title, current_blocks))
    if not sections:
        sections.append(("整体规范综述", body))
    return title, sections


BRACKET_RE = re.compile(r"^\s*(【[^】]+】)\s*(.*)$")
CONVERSION_RE = re.compile(r"^[一-龥A-Za-z·]{2,12}\s*[+＋]\s*\d+(?:\.\d+)?\s*%$")
# "前缀：内容" lead-in, e.g. 总结：…/字数范围：…/卖点建议顺序：… — gets a red square + pink highlight.
COLON_PREFIX_RE = re.compile(r"^([^：:\n]{1,18}[：:])(.+)$")
# A conversion metric embedded inside a longer label, e.g. "2.优化前后图 商详转化率+2%".
METRIC_INLINE_RE = re.compile(r"[一-龥A-Za-z·]{2,12}\s*[+＋]\s*\d+(?:\.\d+)?\s*%")


def split_label_metric(label: str) -> tuple[str | None, str | None]:
    """Pull an embedded "XX率+X%" metric out of a label so it can be rendered as a
    standalone green emphasis bar, leaving the rest of the label as the title."""
    match = METRIC_INLINE_RE.search(label)
    if not match:
        return label, None
    metric = clean_text(match.group(0))
    rest = clean_text(label[: match.start()] + " " + label[match.end():]).strip(" 　：:")
    return (rest or None), metric


def is_conversion_metric(text: str) -> bool:
    """A standalone "XX率+X%" style metric line that deserves green emphasis."""
    return bool(CONVERSION_RE.fullmatch(clean_text(text)))


METRIC_ARROW_SVG = (
    '<svg class="metric-arrow" xmlns="http://www.w3.org/2000/svg" '
    'viewBox="0 0 32.34917 40.82425" fill="none" aria-hidden="true">'
    '<path d="M16.405537,0L1.1559057,14.706532L11.676984,15.342317Q11.316808,35.421619,0,40.824245'
    'Q19.693825,39.944675,21.828087,15.110984L32.349167,15.74677L16.405537,0Z" fill="#47B250"/></svg>'
)


def metric_emphasis(text: str) -> str:
    cleaned = clean_text(text)
    # Label stays black; the "+X%" (sign included) goes green with an up-arrow.
    match = re.match(r"^(.*?)\s*([+＋]\s*\d+(?:\.\d+)?\s*%)$", cleaned)
    if match and match.group(1).strip():
        head, value = match.group(1).strip(), match.group(2)
        inner = (
            f'<span class="metric-text">{esc(head)}</span>'
            f'<span class="metric-value">{esc(value)}{METRIC_ARROW_SVG}</span>'
        )
    else:
        inner = f'<span class="metric-text">{esc(cleaned)}</span>'
    return f'<div class="metric-emphasis">{inner}</div>'


def lead_block(text: str) -> str:
    return f'<p class="lead">{esc(text)}</p>'


def caption_line(text: str) -> str:
    """Grey-square caption (no red square, no highlight) for example images."""
    return f'<div class="caption-line">{esc(text)}</div>'


def red_list_block(items: list[str]) -> str:
    rows = []
    for item in items:
        item = clean_text(item)
        if not item:
            continue
        bracket = BRACKET_RE.match(item)
        if bracket:
            rows.append(f"<li><b>{esc(bracket.group(1))}</b>{esc(bracket.group(2))}</li>")
            continue
        colon = COLON_PREFIX_RE.match(item)
        if colon:
            rows.append(f"<li><b>{esc(colon.group(1))}</b>{esc(colon.group(2).strip())}</li>")
            continue
        rows.append(f"<li>{esc(item)}</li>")
    if not rows:
        return ""
    return f'<div class="text-block"><ul class="red-list">{"".join(rows)}</ul></div>'


def module_layout(items: list[str], fallback: str) -> str:
    """Faithfully redraw a「首张主图模块化布局图」schematic as clean coloured
    blocks, driven by the captured 主图首张 module names — same spatial layout
    (品牌 top-left, 主要功能卖点 banner, 主品 centre, 赠品/物流 corners, 材质/营销
    bottom row), no watermark. Falls back to the raw image when the expected
    modules can't be found."""
    names = [re.split(r"[：:]", it, 1)[0].strip() for it in items]

    def find(*keys: str) -> str | None:
        for name in names:
            if any(k in name for k in keys):
                return name
        return None

    brand = find("品牌")
    band = find("主要功能卖点", "核心卖点", "功能卖点")
    sub = find("其他功能", "多功能")
    center = find("场景", "产品展示", "主品")
    material = find("材质")
    market = find("营销卖点", "营销")
    if not (band and center):
        return fallback

    center_label = "主品（场景化展示）" if center and "场景" in center else (center or "主品")

    def blk(cls: str, text: str) -> str:
        return f'<div class="{cls}">{esc(text)}</div>'

    frame = blk("ml-band", band)
    if sub:
        frame += blk("ml-sub", sub)
    frame += blk("ml-center", center_label)
    frame += blk("ml-corner ml-gift", "赠品（可选）")
    frame += blk("ml-corner ml-logi", "物流信息 / 质保时间")
    return (
        '<div class="module-layout">'
        + blk("ml-brand", f"{brand or '品牌'} LOGO")
        + f'<div class="ml-frame">{frame}</div>'
        + '<div class="ml-bottom">'
        + blk("ml-material", material or "材质")
        + blk("ml-market", market or "主要营销卖点")
        + "</div></div>"
    )


def grouped_text_block(label: str | None, items: list[str], images: list[str], blobs: dict[str, bytes], half: bool, metric: str | None = None) -> str:
    """One white module that merges a label with its sub-items and example images.
    An embedded "XX率+X%" metric (pulled out of the label) renders as a green bar
    directly under the title, so it spans the module's inner width."""
    inner = label_line(label) if label else ""
    if metric:
        inner += metric_emphasis(metric)
    inner += source_list(items)
    real_images = [t for t in images if t in blobs]
    if real_images:
        # When the images sit under a label, indent the caption and images one
        # level (28px) so they align with .source-list sub-items.
        indent = " indent" if label else ""
        inner += f'<div class="caption-line{indent}">示例图：</div>'
        if half:
            holder_class = "image-holder half-image" + indent
        elif len(real_images) >= 2:
            # Multiple side-by-side reference images (e.g. 短标题「示例：」): half
            # the container width, all equal width.
            holder_class = "image-holder sample-image" + indent
        else:
            holder_class = "image-holder" + indent
        for target in real_images:
            inner += f'<div class="{holder_class}">{image_tag(target, blobs, "示例图：")}</div>'
    return f'<div class="text-block">{inner}</div>' if inner else ""


def classify_table(table: Table) -> str:
    if not table.rows:
        return "generic"
    header = " ".join(clean_text(cell.text) for cell in table.rows[0].cells)
    ncol = len(table.rows[0].cells)
    if "优化前" in header or "优化后" in header:
        return "before_after"
    if ncol >= 3 and ("内容要求" in header or "示例" in header):
        return "spec"
    if ncol == 2:
        return "before_after"
    return "generic"


def before_after(table: Table, doc: DocumentObject, blobs: dict[str, bytes]) -> str:
    rows = list(table.rows)
    if not rows:
        return ""
    header = [clean_text(cell.text) for cell in rows[0].cells]
    cols: list[str] = []
    for ci in range(len(rows[0].cells)):
        head = header[ci] if ci < len(header) else ""
        is_before = "前" in head or (ci == 0 and "后" not in head)
        head_class = "ba-before" if is_before else "ba-after"
        body = ""
        for row in rows[1:]:
            if ci >= len(row.cells):
                continue
            cell = row.cells[ci]
            for paragraph in cell.paragraphs:
                for target in paragraph_images(doc, paragraph):
                    if target in blobs:
                        body += f'<div class="image-holder">{image_tag(target, blobs, head)}</div>'
                txt = clean_text(paragraph.text)
                if txt:
                    body += f'<p class="ba-text">{esc(txt)}</p>'
        cols.append(f'<div class="ba-col"><div class="ba-head {head_class}">{esc(head)}</div>{body}</div>')
    return f'<div class="ba-compare">{"".join(cols)}</div>'


def spec_table(table: Table, doc: DocumentObject, blobs: dict[str, bytes]) -> str:
    rows_html: list[str] = []
    for row_idx, row in enumerate(table.rows):
        cells_html: list[str] = []
        for cell in row.cells:
            content = ""
            for paragraph in cell.paragraphs:
                txt = clean_text(paragraph.text)
                if txt:
                    content += f"<span>{esc(txt)}</span>"
                for target in paragraph_images(doc, paragraph):
                    if target in blobs:
                        content += f'<div class="image-holder">{image_tag(target, blobs, "示例")}</div>'
            cells_html.append(f'<div class="spec-cell">{content}</div>')
        row_class = "spec-row spec-head" if row_idx == 0 else "spec-row"
        rows_html.append(f'<div class="{row_class}">{"".join(cells_html)}</div>')
    return f'<div class="spec-table">{"".join(rows_html)}</div>'


def table_group(label: str | None, table: Table, doc: DocumentObject, blobs: dict[str, bytes], metric: str | None = None) -> str:
    kind = classify_table(table)
    if kind == "before_after":
        inner = before_after(table, doc, blobs)
    elif kind == "spec":
        inner = spec_table(table, doc, blobs)
    else:
        label_html = label_line(label) if label else ""
        return f'<div class="text-block">{label_html}</div>{render_table(table, doc, blobs)}' if label else render_table(table, doc, blobs)
    label_html = label_line(label) if label else ""
    # Embedded metric (e.g. 商详转化率+2%) sits under the title, spanning the same
    # width as the before/after (优化前+优化后) columns below it.
    metric_html = metric_emphasis(metric) if metric else ""
    return f'<div class="text-block">{label_html}{metric_html}{inner}</div>'


def render_section_blocks(blocks: list[ParagraphBlock | TableBlock], doc: DocumentObject, blobs: dict[str, bytes], *, is_intro: bool = False, half_images: bool = False) -> str:
    rendered: list[str] = []
    plain_items: list[str] = []
    bracket_items: list[str] = []
    pending_label: str | None = None
    pending_metric: str | None = None
    pending_items: list[str] = []
    pending_images: list[str] = []
    module_items: list[str] = []  # captured 主图首张 module names, for the layout redraw
    lead_done = False

    def flush_plain() -> None:
        nonlocal plain_items
        if plain_items:
            rendered.append(plain_block(plain_items))
            plain_items = []

    def flush_bracket() -> None:
        nonlocal bracket_items
        if bracket_items:
            rendered.append(red_list_block(bracket_items))
            bracket_items = []

    def flush_label() -> None:
        nonlocal pending_label, pending_metric, pending_items, pending_images, module_items
        if pending_label is None and not pending_images and not pending_metric:
            return
        # Capture the 主图首张 module list so a later 模块化布局图 can be redrawn.
        if pending_label and "主图首张" in pending_label:
            mods = [it for it in pending_items if ("：" in it or ":" in it)]
            if mods:
                module_items = mods
        # Redraw "首张主图模块化布局图" as a clean module schematic (no watermark)
        # instead of embedding the raw reference screenshot.
        if pending_label and "模块化布局图" in pending_label and pending_images and module_items:
            real = [t for t in pending_images if t in blobs]
            fallback = "".join(
                f'<div class="image-holder">{image_tag(t, blobs, "模块化布局图")}</div>' for t in real
            )
            ml = module_layout(module_items, fallback)
            rendered.append(f'<div class="text-block">{label_line(pending_label)}{ml}</div>')
        else:
            rendered.append(grouped_text_block(pending_label, pending_items, pending_images, blobs, half_images, pending_metric))
        pending_label = None
        pending_metric = None
        pending_items = []
        pending_images = []

    for block in blocks:
        if isinstance(block, TableBlock):
            flush_plain()
            flush_bracket()
            if pending_label is not None and not pending_items and not pending_images:
                label = pending_label
                metric = pending_metric
                pending_label = None
                pending_metric = None
                rendered.append(table_group(label, block.table, doc, blobs, metric=metric))
            else:
                flush_label()
                rendered.append(table_group(None, block.table, doc, blobs))
            continue

        if block.images and not block.text:
            flush_plain()
            flush_bracket()
            # Always accumulate into pending_images (even with no label) so that
            # consecutive image-only paragraphs land in ONE module and share a
            # single uniform width class, instead of each becoming its own
            # natural-width block.
            pending_images.extend(block.images)
            continue

        text = block.text
        if not text:
            continue

        if is_conversion_metric(text):
            flush_plain()
            flush_bracket()
            flush_label()
            rendered.append(metric_emphasis(text))
            continue

        if BRACKET_RE.match(text):
            flush_plain()
            flush_label()
            bracket_items.append(text)
            continue

        if is_label(text):
            flush_plain()
            flush_bracket()
            flush_label()
            rest, metric = split_label_metric(text)
            pending_label = rest
            pending_metric = metric
            pending_items = []
            pending_images = []
            continue

        if pending_label is not None:
            pending_items.append(text)
            continue

        # No active label: flush any loose accumulated images first so they keep
        # their original position relative to the text that follows.
        flush_label()

        if is_intro and not lead_done and not rendered and not plain_items and not bracket_items:
            rendered.append(lead_block(text))
            lead_done = True
        elif COLON_PREFIX_RE.match(text):
            # Top-level "前缀：内容" lead-in becomes a red-square + pink-highlight list item.
            flush_plain()
            bracket_items.append(text)
        else:
            flush_bracket()
            plain_items.append(text)

    flush_plain()
    flush_bracket()
    flush_label()
    return "".join(rendered) or plain_block([""])


def hero_title_html(title: str) -> str:
    """Force the "商品信息运营规范" category suffix onto the hero title's second line."""
    title = strip_title_prefix(title)
    marker = "商品信息运营规范"
    if marker in title:
        prefix = title[: title.index(marker)].rstrip()
        return f"{esc(prefix)}<br>{esc(marker)}" if prefix else esc(marker)
    return esc(title)


def section_head(num: int | None, title: str) -> str:
    chapter = f'<span class="chapter">{num:02d}</span>' if num is not None else ""
    return (
        '<div class="section-head spec-head">'
        f"{chapter}"
        f"<h2>{{ {esc(title)} }}</h2>"
        '<div class="en-label"><strong>INTRODUCTION</strong><span></span></div>'
        "</div>"
    )


def normalize_update_label(value: str | None, docx_path: Path | None = None) -> str:
    if not value:
        # Default to the source file's last-modified date, to the day.
        d = date.fromtimestamp(docx_path.stat().st_mtime) if docx_path is not None else date.today()
        return f"更新日期 {d.year}年{d.month}月{d.day}日"
    text = clean_text(value)
    if re.search(r"[\u4e00-\u9fff]", text):
        return text
    match = re.search(r"(20\d{2})[.\-/年 ]+(\d{1,2})(?:[.\-/月 ]+(\d{1,2}))?", text)
    if match:
        year, month, day = match.group(1), int(match.group(2)), match.group(3)
        if day:
            return f"更新日期 {year}年{month}月{int(day)}日"
        return f"更新日期 {year}年{month}月"
    return f"更新日期 {text}"


def render_card(title: str, body: str, num: int | None) -> str:
    card_class = "card intro-card" if num is None else "card spec-card"
    return f'<section class="{card_class}">{section_head(num, title)}<div class="gray-panel spec-text">{body}</div></section>'


def download_runtime() -> str:
    """A floating "下载整页图片" button that rasterises the whole poster to one
    PNG via an embedded html2canvas, so the page stays self-contained/offline.
    Scale is clamped so the canvas height stays under the browser limit."""
    if not DEFAULT_H2C.exists():
        return ""
    lib = DEFAULT_H2C.read_text(encoding="utf-8")
    return (
        '<button id="dl-page-btn" class="dl-page-btn" data-html2canvas-ignore>下载整页图片</button>\n'
        f"<script>{lib}</script>\n"
        "<script>(function(){var b=document.getElementById('dl-page-btn');"
        "if(!b||!window.html2canvas)return;"
        "b.addEventListener('click',function(){var p=document.querySelector('.poster');if(!p)return;"
        "var t=b.textContent;b.textContent='生成中…';b.disabled=true;"
        "var h=p.scrollHeight,s=Math.min(2,Math.max(0.3,32000/h));"
        "html2canvas(p,{scale:s,backgroundColor:'#dcedff',useCORS:true,logging:false,"
        "windowWidth:p.scrollWidth,windowHeight:h}).then(function(c){c.toBlob(function(bl){"
        "var a=document.createElement('a');a.href=URL.createObjectURL(bl);"
        "a.download=(document.title||'page')+'.png';document.body.appendChild(a);a.click();a.remove();"
        "setTimeout(function(){URL.revokeObjectURL(a.href);},1500);b.textContent=t;b.disabled=false;},'image/png');"
        "}).catch(function(e){console.error(e);b.textContent='下载失败，重试';b.disabled=false;});});})();</script>"
    )


def render_html(docx_path: Path, design_path: Path, font_path: Path | None, updated_label: str, editable: bool) -> str:
    doc = Document(docx_path)
    blobs = image_target_to_blob(doc)
    title, sections = split_sections(iter_blocks(doc))
    css = extract_css(design_path, font_path)
    cards: list[str] = []
    chapter = 1
    for section_title, section_blocks in sections:
        is_intro = section_title == "整体规范综述" and not cards
        half_images = "图文详情" in section_title
        body = render_section_blocks(
            section_blocks, doc, blobs, is_intro=is_intro, half_images=half_images
        )
        if is_intro:
            cards.append(render_card(section_title, body, None))
        else:
            cards.append(render_card(section_title, body, chapter))
            chapter += 1
    editable_runtime = EDITABLE_RUNTIME if editable else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{esc(title)}</title>
  <style>{css}</style>
</head>
<body>
<main class="poster auto-doc">
  <section class="hero">
    <div class="rings">
      <div class="ring ring-one"></div>
      <div class="ring ring-two"></div>
      <div class="ring ring-three"></div>
    </div>
    <h1>{hero_title_html(title)}</h1>
    <p class="updated">{esc(updated_label)}</p>
    <div class="hero-mark">OPERATION<br>STANDARDS</div>
    <div class="hero-rule"></div>
  </section>
  {''.join(cards)}
</main>
{download_runtime()}
{editable_runtime}
</body>
</html>
"""


def docx_inputs(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path]
    return sorted(path for path in input_path.rglob("*.docx") if not path.name.startswith("~$"))


def generate_one(docx_path: Path, output_dir: Path, design_path: Path, font_path: Path | None, updated_value: str | None, editable: bool) -> dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    out_html = output_dir / f"{slugify(docx_path.stem)}-output.html"
    report_path = output_dir / f"{slugify(docx_path.stem)}-report.json"
    updated_label = normalize_update_label(updated_value, docx_path)
    html_text = render_html(docx_path, design_path, font_path, updated_label, editable)
    out_html.write_text(html_text, encoding="utf-8")
    report = validate(docx_path, out_html)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return {
        "source": str(docx_path),
        "html": str(out_html),
        "report": str(report_path),
        "passed": report["passed"],
        "warnings": report["warnings"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Batch-generate MPDN50EU-style single-file HTML from DOCX files.")
    parser.add_argument("input", type=Path, help="A .docx file or a folder containing .docx files.")
    parser.add_argument("output_dir", type=Path, help="Folder for generated HTML and reports.")
    parser.add_argument("--design", type=Path, default=DEFAULT_DESIGN)
    parser.add_argument("--font", type=Path, default=DEFAULT_FONT)
    parser.add_argument("--updated", default=None, help='Hero update label. Examples: "2026.06" or "更新日期 2026年06月".')
    parser.add_argument("--editable", action="store_true", help="Add an optional in-page text editing toolbar and download button.")
    parser.add_argument("--strict", action="store_true", help="Exit non-zero if any generated report has warnings.")
    args = parser.parse_args()

    inputs = docx_inputs(args.input)
    if not inputs:
        raise SystemExit(f"No .docx files found: {args.input}")

    font_path = args.font if args.font and args.font.exists() else None
    summary = [generate_one(path, args.output_dir, args.design, font_path, args.updated, args.editable) for path in inputs]
    summary_path = args.output_dir / "batch-summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"count": len(summary), "summary": str(summary_path), "items": summary}, ensure_ascii=False, indent=2))
    return 1 if args.strict and any(not item["passed"] for item in summary) else 0


if __name__ == "__main__":
    raise SystemExit(main())
