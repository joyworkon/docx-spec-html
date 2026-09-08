#!/usr/bin/env python3
"""Extract a PyMuPDF-backed PDF structure manifest for model-led HTML reconstruction.

This is the PDF-only counterpart of ``extract_docx_manifest.py``: when no
companion DOCX exists, the PDF is the single source of truth, so the manifest
carries reading-order text lines (with font size/weight for hierarchy
decisions), image placements (dumped to files), border-detected tables, and
optional full-page renders for visual cross-checking.
"""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any

from pdf_source import clean_text, extract_page_tables, is_pdf, iter_text_lines, open_pdf


def body_font_size(lines: list[dict[str, Any]]) -> float:
    sizes = [round(line["size"], 1) for line in lines if line["size"] > 0]
    if not sizes:
        return 0.0
    return Counter(sizes).most_common(1)[0][0]


def heading_level(line: dict[str, Any], body_size: float) -> int | None:
    if not body_size or line["size"] <= body_size * 1.05:
        return None
    if line["size"] >= body_size * 1.6:
        return 1
    if line["size"] >= body_size * 1.25:
        return 2
    return 3 if line["bold"] else None


def extract_manifest(
    pdf_path: Path,
    images_dir: Path | None = None,
    render_pages: bool = False,
) -> dict[str, Any]:
    if not is_pdf(pdf_path):
        raise ValueError(f"not a PDF source: {pdf_path}")

    blocks: list[dict[str, Any]] = []
    headings: list[dict[str, Any]] = []
    image_count = 0
    table_count = 0
    xref_files: dict[int, str] = {}

    with open_pdf(pdf_path) as doc:
        import fitz  # noqa: PLC0415 - guaranteed importable after open_pdf

        all_lines: list[dict[str, Any]] = []
        per_page: list[tuple[int, list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]] = []
        for page_index, page in enumerate(doc):
            lines = list(iter_text_lines(page.get_text("dict")))
            infos = page.get_image_info(xrefs=True)
            tables = extract_page_tables(page, page_index)
            per_page.append((page_index, lines, infos, tables))
            all_lines.extend(lines)

        body_size = body_font_size(all_lines)

        if render_pages:
            renders_dir = (images_dir or pdf_path.with_suffix("")) / "page-renders"
            renders_dir.mkdir(parents=True, exist_ok=True)
            for page_index, page in enumerate(doc):
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                pixmap.save(str(renders_dir / f"page-{page_index + 1:02d}.png"))

        if images_dir:
            images_dir.mkdir(parents=True, exist_ok=True)

        for page_index, lines, infos, tables in per_page:
            page_blocks: list[dict[str, Any]] = []
            for line in lines:
                level = heading_level(line, body_size)
                block = {
                    "type": "text",
                    "page": page_index + 1,
                    "bbox": line["bbox"],
                    "text": line["text"],
                    "font_size": round(line["size"], 1),
                    "is_bold": line["bold"],
                    "fonts": line["fonts"],
                    "heading_level": level,
                    "is_heading_candidate": level is not None,
                }
                page_blocks.append(block)
                if level is not None:
                    headings.append(
                        {
                            "page": page_index + 1,
                            "text": line["text"],
                            "level": level,
                            "font_size": block["font_size"],
                        }
                    )
            for image_index, info in enumerate(infos, 1):
                image_count += 1
                stored_path = None
                if images_dir:
                    xref = int(info.get("xref", 0))
                    if xref and xref in xref_files:
                        stored_path = xref_files[xref]
                    else:
                        try:
                            extracted = doc.extract_image(xref) if xref else None
                        except Exception:  # noqa: BLE001 - keep manifest usable
                            extracted = None
                        if extracted:
                            filename = f"page-{page_index + 1:02d}-img-{image_index:02d}.{extracted['ext']}"
                            (images_dir / filename).write_bytes(extracted["image"])
                            stored_path = str(images_dir / filename)
                            if xref:
                                xref_files[xref] = stored_path
                page_blocks.append(
                    {
                        "type": "image",
                        "page": page_index + 1,
                        "bbox": [round(v, 2) for v in info.get("bbox", [])],
                        "width": info.get("width"),
                        "height": info.get("height"),
                        "xref": int(info.get("xref", 0)),
                        "path": stored_path,
                    }
                )
            for table in tables:
                table_count += 1
                page_blocks.append(
                    {
                        "type": "table",
                        "table_index": table_count,
                        **table,
                    }
                )
            page_blocks.sort(
                key=lambda block: ((block.get("bbox") or [0, 0])[1], (block.get("bbox") or [0, 0])[0])
            )
            blocks.extend(page_blocks)

    for block_index, block in enumerate(blocks):
        block["block_index"] = block_index

    texts = [block["text"] for block in blocks if block["type"] == "text"]
    return {
        "schema_version": 1,
        "source": str(pdf_path),
        "extractor": {"name": "PyMuPDF", "version": getattr(fitz, "VersionBind", "unknown")},
        "page_count": len(per_page),
        "block_count": len(blocks),
        "table_count": table_count,
        "image_count": image_count,
        "visible_text_count": len(texts),
        "body_font_size": body_size,
        "heading_count": len(headings),
        "headings": headings,
        "title_candidate": next((text for text in texts if text), ""),
        "images_dir": str(images_dir) if images_dir else None,
        "blocks": blocks,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract a PyMuPDF-backed PDF structure manifest for model-led HTML reconstruction."
    )
    parser.add_argument("pdf", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument(
        "--images-dir",
        type=Path,
        help="directory for extracted images (default: <pdf-stem>-images next to the PDF)",
    )
    parser.add_argument(
        "--render-pages",
        action="store_true",
        help="also render every page to <images-dir>/page-renders/page-NN.png for visual review",
    )
    args = parser.parse_args()

    images_dir = args.images_dir or args.pdf.with_name(f"{args.pdf.stem}-images")

    try:
        manifest = extract_manifest(args.pdf, images_dir, args.render_pages)
    except (FileNotFoundError, ValueError, RuntimeError) as exc:
        parser.error(str(exc))
    text = json.dumps(manifest, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
