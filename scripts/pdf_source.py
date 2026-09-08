#!/usr/bin/env python3
"""Shared PyMuPDF-backed PDF extraction helpers.

Used when no companion DOCX exists and the PDF is the single source of truth:
``validate_output`` counts source text/image/table expectations from these
helpers, and ``extract_pdf_manifest`` builds the reconstruction manifest from
the same primitives. PyMuPDF is imported lazily so DOCX-only workflows never
require it.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Iterable

ZERO_WIDTH_RE = re.compile(r"[\u200b-\u200f\ufeff]")


def clean_text(value: str) -> str:
    value = ZERO_WIDTH_RE.sub("", value or "")
    return re.sub(r"\s+", " ", value).strip()


def is_pdf(path: Path) -> bool:
    return path.suffix.lower() == ".pdf"


def open_pdf(pdf_path: Path) -> Any:
    if not is_pdf(pdf_path):
        raise ValueError(f"not a PDF source: {pdf_path}")
    if not pdf_path.is_file():
        raise FileNotFoundError(pdf_path)
    try:
        import fitz  # PyMuPDF
    except ImportError as exc:
        raise RuntimeError(
            "PDF sources require PyMuPDF: pip install pymupdf"
        ) from exc
    return fitz.open(pdf_path)


def span_is_bold(span: dict[str, Any]) -> bool:
    # PyMuPDF span flags: bit 4 (16) marks synthetic bold; embedded bold fonts
    # only show up in the font name.
    return bool(span.get("flags", 0) & 16) or "bold" in str(span.get("font", "")).lower()


def iter_text_lines(page_dict: dict[str, Any]) -> Iterable[dict[str, Any]]:
    """Yield one record per text line, in the order PyMuPDF reports blocks."""
    for block in page_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            text = clean_text("".join(str(span.get("text", "")) for span in spans))
            if not text:
                continue
            yield {
                "text": text,
                "bbox": [round(v, 2) for v in line.get("bbox", block.get("bbox", []))],
                "size": max((float(span.get("size", 0)) for span in spans), default=0.0),
                "bold": any(span_is_bold(span) for span in spans),
                "fonts": sorted({str(span.get("font", "")) for span in spans if span.get("font")}),
            }


def pdf_texts(pdf_path: Path) -> list[str]:
    """Visible text as line-level fragments, page by page.

    Lines (not paragraphs) are the right granularity for validation: every PDF
    line must survive as a substring of the rendered HTML, while PDF line
    wrapping never has to match DOCX paragraph boundaries.
    """
    texts: list[str] = []
    with open_pdf(pdf_path) as doc:
        for page in doc:
            for line in iter_text_lines(page.get_text("dict")):
                texts.append(line["text"])
    return texts


def pdf_image_count(pdf_path: Path) -> int:
    """Count image placements (occurrences), mirroring DOCX ``a:blip`` counting."""
    count = 0
    with open_pdf(pdf_path) as doc:
        for page in doc:
            count += len(page.get_image_info())
    return count


def extract_page_tables(page: Any, page_index: int) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    for table in page.find_tables().tables:
        rows = [
            [clean_text(cell or "") for cell in row]
            for row in table.extract()
        ]
        tables.append(
            {
                "page": page_index + 1,
                "bbox": [round(v, 2) for v in table.bbox],
                "row_count": table.row_count,
                "column_count": table.col_count,
                "rows": rows,
            }
        )
    return tables


def pdf_tables(pdf_path: Path) -> list[dict[str, Any]]:
    """Border-detected tables via ``page.find_tables()`` (line strategies).

    Borderless layout grids are intentionally not guessed here; treat the
    count as a lower bound and resolve borderless structures with model
    judgment against the page render.
    """
    tables: list[dict[str, Any]] = []
    with open_pdf(pdf_path) as doc:
        for page_index, page in enumerate(doc):
            tables.extend(extract_page_tables(page, page_index))
    return tables


def pdf_table_count(pdf_path: Path) -> int:
    return len(pdf_tables(pdf_path))
