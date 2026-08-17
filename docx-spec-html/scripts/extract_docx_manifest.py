#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
from pathlib import Path
from typing import Any, Iterable


ZERO_WIDTH_RE = re.compile(r"[\u200b-\u200f\ufeff]")
STRUCTURAL_FORMAT_KEYS = {
    "align",
    "anchor",
    "behindText",
    "bold",
    "bold.cs",
    "colspan",
    "color",
    "firstLineIndent",
    "height",
    "hAlign",
    "hPosition",
    "hRelative",
    "indent",
    "italic",
    "italic.cs",
    "layout",
    "leftIndent",
    "listStyle",
    "name",
    "numFmt",
    "numId",
    "numLevel",
    "outlineLvl",
    "relId",
    "rightIndent",
    "rows",
    "size",
    "valign",
    "vAlign",
    "vmerge",
    "vPosition",
    "vRelative",
    "width",
    "wrap",
    "_gridCols",
    "colWidths",
    "cols",
}
STRUCTURAL_FORMAT_PREFIXES = ("shading.",)


class OfficeCLIError(RuntimeError):
    pass


def clean_text(value: str) -> str:
    value = ZERO_WIDTH_RE.sub("", value or "")
    return re.sub(r"\s+", " ", value).strip()


def compact_format(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        return {}
    return {
        key: item
        for key, item in value.items()
        if key in STRUCTURAL_FORMAT_KEYS or key.startswith(STRUCTURAL_FORMAT_PREFIXES)
    }


def node_children(node: dict[str, Any], node_type: str | None = None) -> list[dict[str, Any]]:
    children = [child for child in node.get("children", []) if isinstance(child, dict)]
    if node_type is None:
        return children
    return [child for child in children if child.get("type") == node_type]


def walk_nodes(node: dict[str, Any]) -> Iterable[dict[str, Any]]:
    yield node
    for child in node_children(node):
        yield from walk_nodes(child)


def parse_int(value: Any, default: int = 0) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def officecli_executable(requested: str) -> str:
    path = Path(requested).expanduser()
    if path.is_file():
        return str(path.resolve())
    resolved = shutil.which(requested)
    if resolved:
        return resolved
    raise OfficeCLIError(
        "OfficeCLI is required for DOCX extraction. Install it from "
        "https://github.com/iOfficeAI/OfficeCLI and verify `officecli --version`."
    )


def run_officecli(executable: str, *args: str, timeout: int = 120) -> dict[str, Any]:
    env = os.environ.copy()
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    completed = subprocess.run(
        [executable, *args],
        check=False,
        capture_output=True,
        text=True,
        timeout=timeout,
        env=env,
    )
    if completed.returncode != 0:
        detail = clean_text(completed.stderr) or clean_text(completed.stdout)
        raise OfficeCLIError(f"OfficeCLI failed ({completed.returncode}): {detail}")
    try:
        payload = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise OfficeCLIError(f"OfficeCLI returned invalid JSON: {exc}") from exc
    if isinstance(payload, dict) and payload.get("success") is False:
        raise OfficeCLIError(f"OfficeCLI reported failure: {payload}")
    return payload


def officecli_version(executable: str) -> str:
    env = os.environ.copy()
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    completed = subprocess.run(
        [executable, "--version"],
        check=False,
        capture_output=True,
        text=True,
        timeout=15,
        env=env,
    )
    if completed.returncode != 0:
        raise OfficeCLIError(clean_text(completed.stderr) or "Unable to read OfficeCLI version")
    return clean_text(completed.stdout)


def officecli_body(docx_path: Path, executable: str, depth: int) -> dict[str, Any]:
    payload = run_officecli(
        executable,
        "get",
        str(docx_path),
        "/body",
        "--depth",
        str(depth),
        "--json",
    )
    data = payload.get("data", {}) if isinstance(payload, dict) else {}
    results = data.get("results") or data.get("Results") or []
    if not results or not isinstance(results[0], dict) or results[0].get("type") != "body":
        raise OfficeCLIError("OfficeCLI did not return a /body document node")
    return results[0]


def image_manifest(node: dict[str, Any], index: int) -> dict[str, Any]:
    fmt = compact_format(node.get("format"))
    wrap = fmt.get("wrap")
    return {
        "index_in_paragraph": index,
        "path": node.get("path"),
        "relationship_id": fmt.get("relId"),
        "target": None,
        "name": fmt.get("name") or clean_text(str(node.get("text", ""))),
        "content_type": None,
        "byte_size": None,
        "width": fmt.get("width"),
        "height": fmt.get("height"),
        "wrap": wrap,
        "is_floating": bool(wrap and wrap != "inline"),
        "format": fmt,
    }


def paragraph_images(node: dict[str, Any]) -> list[dict[str, Any]]:
    pictures = [item for item in walk_nodes(node) if item.get("type") in {"picture", "image"}]
    return [image_manifest(item, index) for index, item in enumerate(pictures, 1)]


def run_manifest(node: dict[str, Any]) -> dict[str, Any]:
    return {
        "path": node.get("path"),
        "type": node.get("type"),
        "text": clean_text(str(node.get("text", ""))),
        "format": compact_format(node.get("format")),
    }


def paragraph_manifest(node: dict[str, Any], paragraph_index: int) -> dict[str, Any]:
    fmt = compact_format(node.get("format"))
    raw_outline = fmt.get("outlineLvl")
    outline_zero_based = parse_int(raw_outline, -1) if raw_outline is not None else None
    outline_level = outline_zero_based + 1 if outline_zero_based is not None else None
    style = str(node.get("style") or "")
    images = paragraph_images(node)
    run_nodes = [
        child
        for child in node_children(node)
        if child.get("type") in {"run", "picture", "image"}
        and (clean_text(str(child.get("text", ""))) or child.get("type") in {"picture", "image"})
    ]
    return {
        "paragraph_index": paragraph_index,
        "path": node.get("path"),
        "text": clean_text(str(node.get("text", ""))),
        "style": style,
        "outline_level": outline_level,
        "outline_level_zero_based": outline_zero_based,
        "list_level": parse_int(fmt.get("numLevel"), 0) if fmt.get("numId") is not None else None,
        "list_kind": fmt.get("listStyle") or fmt.get("numFmt"),
        "numbering_id": fmt.get("numId"),
        "is_heading_style": bool(
            outline_level is not None or "Heading" in style or "标题" in style
        ),
        "format": fmt,
        # A single run duplicates paragraph-level formatting. Preserve run detail only
        # when multiple runs carry localized emphasis or mixed content.
        "runs": [run_manifest(child) for child in run_nodes] if len(run_nodes) > 1 else [],
        "images": images,
        "image_count": len(images),
    }


def cell_manifest(node: dict[str, Any]) -> dict[str, Any]:
    paragraphs: list[dict[str, Any]] = []
    paragraph_index = 0
    for child in node_children(node):
        if child.get("type") != "paragraph":
            continue
        manifest = paragraph_manifest(child, paragraph_index)
        if manifest["text"] or manifest["images"]:
            paragraphs.append(manifest)
        paragraph_index += 1
    paragraph_text = clean_text(" ".join(paragraph["text"] for paragraph in paragraphs))
    return {
        "path": node.get("path"),
        "text": paragraph_text or clean_text(str(node.get("text", ""))),
        "format": compact_format(node.get("format")),
        "paragraphs": paragraphs,
    }


def cells_with_grid_columns(row: dict[str, Any]) -> list[tuple[int, int, dict[str, Any]]]:
    positioned: list[tuple[int, int, dict[str, Any]]] = []
    grid_column = 0
    for cell in node_children(row, "cell"):
        colspan = max(1, parse_int((cell.get("format") or {}).get("colspan"), 1))
        positioned.append((grid_column, colspan, cell))
        grid_column += colspan
    return positioned


def table_manifest(node: dict[str, Any], block_index: int, table_index: int) -> dict[str, Any]:
    row_nodes = node_children(node, "row")
    positioned_rows = [cells_with_grid_columns(row) for row in row_nodes]
    rows: list[dict[str, Any]] = []

    for row_index, positioned_cells in enumerate(positioned_rows):
        semantic_cells: list[dict[str, Any]] = []
        for grid_column, colspan, cell in positioned_cells:
            cell_format = cell.get("format") or {}
            vmerge = str(cell_format.get("vmerge") or "").lower()
            if vmerge == "continue":
                continue
            rowspan = 1
            if vmerge == "restart":
                for following_row in positioned_rows[row_index + 1 :]:
                    following = next(
                        (item for start, _span, item in following_row if start == grid_column),
                        None,
                    )
                    if not following or str((following.get("format") or {}).get("vmerge") or "").lower() != "continue":
                        break
                    rowspan += 1
            manifest = cell_manifest(cell)
            manifest.update(
                {
                    "grid_column": grid_column,
                    "colspan": colspan,
                    "rowspan": rowspan,
                }
            )
            semantic_cells.append(manifest)
        rows.append({"row_index": row_index, "cells": semantic_cells})

    table_format = compact_format(node.get("format"))
    column_count = parse_int(table_format.get("_gridCols"), 0)
    if not column_count:
        column_count = max(
            (start + span for row in positioned_rows for start, span, _cell in row),
            default=0,
        )
    return {
        "block_index": block_index,
        "type": "table",
        "table_index": table_index,
        "path": node.get("path"),
        "row_count": len(row_nodes),
        "column_count": column_count,
        "format": table_format,
        "rows": rows,
    }


def extract_manifest(docx_path: Path, officecli: str = "officecli", depth: int = 16) -> dict[str, Any]:
    if not docx_path.is_file():
        raise FileNotFoundError(docx_path)
    executable = officecli_executable(officecli)
    version = officecli_version(executable)
    body = officecli_body(docx_path, executable, depth)

    blocks: list[dict[str, Any]] = []
    paragraph_index = 0
    table_index = 0
    for block_index, child in enumerate(node_children(body)):
        node_type = child.get("type")
        if node_type == "paragraph":
            manifest = paragraph_manifest(child, paragraph_index)
            manifest.update({"block_index": block_index, "type": "paragraph"})
            if manifest["text"] or manifest["images"]:
                blocks.append(manifest)
            paragraph_index += 1
        elif node_type == "table":
            table_index += 1
            blocks.append(table_manifest(child, block_index, table_index))

    all_texts: list[str] = []
    for block in blocks:
        if block["type"] == "paragraph" and block["text"]:
            all_texts.append(block["text"])
        elif block["type"] == "table":
            for row in block["rows"]:
                for cell in row["cells"]:
                    if cell["text"]:
                        all_texts.append(cell["text"])

    paragraph_image_count = sum(
        len(block.get("images", []))
        for block in blocks
        if block["type"] == "paragraph"
    )
    table_image_count = sum(
        len(paragraph["images"])
        for block in blocks
        if block["type"] == "table"
        for row in block["rows"]
        for cell in row["cells"]
        for paragraph in cell["paragraphs"]
    )
    headings = [
        {
            "block_index": block["block_index"],
            "paragraph_index": block["paragraph_index"],
            "path": block.get("path"),
            "text": block["text"],
            "style": block["style"],
            "level": block["outline_level"],
        }
        for block in blocks
        if block["type"] == "paragraph" and block["is_heading_style"]
    ]

    return {
        "schema_version": 2,
        "source": str(docx_path),
        "extractor": {
            "name": "OfficeCLI",
            "version": version,
            "executable": executable,
            "depth": depth,
        },
        "body_child_count": parse_int(body.get("childCount"), len(node_children(body))),
        "block_count": len(blocks),
        "paragraph_count": paragraph_index,
        "table_count": table_index,
        "image_count": paragraph_image_count + table_image_count,
        "visible_text_count": len([text for text in all_texts if text]),
        "heading_count": len(headings),
        "headings": headings,
        "title_candidate": next(
            (block["text"] for block in blocks if block["type"] == "paragraph" and block["text"]),
            "",
        ),
        "blocks": blocks,
    }


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract an OfficeCLI-backed DOCX structure manifest for model-led HTML reconstruction."
    )
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path)
    parser.add_argument("--officecli", default="officecli", help="OfficeCLI executable or absolute path")
    parser.add_argument("--depth", type=int, default=16, help="OfficeCLI DOM expansion depth")
    args = parser.parse_args()

    try:
        manifest = extract_manifest(args.docx, args.officecli, args.depth)
    except (FileNotFoundError, OfficeCLIError, subprocess.TimeoutExpired) as exc:
        parser.error(str(exc))
    text = json.dumps(manifest, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
