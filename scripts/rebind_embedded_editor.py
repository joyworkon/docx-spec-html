#!/usr/bin/env python3
"""Replace a final page's embedded editor without changing its document content."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import re
import tempfile
from pathlib import Path


EDITOR_SCRIPT = re.compile(
    r'(?P<open><script\s+type="application/octet-stream"\s+id="editor-src-b64"\s+'
    r'data-editor-sha256=")(?P<sha>[0-9a-f]{64})(?P<after_sha>"[^>]*>)'
    r'(?P<payload>.*?)(?P<close></script>)',
    re.DOTALL,
)
GENERATOR_META = re.compile(
    r'(<meta\s+name="generator"\s+content="docx-spec-html/)([^"]+)(">)',
)


def inspect(page: Path, editor: Path) -> dict:
    html = page.read_text(encoding="utf-8")
    matches = list(EDITOR_SCRIPT.finditer(html))
    if len(matches) != 1:
        raise ValueError(f"expected exactly one embedded editor, found {len(matches)}")
    match = matches[0]
    try:
        embedded = base64.b64decode(match.group("payload").strip(), validate=True)
    except ValueError as exc:
        raise ValueError("embedded editor payload is not valid base64") from exc
    actual_sha = hashlib.sha256(embedded).hexdigest()
    declared_sha = match.group("sha")
    if declared_sha != actual_sha:
        raise ValueError("embedded editor SHA-256 marker does not match its payload")
    expected = editor.read_bytes()
    expected_sha = hashlib.sha256(expected).hexdigest()
    return {
        "page": str(page),
        "editor": str(editor),
        "embedded_sha256": actual_sha,
        "expected_sha256": expected_sha,
        "embedded_bytes": len(embedded),
        "expected_bytes": len(expected),
        "matches": actual_sha == expected_sha,
        "generator_release": (
            GENERATOR_META.search(html).group(2) if GENERATOR_META.search(html) else None
        ),
    }


def rebind(page: Path, editor: Path, release: str) -> dict:
    before = inspect(page, editor)
    html = page.read_text(encoding="utf-8")
    editor_bytes = editor.read_bytes()
    editor_sha = hashlib.sha256(editor_bytes).hexdigest()
    payload = base64.b64encode(editor_bytes).decode("ascii")

    def replace_editor(match: re.Match[str]) -> str:
        return f'{match.group("open")}{editor_sha}{match.group("after_sha")}{payload}{match.group("close")}'

    html, editor_changes = EDITOR_SCRIPT.subn(replace_editor, html, count=1)
    html, release_changes = GENERATOR_META.subn(rf'\g<1>{release}\g<3>', html, count=1)
    if editor_changes != 1 or release_changes != 1:
        raise ValueError("page must contain one canonical editor payload and generator marker")

    with tempfile.NamedTemporaryFile(
        mode="w", encoding="utf-8", dir=page.parent, suffix=".html.tmp", delete=False
    ) as handle:
        handle.write(html)
        temporary = Path(handle.name)
    try:
        os.replace(temporary, page)
    finally:
        temporary.unlink(missing_ok=True)
    after = inspect(page, editor)
    after["previous_embedded_sha256"] = before["embedded_sha256"]
    after["generator_release"] = release
    after["updated"] = True
    return after


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("page", type=Path, help="final single-file HTML page")
    parser.add_argument(
        "--editor",
        type=Path,
        default=Path(__file__).resolve().parents[1] / "assets/vendor/html-editor.html",
        help="canonical editor HTML (default: this skill's vendor asset)",
    )
    parser.add_argument("--release", help="new docx-spec-html release for the generator meta tag")
    parser.add_argument("--write", action="store_true", help="perform the replacement")
    args = parser.parse_args()
    if not args.page.is_file() or not args.editor.is_file():
        parser.error("page and editor must both be existing files")
    if args.write and not args.release:
        parser.error("--release is required with --write")

    result = rebind(args.page, args.editor, args.release) if args.write else inspect(args.page, args.editor)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
