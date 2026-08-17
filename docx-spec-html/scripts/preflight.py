#!/usr/bin/env python3
from __future__ import annotations

import argparse
import importlib
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any


MINIMUM_PYTHON = (3, 9)
REQUIRED_FILES = (
    "SKILL.md",
    "requirements.txt",
    "agents/openai.yaml",
    "assets/styles.css",
    "assets/examples/auto-oil-golden-reference.html",
    "assets/examples/golden-components.webp",
    "assets/examples/golden-hero.webp",
    "assets/examples/golden-table.webp",
    "assets/fonts/JINGDONGLangZhengTi1-Bold.woff2",
    "assets/vendor/html-editor.html",
    "assets/vendor/html2canvas.min.js",
    "references/components.md",
    "references/structure.md",
    "references/visual-qa.md",
    "scripts/batch_generate.py",
    "scripts/dom_contracts.py",
    "scripts/extract_docx_manifest.py",
    "scripts/finalize_output.py",
    "scripts/rebind_embedded_editor.py",
    "scripts/review_gate.py",
    "scripts/skill_fingerprint.py",
    "scripts/validate_output.py",
)


def check_files(root: Path) -> dict[str, Any]:
    missing = [
        relative
        for relative in REQUIRED_FILES
        if not (root / relative).is_file() or (root / relative).stat().st_size == 0
    ]
    return {
        "ok": not missing,
        "detail": f"{len(REQUIRED_FILES) - len(missing)}/{len(REQUIRED_FILES)} required files",
        "missing": missing,
    }


def check_python() -> dict[str, Any]:
    version = tuple(sys.version_info[:3])
    return {
        "ok": version >= MINIMUM_PYTHON,
        "detail": f"Python {'.'.join(map(str, version))}",
        "executable": sys.executable,
        "required": ".".join(map(str, MINIMUM_PYTHON)) + "+",
    }


def check_python_docx() -> dict[str, Any]:
    try:
        module = importlib.import_module("docx")
    except ImportError:
        return {"ok": False, "detail": "python-docx is not installed"}
    return {
        "ok": True,
        "detail": f"python-docx {getattr(module, '__version__', 'unknown')}",
    }


def check_officecli() -> dict[str, Any]:
    requested = os.environ.get("OFFICECLI_BIN", "officecli")
    executable = requested if Path(requested).is_file() else shutil.which(requested)
    if not executable:
        return {"ok": False, "detail": "officecli is not available on PATH"}
    try:
        completed = subprocess.run(
            [str(executable), "--version"],
            check=False,
            capture_output=True,
            text=True,
            timeout=20,
            env={**os.environ, "OFFICECLI_SKIP_UPDATE": "1"},
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        return {"ok": False, "detail": f"officecli check failed: {exc}"}
    detail = (completed.stdout or completed.stderr).strip()
    return {
        "ok": completed.returncode == 0,
        "detail": detail or f"officecli exited with {completed.returncode}",
        "executable": str(executable),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a docx-spec-html Skill installation.")
    parser.add_argument(
        "--skill-root",
        type=Path,
        default=Path(__file__).resolve().parents[1],
        help="Directory containing SKILL.md; defaults to this script's parent Skill.",
    )
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON.")
    args = parser.parse_args()

    root = args.skill_root.expanduser().resolve()
    checks = {
        "files": check_files(root),
        "python": check_python(),
        "python_docx": check_python_docx(),
        "officecli": check_officecli(),
    }
    payload = {
        "ok": all(item["ok"] for item in checks.values()),
        "skill_root": str(root),
        "checks": checks,
    }

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        for name, item in checks.items():
            status = "PASS" if item["ok"] else "FAIL"
            print(f"[{status}] {name}: {item['detail']}")
            for missing in item.get("missing", []):
                print(f"  - missing: {missing}")
        print("READY" if payload["ok"] else "NOT READY")
    return 0 if payload["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
