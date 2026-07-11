#!/usr/bin/env python3
"""The only supported production exit for docx-spec-html candidates."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import tempfile
from pathlib import Path

from review_gate import review


def finalize(docx: Path, candidate: Path, output: Path, report_path: Path, profile: str) -> dict:
    result = review(docx, candidate, profile)
    if not result["passed"]:
        result["finalized"] = False
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        return result

    output.parent.mkdir(parents=True, exist_ok=True)
    final_result = result
    if candidate.resolve() != output.resolve():
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".html.tmp", delete=False) as handle:
            temporary = Path(handle.name)
        try:
            shutil.copyfile(candidate, temporary)
            staged_result = review(docx, temporary, profile)
            if not staged_result["passed"]:
                staged_result["finalized"] = False
                staged_result["candidate"] = str(candidate)
                report_path.parent.mkdir(parents=True, exist_ok=True)
                report_path.write_text(json.dumps(staged_result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
                return staged_result
            os.replace(temporary, output)
            final_result = staged_result
        finally:
            temporary.unlink(missing_ok=True)

    final_result["finalized"] = True
    final_result["candidate"] = str(candidate)
    final_result["html"] = str(output)

    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(final_result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return final_result


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate, DOM-contract-check, hash-bind, and publish one final HTML file."
    )
    parser.add_argument("docx", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--profile", choices=["auto", "body-care", "generic"], default="auto")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    report_path = args.report or args.output.with_name(args.output.stem + "-review-report.json")
    result = finalize(args.docx, args.candidate, args.output, report_path, args.profile)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("finalized") else 1


if __name__ == "__main__":
    raise SystemExit(main())
