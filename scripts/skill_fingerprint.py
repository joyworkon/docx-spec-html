#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


IGNORED_NAMES = {".DS_Store"}
IGNORED_SUFFIXES = {".pyc"}


def fingerprint(root: Path) -> dict:
    digest = hashlib.sha256()
    files = 0
    size = 0
    for path in sorted(p for p in root.rglob("*") if p.is_file()):
        relative = path.relative_to(root)
        if any(part == "__pycache__" for part in relative.parts):
            continue
        if path.name in IGNORED_NAMES or path.suffix in IGNORED_SUFFIXES:
            continue
        data = path.read_bytes()
        digest.update(str(relative).encode("utf-8"))
        digest.update(b"\0")
        digest.update(data)
        digest.update(b"\0")
        files += 1
        size += len(data)
    return {"path": str(root.resolve()), "sha256": digest.hexdigest(), "files": files, "bytes": size}


def main() -> int:
    parser = argparse.ArgumentParser(description="Fingerprint one or two Skill directories.")
    parser.add_argument("paths", nargs="+", type=Path)
    args = parser.parse_args()
    if len(args.paths) not in (1, 2):
        parser.error("provide one path, or two paths to compare")
    results = [fingerprint(path) for path in args.paths]
    output = {"items": results}
    if len(results) == 2:
        output["equal"] = results[0]["sha256"] == results[1]["sha256"]
    print(json.dumps(output, ensure_ascii=False, indent=2))
    return 0 if output.get("equal", True) else 1


if __name__ == "__main__":
    raise SystemExit(main())
