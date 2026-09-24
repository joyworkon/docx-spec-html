#!/usr/bin/env python3
"""Make a final spec page small enough — and font-complete enough — to edit.

Why this step exists
--------------------
A raw model-led page embeds every source image at its original pixel size and
every font as a whole face. Real deliveries reached 25–55 MB, which is the one
thing third-party HTML editors cannot survive: they either refuse the file,
hang while parsing one 6 MB base64 attribute, or silently drop the tail of the
document. Nothing about the page is locked — it is plain ``<p>``/``<li>``/
``<td>`` text — so shrinking the payload is what actually restores editing.

Two things happen here, both idempotent:

1. Fonts. Only MiSans-Normal and MiSans-Bold ship (plus the hero title face).
   Each is subsetted to the characters the page uses **plus the 3755 GB2312
   level-1 common Chinese characters**, so a reviewer who retypes a sentence in
   another editor still gets MiSans glyphs instead of a system fallback. The
   faces are injected as woff2 data-URIs inside the single canonical ``<style>``
   block, ONE copy per weight, registered on the ``MiSans`` family at its
   numeric weight (400 / 700). The stylesheet asks for ``MiSans`` plus an
   explicit ``font-weight``, so Normal vs Bold resolve from the embedded faces
   without a second copy of the same glyphs.
2. Raster images. Every ``<img src="data:image/…;base64,…">`` is decoded,
   downscaled to at most ``--max-image-width`` device pixels (the poster is
   1280px wide, so 1600px already covers 2x review zoom) and re-encoded with
   the smallest lossless-or-visually-lossless codec. A candidate encoding is
   kept only when it is smaller than the original bytes.

3. Text hygiene. Any U+200B left over from the retired "inject a zero-width
   space between every CJK character to force justification" trick is removed.
   Blink justifies CJK without it, and an invisible character between every
   glyph breaks copy, search, spell-check and find-and-replace in every
   downstream editor.

Run it on the candidate, before ``review_gate.py`` / ``finalize_output.py``.
"""

from __future__ import annotations

import argparse
import base64
import binascii
import io
import json
import re
from html.parser import HTMLParser
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parents[1]
MISANS_SRC_DIR = SKILL_ROOT / "assets" / "fonts" / "misans-src"

# Only these two weights are part of the design system: Normal for body copy,
# Bold for every emphasis role (bracket titles, INTRODUCTION, red/grey square
# headings, metric box, table headers, buttons, outbound link text).
MISANS_WEIGHTS = {"Normal": 400, "Bold": 700}
TITLE_FAMILY = "JINGDONGLangZhengTi1-Bold"

PUNCTUATION = (
    "　「」『』【】〔〕《》〈〉（）［］｛｝·—…～‰℃×÷±≈≤≥≠√↑↓←→★☆❌✅"
    "①②③④⑤⑥⑦⑧⑨⑩⑪⑫、。，．；：？！“”‘’％＋－＝／｜＜＞＃＆＊"
)

IMG_DATA_URI_RE = re.compile(
    r'(<img\b[^>]*?\bsrc=")data:image/(?P<fmt>[a-zA-Z0-9.+-]+);base64,(?P<payload>[A-Za-z0-9+/=\s]+?)(")',
    re.IGNORECASE | re.DOTALL,
)
FONT_FACE_RE = re.compile(r"@font-face\s*\{[^}]*\}", re.IGNORECASE | re.DOTALL)


# ---------------------------------------------------------------------------
# character coverage
# ---------------------------------------------------------------------------
def gb2312_level1() -> str:
    """The 3755 most common simplified Chinese characters (GB2312 level 1)."""
    chars: list[str] = []
    for high in range(0xB0, 0xD8):
        for low in range(0xA1, 0xFF):
            try:
                chars.append(bytes((high, low)).decode("gb2312"))
            except UnicodeDecodeError:
                continue
    return "".join(chars)


def gb2312_full() -> str:
    chars: list[str] = []
    for high in range(0xB0, 0xF8):
        for low in range(0xA1, 0xFF):
            try:
                chars.append(bytes((high, low)).decode("gb2312"))
            except UnicodeDecodeError:
                continue
    return "".join(chars)


class _VisibleText(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in {"style", "script", "noscript"}:
            self.skip += 1

    def handle_endtag(self, tag):
        if tag in {"style", "script", "noscript"} and self.skip:
            self.skip -= 1

    def handle_data(self, data):
        if not self.skip:
            self.parts.append(data)


def page_visible_characters(html: str) -> str:
    parser = _VisibleText()
    parser.feed(html)
    return "".join(parser.parts)


def coverage_text(html: str, coverage: str) -> str:
    chars = set(page_visible_characters(html))
    chars.update(chr(code) for code in range(0x20, 0x7F))
    chars.update(PUNCTUATION)
    if coverage == "common":
        chars.update(gb2312_level1())
    elif coverage == "full":
        chars.update(gb2312_full())
    chars.discard("\u200b")
    return "".join(sorted(chars))


# ---------------------------------------------------------------------------
# fonts
# ---------------------------------------------------------------------------
def subset_woff2(font_bytes: bytes, text: str) -> bytes | None:
    try:
        from fontTools.subset import Options, Subsetter
        from fontTools.ttLib import TTFont
        import brotli  # noqa: F401  (woff2 flavour needs it)
    except ImportError:
        return None
    try:
        options = Options()
        options.flavor = "woff2"
        options.hinting = False
        options.desubroutinize = True
        options.layout_features = ["*"]
        options.name_IDs = ["*"]
        options.notdef_outline = True
        options.recalc_bounds = True
        options.recalc_average_width = True
        options.drop_tables += ["DSIG"]
        font = TTFont(io.BytesIO(font_bytes), fontNumber=0, lazy=True)
        subsetter = Subsetter(options)
        subsetter.populate(text=text)
        subsetter.subset(font)
        buffer = io.BytesIO()
        font.flavor = "woff2"
        font.save(buffer)
        return buffer.getvalue()
    except Exception:  # noqa: BLE001 - never break a delivery over a font
        return None


def embedded_font_payload(css: str, family: str) -> bytes | None:
    """Recover the font bytes a page already embedded for ``family``."""
    for rule in FONT_FACE_RE.findall(css):
        if f'"{family}"' not in rule and f"'{family}'" not in rule:
            continue
        match = re.search(r"base64,\s*([A-Za-z0-9+/=\s]+?)\s*\"\s*\)", rule)
        if not match:
            continue
        try:
            return base64.b64decode(re.sub(r"\s+", "", match.group(1)), validate=True)
        except (binascii.Error, ValueError):
            continue
    return None


def face_rule(family: str, weight: int, blob: bytes) -> str:
    uri = "data:font/woff2;base64," + base64.b64encode(blob).decode("ascii")
    return (
        "@font-face{"
        f'font-family:"{family}";font-weight:{weight};font-style:normal;'
        f'font-display:block;src:url("{uri}") format("woff2");'
        "}"
    )


def rebuild_fonts(html: str, coverage: str) -> tuple[str, dict]:
    style_open = html.find("<style")
    if style_open < 0:
        return html, {"embedded": [], "note": "no <style> block; fonts untouched"}
    style_body_start = html.find(">", style_open) + 1
    style_end = html.find("</style>", style_body_start)
    if style_body_start <= 0 or style_end < 0:
        return html, {"embedded": [], "note": "unparsable <style> block; fonts untouched"}

    css = html[style_body_start:style_end]
    text = coverage_text(html, coverage)

    title_source = embedded_font_payload(css, TITLE_FAMILY)
    if title_source is None:
        for candidate in sorted((SKILL_ROOT / "assets" / "fonts").glob(f"{TITLE_FAMILY}.*")):
            if candidate.suffix.lower() in {".ttf", ".otf", ".woff2", ".woff"}:
                title_source = candidate.read_bytes()
                break

    # Drop every existing @font-face for the families this script owns; they are
    # re-emitted below so repeated runs converge instead of stacking.
    owned = {TITLE_FAMILY, "MiSans", *(f"MiSans-{name}" for name in MISANS_WEIGHTS)}

    def drop_owned(match: re.Match[str]) -> str:
        rule = match.group(0)
        return "" if any(f'"{family}"' in rule for family in owned) else rule

    css = FONT_FACE_RE.sub(drop_owned, css)

    rules: list[str] = []
    embedded: list[dict] = []

    if title_source:
        # The hero face only ever renders the page title, so the page's own
        # character set is the right coverage; the whole face is 7.6 MB.
        blob = subset_woff2(title_source, coverage_text(html, "page"))
        if blob:
            rules.append(face_rule(TITLE_FAMILY, 700, blob))
            embedded.append({"family": TITLE_FAMILY, "weight": 700, "bytes": len(blob),
                             "source_bytes": len(title_source)})

    for name, weight in MISANS_WEIGHTS.items():
        source = MISANS_SRC_DIR / f"MiSans-{name}.ttf"
        if not source.exists():
            continue
        blob = subset_woff2(source.read_bytes(), text)
        if not blob:
            continue
        # One copy per weight. A second @font-face for the "MiSans-Normal" /
        # "MiSans-Bold" aliases would duplicate ~0.5 MB of identical glyphs,
        # which is exactly the payload bloat this script exists to remove.
        rules.append(face_rule("MiSans", weight, blob))
        embedded.append({"family": f"MiSans-{name}", "weight": weight, "bytes": len(blob),
                         "source_bytes": source.stat().st_size})

    css = "\n".join(rules) + "\n" + css.lstrip("\n")
    html = html[:style_body_start] + css + html[style_end:]
    return html, {
        "embedded": embedded,
        "coverage": coverage,
        "coverage_characters": len(text),
    }


# ---------------------------------------------------------------------------
# images
# ---------------------------------------------------------------------------
def recompress(raw: bytes, fmt: str, max_width: int, jpeg_quality: int) -> tuple[bytes, str] | None:
    try:
        from PIL import Image
    except ImportError:
        return None
    try:
        image = Image.open(io.BytesIO(raw))
        image.load()
    except Exception:  # noqa: BLE001
        return None
    if getattr(image, "n_frames", 1) > 1:  # animated source: leave it alone
        return None

    if image.width > max_width:
        height = max(1, round(image.height * max_width / image.width))
        image = image.resize((max_width, height), Image.LANCZOS)

    has_alpha = image.mode in {"RGBA", "LA", "PA"} or "transparency" in image.info
    candidates: list[tuple[bytes, str]] = []

    png_source = image.convert("RGBA") if has_alpha else image.convert("RGB")
    buffer = io.BytesIO()
    png_source.save(buffer, format="PNG", optimize=True)
    candidates.append((buffer.getvalue(), "png"))

    # Flat UI screenshots and schematics quantise losslessly-looking to 256
    # colours and get dramatically smaller.
    try:
        quantised = png_source.convert("RGB").quantize(colors=256, method=Image.MEDIANCUT, dither=Image.NONE)
        buffer = io.BytesIO()
        quantised.save(buffer, format="PNG", optimize=True)
        if not has_alpha:
            candidates.append((buffer.getvalue(), "png"))
    except Exception:  # noqa: BLE001
        pass

    if not has_alpha:
        buffer = io.BytesIO()
        image.convert("RGB").save(buffer, format="JPEG", quality=jpeg_quality, optimize=True,
                                  progressive=True, subsampling=0)
        candidates.append((buffer.getvalue(), "jpeg"))

    best = min(candidates, key=lambda item: len(item[0]))
    return best if len(best[0]) < len(raw) else None


def shrink_images(html: str, max_width: int, jpeg_quality: int, min_bytes: int) -> tuple[str, dict]:
    stats = {"images": 0, "rewritten": 0, "before_bytes": 0, "after_bytes": 0}

    def replace(match: re.Match[str]) -> str:
        stats["images"] += 1
        payload = re.sub(r"\s+", "", match.group("payload"))
        try:
            raw = base64.b64decode(payload, validate=True)
        except (binascii.Error, ValueError):
            return match.group(0)
        stats["before_bytes"] += len(raw)
        if len(raw) < min_bytes:
            stats["after_bytes"] += len(raw)
            return match.group(0)
        result = recompress(raw, match.group("fmt"), max_width, jpeg_quality)
        if result is None:
            stats["after_bytes"] += len(raw)
            return match.group(0)
        blob, fmt = result
        stats["after_bytes"] += len(blob)
        stats["rewritten"] += 1
        encoded = base64.b64encode(blob).decode("ascii")
        return f"{match.group(1)}data:image/{fmt};base64,{encoded}{match.group(4)}"

    return IMG_DATA_URI_RE.sub(replace, html), stats


# ---------------------------------------------------------------------------
def optimize(path: Path, out: Path, coverage: str, max_width: int, jpeg_quality: int,
             min_bytes: int, skip_images: bool, skip_fonts: bool) -> dict:
    html = path.read_text(encoding="utf-8")
    before = len(html.encode("utf-8"))
    report: dict = {"input": str(path), "output": str(out), "bytes_before": before}

    zwsp = html.count("\u200b")
    if zwsp:
        html = html.replace("\u200b", "")
    report["zero_width_spaces_removed"] = zwsp

    if not skip_fonts:
        html, font_report = rebuild_fonts(html, coverage)
        report["fonts"] = font_report
    if not skip_images:
        html, image_report = shrink_images(html, max_width, jpeg_quality, min_bytes)
        report["images"] = image_report

    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")
    after = len(html.encode("utf-8"))
    report["bytes_after"] = after
    report["mb_before"] = round(before / 1_000_000, 2)
    report["mb_after"] = round(after / 1_000_000, 2)
    report["saved_percent"] = round((before - after) * 100 / before, 1) if before else 0.0
    report["editor_friendly"] = after <= 24_000_000
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("page", type=Path, help="candidate single-file HTML")
    parser.add_argument("--out", type=Path, help="write here instead of in place")
    parser.add_argument("--font-coverage", choices=["page", "common", "full"], default="common",
                        help="page = only characters already on the page; common = + GB2312 "
                             "level 1 (3755 chars, keeps reviewer-typed text in MiSans); "
                             "full = + GB2312 level 2")
    parser.add_argument("--max-image-width", type=int, default=1600)
    parser.add_argument("--jpeg-quality", type=int, default=86)
    parser.add_argument("--min-image-bytes", type=int, default=40_000,
                        help="leave images smaller than this untouched")
    parser.add_argument("--skip-images", action="store_true")
    parser.add_argument("--skip-fonts", action="store_true")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()

    if not args.page.is_file():
        parser.error(f"not a file: {args.page}")
    out = args.out or args.page
    report = optimize(args.page, out, args.font_coverage, args.max_image_width,
                      args.jpeg_quality, args.min_image_bytes, args.skip_images, args.skip_fonts)
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
