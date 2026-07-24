#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import re
from pathlib import Path

from batch_generate import DEFAULT_EDITOR, SKILL_RELEASE
from dom_contracts import evaluate_dom_contracts
from validate_output import count_class, strip_tags, validate


BODY_CARE_MODULES = [
    "主图规范",
    "主图视频",
    "长标题",
    "短标题",
    "通用卖点",
    "主推标签",
    "品质标签",
    "属性",
]


def body_care_checks(html: str, report: dict) -> dict[str, bool]:
    h2_texts = [
        strip_tags(match).strip("{} ")
        for match in re.findall(r"<h2\b[^>]*>(.*?)</h2>", html, flags=re.I | re.S)
    ]
    local_pink = len(
        re.findall(
            r'class="[^"]*\blabel-text\b[^"]*"[^>]*>\s*[（(][1-4][）)]',
            html,
        )
    )
    numbered_item_bolded = bool(
        re.search(r"<li\b[^>]*>\s*<b>\s*[123][、.．]", html)
    )
    return {
        "fixed_eight_modules": h2_texts[1:] == BODY_CARE_MODULES,
        "image_contract_49_plus_1_redraw": report.get("expected_image_count") == 50
        and report.get("actual_image_count") == 49
        and report.get("redrawn_image_count") == 1,
        "table_contract_11": report.get("expected_table_count") == 11
        and report.get("actual_table_like_count") == 11,
        "three_metric_bars": count_class(html, "metric-emphasis") == 3,
        "three_overview_sublevels": count_class(html, "sublevel") == 3,
        "four_pink_local_subtitles": local_pink == 4,
        "nine_short_row_headers": count_class(html, "row-head") == 9,
        "six_red_attribute_headers": count_class(html, "attr-head-red") == 6
        and count_class(html, "attr-head-gray") == 0,
        "numbered_video_siblings_equal": not numbered_item_bolded,
        "video_pair_and_play_icon": count_class(html, "video-case-card") == 1
        and count_class(html, "video-demo") == 1
        and count_class(html, "vd-icon") == 1,
        "merged_showcase_header": bool(
            re.search(r'<th\b[^>]*colspan="3"[^>]*>展现样式</th>', html)
        ),
        "merged_inconsistency_header": bool(
            re.search(r'<th\b[^>]*colspan="3"[^>]*>信息不一致案例</th>', html)
        ),
        "upload_path_kept_together": bool(
            re.search(
                r'（4）丰富素材图类型</span></div><div class="label-rest">'
                r'（上传路径：京麦-商品-素材管理-商品素材）</div>',
                html,
            )
        ),
        "faithful_six_block_module_redraw": count_class(html, "module-layout") == 1
        and 'data-redraw-source="word/media/image1.png"' in html
        and all(
            text in html
            for text in ("品牌名", "场景", "商品主体", "卖点", "价促信息/商品名", "商品<br/>细节")
        ),
        "material_body_cells_use_light_gray": bool(
            re.search(
                r"\.poster\.auto-doc\s+\.material-table\s+td,[^{}]*\{"
                r"[^}]*background:\s*var\(--table-body-bg\)",
                html,
                flags=re.S,
            )
        ),
        "table_media_uses_equal_insets": all(
            bool(re.search(pattern, html, flags=re.S))
            for pattern in (
                r"\.poster\.auto-doc\s+\.compare-matrix\s+td\.cm-img\s*\{[^}]*padding:\s*var\(--table-media-padding\)",
                r"\.poster\.auto-doc\s+\.material-table\s+td\.mt-eg\s*\{[^}]*padding:\s*var\(--table-media-padding\)",
                r"\.poster\.auto-doc\s+\.ba-col\s+\.image-holder\s*\{[^}]*padding:\s*var\(--table-media-padding\)",
                r"\.poster\.auto-doc\s+\.doc-table\s+td\.table-media-cell\s*\{[^}]*padding:\s*var\(--table-media-padding\)",
                r"\.poster\.auto-doc\s+\.spec-cell\.spec-media-cell\s*\{[^}]*padding:\s*var\(--table-media-padding\)",
                r"\.poster\.auto-doc\s+\.attr-table\s+\.attr-img\s*\{[^}]*padding:\s*var\(--table-media-padding\)",
                r"\.poster\.auto-doc\s+\.video-case-media\s+\.image-holder\s*\{[^}]*padding:\s*var\(--table-media-padding\)",
            )
        ),
        "table_body_copy_is_centered": all(
            bool(re.search(pattern, html, flags=re.S))
            for pattern in (
                r"\.poster\.auto-doc\s+\.doc-table\s+td\s*\{[^}]*text-align:\s*center;[^}]*text-align-last:\s*center",
                r"\.poster\.auto-doc\s+\.compare-matrix\s+td,[^{}]*\{[^}]*text-align:\s*center;[^}]*text-align-last:\s*center",
                r"\.poster\.auto-doc\s+\.ba-text\s*\{[^}]*text-align:\s*center;[^}]*text-align-last:\s*center",
                r"\.poster\.auto-doc\s+\.spec-cell\s*\{[^}]*text-align:\s*center;[^}]*text-align-last:\s*center",
            )
        ),
    }


def review(docx: Path, html_path: Path, profile: str | None = "auto") -> dict:
    html = html_path.read_text(encoding="utf-8")
    source_report = validate(docx, html_path)
    requested_profile = profile or "auto"
    dom_profile = None if requested_profile == "generic" else requested_profile
    dom_checks, body_care_profile = evaluate_dom_contracts(html, dom_profile)
    effective_profile = "body-care" if body_care_profile else None
    editor_match = re.search(
        r'<script\b[^>]*\bid="editor-src-b64"[^>]*>([^<]+)</script>',
        html,
        flags=re.I | re.S,
    )
    try:
        embedded_editor_matches_vendor = bool(editor_match) and DEFAULT_EDITOR.exists() and (
            base64.b64decode(editor_match.group(1).strip()) == DEFAULT_EDITOR.read_bytes()
        )
    except (ValueError, TypeError):
        embedded_editor_matches_vendor = False
    generic_checks = {
        "strict_source_validation": bool(source_report.get("passed")),
        "release_marker_present": bool(
            re.search(
                rf'<meta\s+name="generator"\s+content="docx-spec-html/{re.escape(SKILL_RELEASE)}">',
                html,
            )
        ),
        "single_embedded_stylesheet": len(re.findall(r"<style\b", html, flags=re.I)) == 1,
        "no_review_patch_stack": "body-care-review-fixes" not in html,
        "canonical_design_tokens": all(
            token in html
            for token in (
                "--table-font-size: 24px",
                "--table-header-weight: 700",
                "--table-body-weight: 400",
                "--table-body-bg: #f7f7f7",
                "--table-cell-radius: 10px",
                "--table-cell-gap: 8px",
                "--table-media-padding: 12px",
                "--tag-example-media-height: 480px",
                "--nested-group-indent: 42px",
                "--nested-text-offset: 25px",
                "--video-header-height: 72px",
            )
        ),
        "offline_assets_only": not bool(
            re.search(r'<(?:img|script)\b[^>]*\bsrc=["\']https?://', html, flags=re.I)
        ),
        "fixed_review_controls": "data-html2canvas-ignore" in html
        and "id=\"edit-page-btn\"" in html
        and "id=\"dl-page-btn\"" in html,
        "embedded_editor_matches_vendor": embedded_editor_matches_vendor,
        "embedded_export_runtime": 'id="html2canvas-src-b64"' in html
        and "var module=undefined,exports=undefined,define=undefined;" in html
        and "new TextDecoder('utf-8').decode(bytes)" in html
        and "if(!b||!ensureH2C())return;" in html,
        "hero_title_export_color": bool(
            re.search(
                r"\.poster\.auto-doc\s+\.hero\s+h1\s*\{[^}]*color:\s*#fff;"
                r"[^}]*-webkit-text-fill-color:\s*#fff;",
                html,
                flags=re.S,
            )
        ),
        "inline_svg_contract": "<svg" in html and "class=\"metric-arrow\"" in html,
    }
    profile_checks = body_care_checks(html, source_report) if effective_profile == "body-care" else {}
    checks = {**generic_checks, **dom_checks, **profile_checks}
    warnings = [name for name, passed in checks.items() if not passed]
    return {
        "release": SKILL_RELEASE,
        "source": str(docx),
        "html": str(html_path),
        "requested_profile": requested_profile,
        "profile": effective_profile or "generic",
        "source_sha256": hashlib.sha256(docx.read_bytes()).hexdigest(),
        "html_sha256": hashlib.sha256(html_path.read_bytes()).hexdigest(),
        "checks": checks,
        "source_validation": {
            "passed": source_report.get("passed"),
            "missing_text_count": source_report.get("missing_text_count"),
            "underrepresented_text_count": source_report.get("underrepresented_text_count"),
            "expected_image_count": source_report.get("expected_image_count"),
            "actual_image_count": source_report.get("actual_image_count"),
            "redrawn_image_count": source_report.get("redrawn_image_count"),
            "synthetic_image_count": source_report.get("synthetic_image_count"),
            "expected_table_count": source_report.get("expected_table_count"),
            "actual_table_like_count": source_report.get("actual_table_like_count"),
            "expected_metric_count": source_report.get("expected_metric_count"),
            "actual_metric_count": source_report.get("actual_metric_count"),
        },
        "manual_visual_review_required": [
            "hero/overview/complex-table/image-heavy-region screenshots",
            "overlap, clipping, and horizontal overflow",
            "编辑 and 下载整页图片 runtime smoke tests",
        ],
        "passed": not warnings,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Run deterministic post-generation review gates.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("html", type=Path)
    parser.add_argument("--profile", choices=["auto", "body-care", "generic"], default="auto")
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    result = review(args.docx, args.html, args.profile)
    rendered = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(rendered, encoding="utf-8")
    print(rendered, end="")
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
