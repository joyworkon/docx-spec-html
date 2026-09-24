#!/usr/bin/env python3
"""DOM-level component contracts for final DOCX specification HTML.

The generator owns component markup. Model judgment may choose a component and
map source content into it, but it must not invent alternate wrappers/classes.
This module intentionally uses only the Python standard library so every Agent
running the Skill can enforce the same contracts.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from html import unescape
from html.parser import HTMLParser


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

VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link",
    "meta", "param", "source", "track", "wbr",
}
KEPT_ATTRS = {"class", "id", "style", "colspan", "rowspan", "data-component"}


@dataclass(eq=False)
class Node:
    # eq=False keeps identity semantics: a recursive field-by-field comparison
    # on a whole document tree blows the recursion limit (and is never what a
    # contract means by "this element").
    tag: str
    attrs: dict[str, str] = field(default_factory=dict)
    parent: "Node | None" = None
    children: list["Node"] = field(default_factory=list)
    content: list[object] = field(default_factory=list)

    @property
    def classes(self) -> set[str]:
        return set(self.attrs.get("class", "").split())

    def has_class(self, name: str) -> bool:
        return name in self.classes

    def descendants(self, tag: str | None = None) -> list["Node"]:
        found: list[Node] = []
        stack = list(reversed(self.children))
        while stack:
            node = stack.pop()
            if tag is None or node.tag == tag:
                found.append(node)
            stack.extend(reversed(node.children))
        return found

    def elements(self, tag: str | None = None) -> list["Node"]:
        nodes = [self] if tag is None or self.tag == tag else []
        return nodes + self.descendants(tag)

    def text(self) -> str:
        parts = [item.text() if isinstance(item, Node) else str(item) for item in self.content]
        return re.sub(r"\s+", " ", unescape(" ".join(parts))).strip()

    def ancestor(self, *, tag: str | None = None, class_name: str | None = None) -> "Node | None":
        node = self.parent
        while node is not None:
            if (tag is None or node.tag == tag) and (class_name is None or node.has_class(class_name)):
                return node
            node = node.parent
        return None


class ContractHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.root = Node("document")
        self.stack = [self.root]

    def _node(self, tag: str, attrs: list[tuple[str, str | None]]) -> Node:
        kept = {key: value or "" for key, value in attrs if key in KEPT_ATTRS}
        node = Node(tag.lower(), kept, self.stack[-1])
        self.stack[-1].children.append(node)
        self.stack[-1].content.append(node)
        return node

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        node = self._node(tag, attrs)
        if tag.lower() not in VOID_TAGS:
            self.stack.append(node)

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self._node(tag, attrs)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        for index in range(len(self.stack) - 1, 0, -1):
            if self.stack[index].tag == tag:
                del self.stack[index:]
                return

    def handle_data(self, data: str) -> None:
        if self.stack[-1].tag not in {"script", "style"} and data.strip():
            self.stack[-1].content.append(data)


def parse_html(html: str) -> Node:
    parser = ContractHTMLParser()
    parser.feed(html)
    parser.close()
    return parser.root


def section_for_title(root: Node, title: str) -> Node | None:
    for section in root.descendants("section"):
        for heading in section.descendants("h2"):
            if heading.text().strip("{} ") == title:
                return section
    return None


def detected_body_care(root: Node) -> bool:
    titles = [node.text().strip("{} ") for node in root.descendants("h2")]
    return titles[1:] == BODY_CARE_MODULES


def _direct_elements(node: Node, tag: str | None = None) -> list[Node]:
    return [child for child in node.children if tag is None or child.tag == tag]


def _first_row(table: Node) -> Node | None:
    rows = table.descendants("tr")
    return rows[0] if rows else None


def _closest_cell(image: Node, table: Node) -> Node | None:
    node = image.parent
    while node is not None and node is not table:
        if node.tag in {"td", "th"}:
            return node
        node = node.parent
    return None


def _tag_example_contract(table: Node) -> bool:
    if not {"doc-table", "tag-example-table"}.issubset(table.classes):
        return False
    if table.parent is None or not table.parent.has_class("doc-table-wrap"):
        return False
    first_row = _first_row(table)
    if first_row is None:
        return False
    header_cells = [node for node in _direct_elements(first_row) if node.tag in {"th", "td"}]
    if not header_cells or any(node.tag != "th" or not node.text() for node in header_cells):
        return False
    for image in table.descendants("img"):
        cell = _closest_cell(image, table)
        if cell is None or not cell.has_class("table-media-cell"):
            return False
        if image.parent is None or not image.parent.has_class("image-holder"):
            return False
    return True


# ---------------------------------------------------------------------------
# Marker-level ladder: red square (L1) > grey square (L2) > hollow square (L3)
# > round dot (L4). A level is NEVER skipped and NEVER shares its parent's
# left edge: text that hangs under a red square is a grey square, text under a
# grey square is a hollow square, text under a hollow square is a round dot.
# ---------------------------------------------------------------------------
GREY_LEVEL_CLASSES = {"sublevel", "caption-line"}
NUMBERED_ITEM_RE = re.compile(r"^\s*(?:\(|（)?\d{1,2}(?:\)|）|[、.．,:：])")
LETTERED_ITEM_RE = re.compile(r"^\s*[a-hA-H](?:\)|）|[、.．])")
LABEL_ONLY_RE = re.compile(r"[：:]\s*$")


def _marker_level(node: Node) -> int | None:
    """Which rung of the marker ladder this node renders, or None."""
    if node.has_class("dot"):
        return 4
    if node.has_class("deep"):
        return 3
    if node.classes & GREY_LEVEL_CLASSES:
        return 2
    list_parent = node.ancestor(tag="ul") or node.ancestor(tag="ol")
    if list_parent is not None and list_parent.has_class("dot-list"):
        return 4
    if list_parent is not None and list_parent.has_class("source-list"):
        return 2
    if list_parent is not None and list_parent.has_class("red-list"):
        return 1
    if node.has_class("label-line"):
        return 1
    return None


def _marker_items(root: Node) -> list[Node]:
    items = [node for node in root.descendants("li")]
    items += [node for node in root.descendants() if node.has_class("label-line")
              or node.has_class("caption-line")]
    return items


def _nearest_marker_ancestor(node: Node) -> Node | None:
    current = node.parent
    while current is not None:
        if current.tag == "li" or current.has_class("label-line") or current.has_class("caption-line"):
            return current
        current = current.parent
    return None


def marker_ladder_checks(root: Node) -> dict[str, bool]:
    items = _marker_items(root)

    # A nested marker item is exactly one rung below its nearest marker parent.
    no_skip = True
    no_flat_nesting = True
    for node in items:
        level = _marker_level(node)
        parent_item = _nearest_marker_ancestor(node)
        if level is None or parent_item is None:
            continue
        parent_level = _marker_level(parent_item)
        if parent_level is None:
            continue
        if level == parent_level:
            no_flat_nesting = False
        elif level != parent_level + 1:
            no_skip = False

    # Hollow squares only ever appear under a grey square; round dots only ever
    # appear under a hollow square. Anything else is a skipped rung.
    hollow_under_grey = True
    for node in (item for item in items if item.has_class("deep")):
        parent_item = _nearest_marker_ancestor(node)
        source_list = node.ancestor(tag="ul", class_name="source-list")
        if parent_item is None and source_list is None:
            hollow_under_grey = False
    dot_under_hollow = all(
        (_nearest_marker_ancestor(node) is not None
         and _marker_level(_nearest_marker_ancestor(node)) == 3)
        or node.ancestor(tag="ul", class_name="dot-list") is not None
        for node in items if node.has_class("dot")
    )

    # A grey-square group must hang under a red-square heading: nested inside a
    # .red-list item, or preceded in its own card by a red-square element.
    # Table cells are their own parent context and are exempt.
    def grey_group_has_red_parent(group: Node) -> bool:
        if group.ancestor(tag="li") is not None:
            return True
        container = group.parent
        if container is not None and (
            container.tag in {"td", "th"}
            or container.has_class("image-frame")
            or any("cell" in name for name in container.classes)
        ):
            return True
        scope = group.ancestor(tag="section") or root
        order = scope.descendants()
        try:
            index = order.index(group)
        except ValueError:
            return True
        return any(
            node.classes & {"label-line", "red-list", "example-line"}
            for node in order[:index]
        )

    grey_lists = [node for node in root.descendants("ul") if node.has_class("source-list")]

    # Numbered items (1./2./3.) that hang under an open label are children, not
    # siblings of that label: they must carry the next marker rung, not sit flush
    # against the red square.
    numbered_demoted = True
    lettered_demoted = True
    for red_list in (node for node in root.descendants("ul") if node.has_class("red-list")):
        siblings = _direct_elements(red_list, "li")
        for index, node in enumerate(siblings):
            if not LABEL_ONLY_RE.search(node.text()):
                continue
            if node.classes & GREY_LEVEL_CLASSES or node.has_class("deep"):
                continue
            following = siblings[index + 1:]
            numbered = [item for item in following if NUMBERED_ITEM_RE.match(item.text())]
            if len(numbered) >= 2 and any(_marker_level(item) == 1 for item in numbered):
                numbered_demoted = False

    # a./b./c. items belong one rung below the 1./2./3. item they explain —
    # whether the source nested them or left them as flat siblings.
    for node in items:
        if not LETTERED_ITEM_RE.match(node.text()):
            continue
        level = _marker_level(node)
        parent_item = _nearest_marker_ancestor(node)
        if parent_item is None or level is None:
            continue
        if NUMBERED_ITEM_RE.match(parent_item.text()):
            parent_level = _marker_level(parent_item)
            if parent_level is not None and level <= parent_level:
                lettered_demoted = False

    for list_node in root.descendants("ul") + root.descendants("ol"):
        numbered_level: int | None = None
        for item in _direct_elements(list_node, "li"):
            text = item.text()
            level = _marker_level(item)
            if NUMBERED_ITEM_RE.match(text):
                numbered_level = level
                continue
            if LETTERED_ITEM_RE.match(text) and numbered_level is not None:
                if level is not None and level <= numbered_level:
                    lettered_demoted = False

    return {
        "marker_ladder_has_no_skipped_level": no_skip,
        "nested_marker_item_steps_in_one_level": no_flat_nesting,
        "hollow_square_sits_under_grey_square": hollow_under_grey,
        "round_dot_sits_under_hollow_square": dot_under_hollow,
        "grey_square_group_hangs_under_red_heading": all(
            grey_group_has_red_parent(group) for group in grey_lists
        ),
        "numbered_children_are_demoted_one_level": numbered_demoted,
        "lettered_children_are_demoted_below_numbers": lettered_demoted,
    }


# ---------------------------------------------------------------------------
# Source-box integrity, before/after body cells, play-card geometry
# ---------------------------------------------------------------------------
def _cells_of(table: Node) -> list[Node]:
    return [node for node in table.descendants() if node.tag in {"td", "th"}]


def _is_single_cell_table(node: Node) -> bool:
    if node.tag != "table":
        return False
    return len(node.descendants("tr")) == 1 and len(_cells_of(node)) == 1


def source_box_checks(root: Node) -> dict[str, bool]:
    """One bordered/shaded source box renders as exactly one white card.

    The recurring defect is a page that turns every single source LINE into its
    own one-row/one-cell ``.doc-table`` inside its own ``.doc-table-wrap``, so a
    box that reads as one paragraph in the PDF arrives as five stacked white
    cards. A one-cell table is never a table: it is a label or a paragraph.
    """
    single_cell_tables = [node for node in root.descendants("table") if _is_single_cell_table(node)]

    consecutive_wraps = True
    for container in root.descendants():
        run = 0
        for child in container.children:
            wrapped = child.has_class("doc-table-wrap") or child.tag == "table"
            table = child if child.tag == "table" else next(
                (node for node in child.children if node.tag == "table"), None
            )
            if wrapped and table is not None and _is_single_cell_table(table):
                run += 1
                if run >= 2:
                    consecutive_wraps = False
                    break
            else:
                run = 0

    return {
        "no_single_cell_doc_table": not single_cell_tables,
        "source_box_not_fragmented_into_stacked_cards": consecutive_wraps,
    }


def compare_and_play_checks(root: Node) -> dict[str, bool]:
    ba_compares = [node for node in root.descendants() if node.has_class("ba-compare")]

    # Every text row under a 优化前/优化后 header pair is a table BODY cell.
    allowed_col_children = {"ba-head", "ba-text", "image-holder", "ba-media-row",
                            "doc-table-wrap", "video-demo", "metric-emphasis"}
    ba_text_is_body_cell = True
    for compare in ba_compares:
        columns = [node for node in compare.children if node.has_class("ba-col")]
        targets = columns or [compare]
        for column in targets:
            for child in column.children:
                if child.has_class("ba-head"):
                    continue
                if child.classes & allowed_col_children:
                    continue
                # A bare <p>/<div>/<span> of copy dropped straight under the
                # headers has no body-cell background: that is the defect.
                if child.tag in {"p", "div", "span", "ul", "ol"} and child.text():
                    ba_text_is_body_cell = False

    # A 点击播放 card always inherits a cell box so it matches the neighbouring
    # table's width and height instead of floating at its own size.
    play_cards = [node for node in root.descendants() if node.has_class("video-demo")]
    play_hosts = {"table-media-cell", "video-cell", "video-case-media", "ba-col",
                  "image-holder", "spec-media-cell"}
    play_matches_cell = all(
        (node.parent is not None and node.parent.classes & play_hosts)
        or node.classes & {"match-cell", "match-head"}
        for node in play_cards
    )

    return {
        "ba_compare_text_uses_table_body_cells": ba_text_is_body_cell,
        "play_card_matches_neighbour_cell_box": play_matches_cell,
    }


def generic_dom_checks(root: Node) -> dict[str, bool]:
    red_lists = [node for node in root.descendants("ul") if node.has_class("red-list")]
    text_only_blocks: list[Node] = []
    for block in (node for node in root.descendants("div") if node.has_class("text-block")):
        children = _direct_elements(block)
        if children and all(child.tag == "p" for child in children):
            text_only_blocks.append(block)

    tag_tables = [node for node in root.descendants("table") if node.has_class("tag-example-table")]
    table_images: list[tuple[Node, Node]] = []
    for table in root.descendants("table"):
        table_images.extend((image, table) for image in table.descendants("img"))
    def media_cell_matches_table(image: Node, table: Node) -> bool:
        cell = _closest_cell(image, table)
        if cell is None:
            return False
        if table.has_class("doc-table"):
            return cell.has_class("table-media-cell")
        if table.has_class("compare-matrix"):
            return cell.has_class("cm-img")
        if table.has_class("material-table"):
            return cell.has_class("mt-eg")
        if table.has_class("attr-table"):
            return cell.has_class("attr-img")
        return False

    # Indentation, marker size and alignment come from the ladder in
    # assets/styles.css, never from an inline style on a level class.
    protected = {"sublevel", "plain-block", "tag-example-table", "table-media-cell",
                 "doc-table-wrap", "deep", "dot", "source-list", "caption-line",
                 "ba-text", "video-demo"}
    protected_inline_override = any(
        node.classes.intersection(protected)
        and re.search(r"(?:font-size|padding|margin-left|text-align)\s*:", node.attrs.get("style", ""), re.I)
        for node in root.descendants()
    )

    # .ba-compare must sit inside a .text-block white container: the .ba-before
    # header's light-grey fill is invisible against the grey panel otherwise.
    ba_compares = [node for node in root.descendants() if node.has_class("ba-compare")]

    checks = {
        "no_nested_red_list": not any(node.ancestor(tag="ul", class_name="red-list") for node in red_lists),
        "plain_prose_uses_plain_block": all(block.has_class("plain-block") for block in text_only_blocks),
        "tag_example_table_component_contract": not tag_tables
        or all(_tag_example_contract(table) for table in tag_tables),
        "table_images_use_semantic_media_cells": all(media_cell_matches_table(image, table) for image, table in table_images),
        "no_protected_inline_style_overrides": not protected_inline_override,
        "ba_compare_wrapped_in_text_block": all(
            node.ancestor(class_name="text-block") is not None for node in ba_compares
        ),
    }
    checks.update(marker_ladder_checks(root))
    checks.update(source_box_checks(root))
    checks.update(compare_and_play_checks(root))
    return checks


def body_care_dom_checks(root: Node) -> dict[str, bool]:
    intro = next((node for node in root.descendants("section") if node.has_class("intro-card")), None)
    overview_flat = False
    if intro is not None:
        main_parent = next(
            (
                node for node in intro.descendants("li")
                if any(child.tag == "b" and child.text() == "【主图】" for child in node.children)
            ),
            None,
        )
        if main_parent is not None and main_parent.parent is not None:
            siblings = _direct_elements(main_parent.parent, "li")
            try:
                index = siblings.index(main_parent)
            except ValueError:
                index = -1
            expected = ("首张主图：", "主图前5张：", "丰富素材图类型：")
            following = siblings[index + 1:index + 4] if index >= 0 else []
            overview_flat = (
                not any(child.tag == "ul" for child in main_parent.children)
                and len(following) == 3
                and all(node.has_class("sublevel") and node.text().startswith(prefix) for node, prefix in zip(following, expected))
            )

    main_tag = section_for_title(root, "主推标签")
    main_tag_contract = False
    if main_tag is not None:
        target = next(
            (node for node in main_tag.descendants("p") if node.text().startswith("商家可选择对各SPU下的主推荐SKU进行打标")),
            None,
        )
        main_tag_contract = bool(
            target is not None
            and target.parent is not None
            and target.parent.has_class("text-block")
            and target.parent.has_class("plain-block")
        )

    quality = section_for_title(root, "品质标签")
    quality_tables = [] if quality is None else [
        table for table in quality.descendants("table") if "品质标签示例" in table.text()
    ]
    quality_header = False
    quality_media = False
    if len(quality_tables) == 1:
        table = quality_tables[0]
        first_row = _first_row(table)
        cells = [] if first_row is None else [node for node in _direct_elements(first_row) if node.tag in {"th", "td"}]
        quality_header = bool(
            {"doc-table", "tag-example-table"}.issubset(table.classes)
            and table.parent is not None
            and table.parent.has_class("doc-table-wrap")
            and cells
            and all(cell.tag == "th" for cell in cells)
            and "品质标签示例" in cells[0].text()
        )
        images = table.descendants("img")
        quality_media = bool(images) and all(
            (cell := _closest_cell(image, table)) is not None
            and cell.has_class("table-media-cell")
            and image.parent is not None
            and image.parent.has_class("image-holder")
            for image in images
        )

    return {
        "overview_main_image_children_are_flat": overview_flat,
        "main_tag_body_uses_plain_prose_component": main_tag_contract,
        "quality_tag_example_has_real_table_header": quality_header,
        "quality_tag_example_uses_media_cell": quality_media,
    }


def evaluate_dom_contracts(html: str, profile: str | None) -> tuple[dict[str, bool], bool]:
    root = parse_html(html)
    is_body_care = detected_body_care(root)
    effective_profile = "body-care" if profile == "body-care" or (profile == "auto" and is_body_care) else None
    checks = generic_dom_checks(root)
    if effective_profile == "body-care":
        checks.update(body_care_dom_checks(root))
    return checks, effective_profile == "body-care"


def _self_test() -> None:
    good = f"""
    <section class="intro-card"><ul class="red-list">
      <li><b>【主图】</b></li>
      <li class="sublevel">首张主图：内容</li>
      <li class="sublevel">主图前5张：内容</li>
      <li class="sublevel">丰富素材图类型：内容</li>
    </ul></section>
    <section><h2>{{ 主推标签 }}</h2><div class="text-block plain-block"><p>商家可选择对各SPU下的主推荐SKU进行打标，内容</p></div></section>
    <section><h2>{{ 品质标签 }}</h2><div class="doc-table-wrap"><table class="doc-table tag-example-table">
      <tr><th>品质标签示例</th></tr><tr><td class="table-media-cell"><div class="image-holder"><img></div></td></tr>
    </table></div></section>
    <section><h2>{{ 标题 }}</h2><div class="text-block"><div class="ba-compare">
      <div class="ba-col"><div class="ba-head ba-before">优化前</div><p class="ba-text">原始标题文案</p></div>
      <div class="ba-col"><div class="ba-head ba-after">优化后</div><p class="ba-text">优化后标题文案</p></div>
    </div></div></section>
    <section><h2>{{ 层级 }}</h2><div class="text-block">
      <div class="label-line"><span class="label-text">主图基础要求：</span></div>
      <ul class="source-list">
        <li>1. 视觉风格：低饱和配色
          <ul class="source-list"><li class="deep">a. 贴片颜色与背景贴合
            <ul class="dot-list"><li>补充说明</li></ul>
          </li></ul>
        </li>
        <li>2. 信息逻辑：精简聚焦</li>
      </ul>
    </div></section>
    <section><h2>{{ 视频 }}</h2><div class="text-block"><div class="video-case-grid video-case-card">
      <div class="video-case-copy"><div class="video-case-head">内容结构</div><div class="video-case-body"><p class="ba-text">开场</p></div></div>
      <div class="video-case-media"><div class="video-case-head">示例</div><div class="video-demo"><span class="vd-text">点击播放</span></div></div>
    </div></div></section>
    """
    bad = f"""
    <section class="intro-card"><ul class="red-list"><li><b>【主图】</b><ul class="red-list">
      <li class="sublevel">首张主图：内容</li><li class="sublevel">主图前5张：内容</li>
      <li class="sublevel">丰富素材图类型：内容</li></ul></li></ul></section>
    <section><h2>{{ 主推标签 }}</h2><div class="text-block"><p>商家可选择对各SPU下的主推荐SKU进行打标，内容</p></div></section>
    <section><h2>{{ 品质标签 }}</h2><table class="tag-example-table"><tr><th>品质标签示例</th></tr>
      <tr><td class="attr-img"><div class="image-holder"><img></div></td></tr></table></section>
    <section><h2>{{ 标题 }}</h2><div class="ba-compare">
      <div class="ba-col"><div class="ba-head ba-before">优化前</div><p>原始标题文案</p></div>
      <div class="ba-col"><div class="ba-head ba-after">优化后</div><p>优化后标题文案</p></div>
    </div></section>
    <section><h2>{{ 层级 }}</h2><div class="text-block">
      <ul class="red-list">
        <li><b>主图基础要求：</b></li>
        <li>1. 视觉风格：低饱和配色</li>
        <li>a. 贴片颜色须与背景色调贴合</li>
        <li>2. 信息逻辑：精简聚焦</li>
      </ul>
    </div></section>
    <section><h2>{{ 拆框 }}</h2>
      <div class="doc-table-wrap"><table class="doc-table"><tr><th>视觉风格：贴片颜色须与背景色调贴合</th></tr></table></div>
      <div class="doc-table-wrap"><table class="doc-table"><tr><th>破坏整体质感，保持画面干净舒适。</th></tr></table></div>
    </section>
    <section><h2>{{ 视频 }}</h2><div class="text-block"><div class="video-demo"><span class="vd-text">点击播放</span></div></div></section>
    """
    good_root = parse_html(good)
    good_checks = {**generic_dom_checks(good_root), **body_care_dom_checks(good_root)}
    if not all(good_checks.values()):
        raise AssertionError({key: value for key, value in good_checks.items() if not value})
    bad_root = parse_html(bad)
    bad_checks = {**generic_dom_checks(bad_root), **body_care_dom_checks(bad_root)}
    required_failures = {
        "no_nested_red_list",
        "plain_prose_uses_plain_block",
        "tag_example_table_component_contract",
        "table_images_use_semantic_media_cells",
        "ba_compare_wrapped_in_text_block",
        "overview_main_image_children_are_flat",
        "main_tag_body_uses_plain_prose_component",
        "quality_tag_example_has_real_table_header",
        "quality_tag_example_uses_media_cell",
        "ba_compare_text_uses_table_body_cells",
        "play_card_matches_neighbour_cell_box",
        "no_single_cell_doc_table",
        "source_box_not_fragmented_into_stacked_cards",
        "numbered_children_are_demoted_one_level",
        "lettered_children_are_demoted_below_numbers",
    }
    if any(bad_checks.get(name) for name in required_failures):
        raise AssertionError({name: bad_checks.get(name) for name in sorted(required_failures)})
    print("DOM component contract self-test passed")


if __name__ == "__main__":
    _self_test()
