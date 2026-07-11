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


@dataclass
class Node:
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

    protected = {"sublevel", "plain-block", "tag-example-table", "table-media-cell", "doc-table-wrap"}
    protected_inline_override = any(
        node.classes.intersection(protected)
        and re.search(r"(?:font-size|padding|margin-left|text-align)\s*:", node.attrs.get("style", ""), re.I)
        for node in root.descendants()
    )

    return {
        "no_nested_red_list": not any(node.ancestor(tag="ul", class_name="red-list") for node in red_lists),
        "plain_prose_uses_plain_block": all(block.has_class("plain-block") for block in text_only_blocks),
        "tag_example_table_component_contract": not tag_tables
        or all(_tag_example_contract(table) for table in tag_tables),
        "table_images_use_semantic_media_cells": all(media_cell_matches_table(image, table) for image, table in table_images),
        "no_protected_inline_style_overrides": not protected_inline_override,
    }


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
    """
    bad = f"""
    <section class="intro-card"><ul class="red-list"><li><b>【主图】</b><ul class="red-list">
      <li class="sublevel">首张主图：内容</li><li class="sublevel">主图前5张：内容</li>
      <li class="sublevel">丰富素材图类型：内容</li></ul></li></ul></section>
    <section><h2>{{ 主推标签 }}</h2><div class="text-block"><p>商家可选择对各SPU下的主推荐SKU进行打标，内容</p></div></section>
    <section><h2>{{ 品质标签 }}</h2><table class="tag-example-table"><tr><th>品质标签示例</th></tr>
      <tr><td class="attr-img"><div class="image-holder"><img></div></td></tr></table></section>
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
        "overview_main_image_children_are_flat",
        "main_tag_body_uses_plain_prose_component",
        "quality_tag_example_has_real_table_header",
        "quality_tag_example_uses_media_cell",
    }
    if any(bad_checks.get(name) for name in required_failures):
        raise AssertionError({name: bad_checks.get(name) for name in sorted(required_failures)})
    print("DOM component contract self-test passed")


if __name__ == "__main__":
    _self_test()
