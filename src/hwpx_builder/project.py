"""Project-root paths, safe declaration loading, schema validation and AST resolution."""

from __future__ import annotations

import re
import tomllib
from dataclasses import dataclass, replace
from decimal import Decimal
from pathlib import Path

import yaml
from yaml.events import AliasEvent
from yaml.nodes import MappingNode, SequenceNode

from .model import (
    Composition,
    Document,
    Page,
    PageBreak,
    Paragraph,
    ParagraphStyle,
    Section,
    SectionStyle,
    Source,
    Table,
    TextBox,
    TextStyle,
)
from .planner import constrain_table_note
from .resources import Resources
from .text import color, parse_text

D = Decimal


def _pointer(parent, key):
    return parent + "/" + str(key).replace("~", "~0").replace("/", "~1")


@dataclass(frozen=True)
class Node:
    value: object
    source: Source
    marks: dict
    stack: tuple[Path, ...] = ()
    key_marks: dict | None = None

    def child(self, key):
        pointer = _pointer(self.source.pointer, key)
        line, column = self.marks.get(pointer, (self.source.line, self.source.column))
        value = self.value[key] if isinstance(self.value, list) or key in self.value else None
        return Node(
            value,
            Source(self.source.path, pointer, line, column, self.source.reference_chain),
            self.marks,
            self.stack,
            self.key_marks,
        )

    def obj(self, allowed, required=()):
        if not isinstance(self.value, dict):
            self.source.fail("INVALID_VALUE", "객체가 필요합니다.")
        for key in self.value:
            if key not in allowed:
                origin = self.child(key).source
                if self.key_marks and origin.pointer in self.key_marks:
                    origin = replace(
                        origin,
                        line=self.key_marks[origin.pointer][0],
                        column=self.key_marks[origin.pointer][1],
                    )
                origin.fail("UNKNOWN_FIELD", f"알 수 없는 속성: {key}")
        for key in required:
            if key not in self.value or self.value[key] is None:
                self.child(key).source.fail("MISSING_REQUIRED", f"필수 속성: {key}")
        return self.value

    def array(self):
        if not isinstance(self.value, list):
            self.source.fail("INVALID_VALUE", "배열이 필요합니다.")
        return [self.child(i) for i in range(len(self.value))]


def number(value, source, positive=False):
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        source.fail("INVALID_VALUE", "단위 없는 숫자가 필요합니다.")
    result = D(str(value))
    if not result.is_finite() or (result <= 0 if positive else result < 0):
        source.fail("INVALID_VALUE", "양수 또는 허용된 0 값이 필요합니다.")
    return result


def choice(value, options, source):
    if value not in options:
        source.fail("INVALID_VALUE", f"허용 값: {', '.join(options)}")
    return value


class ProjectLoader:
    def __init__(self, root, *, resources=None):
        self.root = Path(root).resolve()
        self.resources_override = resources
        self.files = set()
        self.count = 0

    def file(self, raw, expected, stack=(), reference=None, *, external_resources=False):
        chain = (
            ()
            if reference is None
            else (
                *reference.reference_chain,
                {
                    "source_path": str(reference.path),
                    "pointer": reference.pointer,
                    "line": reference.line,
                    "column": reference.column,
                },
            )
        )

        def origin(path, **kwargs):
            return Source(path, reference_chain=chain, **kwargs)

        if not isinstance(raw, (str, Path)) or raw == "":
            (reference or origin(self.root / "document.yaml")).fail(
                "INVALID_VALUE", "선언 파일 경로는 비어 있지 않은 문자열이어야 합니다."
            )
        path = (self.root / Path(raw)).resolve()
        if not path.is_relative_to(self.root) and not (external_resources and expected == "resources"):
            (reference or origin(path)).fail(
                "DECLARATION_OUTSIDE_ROOT", "선언 파일이 프로젝트 루트를 벗어났습니다."
            )
        if path in stack:
            (reference or origin(path)).fail("CYCLIC_INCLUDE", "순환 선언 참조입니다.")
        if len(stack) >= 32:
            (reference or origin(path)).fail("DECLARATION_LIMIT", "선언 참조 깊이 제한을 넘었습니다.")
        try:
            text = path.read_text(encoding="utf-8")
        except (OSError, UnicodeError) as exc:
            (reference or origin(path)).fail("DECLARATION_READ_FAILED", str(exc))
        if len(text) > 4 * 1024 * 1024:
            (reference or origin(path)).fail("DECLARATION_LIMIT", "선언 파일 크기 제한을 넘었습니다.")
        marks = {}
        key_marks = {}
        if path.suffix.lower() == ".toml":
            try:
                data = tomllib.loads(text)
            except tomllib.TOMLDecodeError as exc:
                mark = getattr(exc, "problem_mark", None)
                origin(
                    path,
                    line=mark.line + 1 if mark else getattr(exc, "lineno", None),
                    column=mark.column + 1 if mark else getattr(exc, "colno", None),
                ).fail("INVALID_DECLARATION", str(exc))
        elif path.suffix.lower() in (".yaml", ".yml"):

            class SafeLoader(yaml.SafeLoader):
                def construct_mapping(self, node, deep=False):
                    result = {}
                    for key_node, value_node in node.value:
                        key = self.construct_object(key_node, deep=deep)
                        if not isinstance(key, str):
                            origin(
                                path, line=key_node.start_mark.line + 1, column=key_node.start_mark.column + 1
                            ).fail("INVALID_VALUE", "매핑 키는 문자열이어야 합니다.")
                        if key in result:
                            origin(
                                path, line=key_node.start_mark.line + 1, column=key_node.start_mark.column + 1
                            ).fail("DUPLICATE_KEY", f"중복 키: {key}")
                        result[key] = self.construct_object(value_node, deep=deep)
                    return result

            try:
                for event in yaml.parse(text):
                    if isinstance(event, AliasEvent):
                        origin(path, line=event.start_mark.line + 1, column=event.start_mark.column + 1).fail(
                            "UNSUPPORTED_DECLARATION", "YAML alias 대신 src 참조를 사용하세요."
                        )
                tree = yaml.compose(text, Loader=SafeLoader)

                def visit(node, pointer=""):
                    marks[pointer] = (node.start_mark.line + 1, node.start_mark.column + 1)
                    if isinstance(node, MappingNode):
                        for k, v in node.value:
                            child_pointer = _pointer(pointer, k.value)
                            key_marks[child_pointer] = (k.start_mark.line + 1, k.start_mark.column + 1)
                            visit(v, child_pointer)
                    elif isinstance(node, SequenceNode):
                        for i, v in enumerate(node.value):
                            visit(v, _pointer(pointer, i))

                if tree is not None:
                    visit(tree)
                data = yaml.load(text, Loader=SafeLoader)
            except yaml.YAMLError as exc:
                mark = getattr(exc, "problem_mark", None)
                origin(
                    path,
                    line=mark.line + 1 if mark else getattr(exc, "lineno", None),
                    column=mark.column + 1 if mark else getattr(exc, "colno", None),
                ).fail("INVALID_DECLARATION", str(exc))
        else:
            origin(path).fail("INVALID_DECLARATION", "선언은 YAML 또는 TOML이어야 합니다.")
        source = Source(path, "", *marks.get("", (None, None)), chain)
        node = Node(data, source, marks, (*stack, path), key_marks)
        if not isinstance(data, dict):
            source.fail("INVALID_VALUE", "선언 루트는 객체여야 합니다.")
        if data.get("schema_version") != "1.0":
            node.child("schema_version").source.fail(
                "INVALID_SCHEMA_VERSION", 'schema_version: "1.0"이 필요합니다.'
            )
        if data.get("type") != expected:
            node.child("type").source.fail("TYPE_MISMATCH", f"type: {expected}이 필요합니다.")
        self.files.add(path)
        return replace(node, value={k: v for k, v in data.items() if k not in ("schema_version", "type")})

    def ref(self, node, expected):
        if isinstance(node.value, dict) and "src" in node.value:
            if set(node.value) != {"src"}:
                node.source.fail("SRC_INLINE_CONFLICT", "src와 인라인 정의를 혼합할 수 없습니다.")
            return self.file(node.value["src"], expected, node.stack, node.child("src").source)
        return node

    def text_style(self, node, parent):
        if node.value is None:
            return parent
        node = self.ref(node, "text_style")
        d = node.obj(
            {"font_family", "font_size", "bold", "italic", "underline", "strike", "color", "highlight"}
        )
        changes = {}
        for key, value in d.items():
            if value is None:
                continue
            origin = node.child(key).source
            if key in ("bold", "italic", "underline", "strike"):
                if not isinstance(value, bool):
                    origin.fail("INVALID_VALUE", f"{key}는 boolean이어야 합니다.")
            elif key == "font_size":
                value = number(value, origin, True)
            elif key in ("color", "highlight"):
                value = color(value, origin)
            elif not isinstance(value, str) or not value.strip():
                origin.fail("INVALID_VALUE", "글꼴 이름이 필요합니다.")
            changes[key] = value
        return replace(parent, **changes)

    def paragraph_style(self, node, parent):
        if node.value is None:
            return parent
        node = self.ref(node, "paragraph_style")
        d = node.obj(
            {
                "align",
                "right_indent",
                "first_line_indent",
                "space_before",
                "space_after",
                "line_spacing",
                "keep_with_next",
                "keep_together",
                "tab_stops",
            }
        )
        changes = {}
        for key, value in d.items():
            if value is None:
                continue
            source = node.child(key).source
            if key == "align":
                value = choice(value, ("left", "center", "right", "justify"), source)
            elif key == "line_spacing":
                if not isinstance(value, str) or not re.fullmatch(r"[1-9]\d*%", value):
                    source.fail("INVALID_VALUE", "줄 간격은 160% 형식이어야 합니다.")
                value = int(value[:-1])
            elif key in ("keep_with_next", "keep_together"):
                if not isinstance(value, bool):
                    source.fail("INVALID_VALUE", "boolean이 필요합니다.")
            elif key == "tab_stops":
                source.fail("UNSUPPORTED_RENDERING", "탭 위치는 아직 지원하지 않습니다.")
            else:
                value = number(value, source)
            changes[key] = value
        return replace(parent, **changes)

    def section_style(self, node, parent):
        if node.value is None:
            return parent
        node = self.ref(node, "section_style")
        data = node.obj({"indent_by_depth"})
        if data.get("indent_by_depth") is None:
            return parent
        values = node.child("indent_by_depth")
        levels = tuple(number(item.value, item.source) for item in values.array())
        if not levels:
            values.source.fail("INVALID_VALUE", "indent_by_depth는 비어 있을 수 없습니다.")
        for i in range(1, len(levels)):
            if levels[i] < levels[i - 1]:
                values.child(i).source.fail("INVALID_VALUE", "들여쓰기는 깊이에 따라 감소할 수 없습니다.")
        return SectionStyle(levels)

    def page(self, node, parent):
        if node.value is None:
            return parent
        node = self.ref(node, "page_setup")
        d = node.obj({"size", "orientation", "margin", "header_margin", "footer_margin", "gutter"})
        changes = {}
        if d.get("size") is not None:
            sizes = {
                "A3": (297, 420),
                "A4": (210, 297),
                "A5": (148, 210),
                "B4": (250, 353),
                "B5": (176, 250),
                "Letter": (215.9, 279.4),
                "Legal": (215.9, 355.6),
            }
            size = d["size"]
            if isinstance(size, dict):
                size_node = node.child("size")
                size_node.obj({"width", "height"}, ("width", "height"))
                changes["width"] = number(size["width"], size_node.child("width").source, True)
                changes["height"] = number(size["height"], size_node.child("height").source, True)
            elif isinstance(size, str) and size in sizes:
                changes["width"], changes["height"] = (D(str(x)) for x in sizes[size])
            else:
                node.child("size").source.fail("INVALID_VALUE", "올바르지 않은 용지 크기입니다.")
        if d.get("orientation") is not None:
            direction = choice(d["orientation"], ("portrait", "landscape"), node.child("orientation").source)
            w, h = changes.get("width", parent.width), changes.get("height", parent.height)
            changes["width"], changes["height"] = (
                (max(w, h), min(w, h)) if direction == "landscape" else (min(w, h), max(w, h))
            )
        if d.get("margin") is not None:
            m = node.child("margin")
            m.obj({"left", "right", "top", "bottom"})
            for k, v in m.value.items():
                if v is not None:
                    changes[k] = number(v, m.child(k).source)
        for key, dest in [("header_margin", "header"), ("footer_margin", "footer")]:
            if d.get(key) is not None:
                changes[dest] = number(d[key], node.child(key).source)
        if d.get("gutter") is not None:
            g = node.child("gutter")
            g.obj({"width", "position"})
            if g.value.get("width") is not None:
                changes["gutter"] = number(g.value["width"], g.source)
            if g.value.get("position") is not None:
                changes["gutter_position"] = choice(g.value["position"], ("left", "right", "top"), g.source)
        page = replace(parent, **changes)
        if page.gutter_position != "left":
            node.source.fail("UNSUPPORTED_RENDERING", "제본 여백은 현재 left만 지원합니다.")
        if (
            page.body_width <= 0
            or page.height - page.top - page.bottom - (page.gutter if page.gutter_position == "top" else 0)
            <= 0
        ):
            node.source.fail("INVALID_VALUE", "페이지의 인쇄 가능한 영역이 없습니다.")
        return page

    def paragraph(self, node, text_style):
        d = node.obj({"type", "text", "runs", "text_style", "paragraph_style", "indent_depth"})
        if "type" in d and d["type"] != "paragraph":
            node.child("type").source.fail("TYPE_MISMATCH", "설명 문단의 type은 paragraph여야 합니다.")
        if ("text" in d) == ("runs" in d):
            node.source.fail("INVALID_VALUE", "text 또는 runs 중 하나만 지정하세요.")
        style = self.text_style(node.child("text_style"), text_style)
        ps = self.paragraph_style(node.child("paragraph_style"), self.default_paragraph)
        depth = d.get("indent_depth", 0)
        if type(depth) is not int or not 0 <= depth < len(self.current_section.indent_by_depth):
            node.child("indent_depth").source.fail(
                "INVALID_INDENT_DEPTH",
                f"indent_depth는 0~{len(self.current_section.indent_by_depth) - 1} 정수여야 합니다.",
            )
        ps = replace(ps, indent=self.current_section.indent_by_depth[depth])
        if "text" in d:
            runs = parse_text(d["text"], style, self.variables, node.child("text").source)
        else:
            runs = []
            for run in node.child("runs").array():
                run.obj({"text", "text_style"}, ("text",))
                rs = self.text_style(run.child("text_style"), style)
                runs.extend(parse_text(run.value["text"], rs, self.variables, run.child("text").source))
            runs = tuple(runs)
        return Paragraph(runs, ps, node.source, depth)

    def blocks(self, node, text_style, parent="layer"):
        result = []
        for block in node.array():
            self.count += 1
            if self.count > 100000:
                block.source.fail("DECLARATION_LIMIT", "블록 개수 제한을 넘었습니다.")
            if not isinstance(block.value, dict):
                block.source.fail("INVALID_VALUE", "블록 객체가 필요합니다.")
            kind = block.value.get("type")
            if kind == "paragraph":
                result.append(self.paragraph(block, text_style))
            elif kind == "include":
                block.obj({"type", "src"}, ("src",))
                included = self.file(block.value["src"], "blocks", block.stack, block.child("src").source)
                included.obj({"contents"}, ("contents",))
                result.extend(self.blocks(included.child("contents"), text_style, parent))
            elif kind == "table":
                if "resource" not in block.value:
                    block.source.fail(
                        "TABLE_RESOURCE_REQUIRED", "표는 완성 XLSX resource만 사용할 수 있습니다."
                    )
                if any(
                    k in block.value for k in ("rows", "text_style", "table_style", "headnote", "footnote")
                ):
                    block.source.fail(
                        "RESOURCE_TABLE_OVERRIDE_FORBIDDEN",
                        "완성 표의 내부 값·서식은 선언에서 덮어쓸 수 없습니다.",
                    )
                block.obj({"type", "resource", "headernote", "footernote"}, ("resource",))
                name = block.value["resource"]
                if not isinstance(name, str):
                    block.source.fail("INVALID_VALUE", "resource 이름은 문자열이어야 합니다.")
                data = self.resources.table(name, block.source)
                for key in ("headernote",):
                    if block.value.get(key) is not None:
                        result.append(
                            constrain_table_note(
                                self.paragraph(block.child(key), text_style),
                                data,
                                self.current_page.body_width,
                            )
                        )
                result.append(Table(data, block.source))
                if block.value.get("footernote") is not None:
                    result.append(
                        constrain_table_note(
                            self.paragraph(block.child("footernote"), text_style),
                            data,
                            self.current_page.body_width,
                        )
                    )
            elif kind == "container":
                block.obj(
                    {"type", "contents", "text_style", "container_style", "headernote", "footernote"},
                    ("contents",),
                )
                cs = block.child("container_style")
                if cs.value is not None:
                    cs = self.ref(cs, "container_style")
                    cs.obj({"layout"})
                    if cs.value.get("layout") is not None:
                        layout = cs.child("layout")
                        layout.obj({"mode", "relative_to", "x", "y", "width", "height", "wrap", "z_index"})
                        if layout.value.get("mode", "inline") == "inline" and (
                            "width" in layout.value or "height" in layout.value
                        ):
                            if (
                                block.value.get("headernote") is not None
                                or block.value.get("footernote") is not None
                            ):
                                block.source.fail(
                                    "UNSUPPORTED_RENDERING",
                                    "고정 Container의 설명은 별도 문단으로 선언하세요.",
                                )
                            result.append(self.composition(block, layout, text_style))
                            continue
                        if layout.value != {"mode": "inline"}:
                            layout.source.fail(
                                "UNSUPPORTED_RENDERING", "현재 Container는 inline 흐름만 지원합니다."
                            )
                inherited = self.text_style(block.child("text_style"), text_style)
                if block.value.get("headernote") is not None:
                    result.append(self.paragraph(block.child("headernote"), inherited))
                result.extend(self.blocks(block.child("contents"), inherited, "container"))
                if block.value.get("footernote") is not None:
                    result.append(self.paragraph(block.child("footernote"), inherited))
            elif kind == "image":
                if parent != "container":
                    block.source.fail("IMAGE_REQUIRES_CONTAINER", "Image의 직접 부모는 Container여야 합니다.")
                block.obj({"type", "resource", "image_style"}, ("resource",))
                style = block.child("image_style")
                values = {}
                if style.value is not None:
                    style = self.ref(style, "image_style")
                    style.obj({"width", "height", "fit", "align", "opacity", "rotation"})
                    for k, v in style.value.items():
                        if v is None:
                            continue
                        if k in ("width", "height"):
                            values[k] = number(v, style.child(k).source, True)
                        elif k == "fit":
                            values[k] = choice(v, ("contain", "cover", "fill", "none"), style.source)
                        elif k == "align":
                            values[k] = choice(v, ("left", "center", "right"), style.source)
                        elif (k == "opacity" and v != 1) or (k == "rotation" and v != 0):
                            style.source.fail(
                                "UNSUPPORTED_RENDERING", "투명도/회전은 아직 지원하지 않습니다."
                            )
                name = block.value["resource"]
                if not isinstance(name, str):
                    block.source.fail("INVALID_VALUE", "resource 이름은 문자열이어야 합니다.")
                result.append(self.resources.picture(name, values, block.source))
            elif kind == "pagebreak":
                block.obj({"type"})
                result.append(PageBreak(block.source))
            elif kind in ("repeat", "condition", "field"):
                block.source.fail("UNSUPPORTED_RENDERING", f"현재 지원하지 않는 블록: {kind}")
            else:
                block.source.fail("INVALID_TYPE", f"알 수 없는 블록 type: {kind}")
        return result

    def composition(self, block, layout, text_style):
        layout.obj({"mode", "width", "height"}, ("width", "height"))
        width = number(layout.value["width"], layout.child("width").source, True)
        height = number(layout.value["height"], layout.child("height").source, True)
        inherited = self.text_style(block.child("text_style"), text_style)
        contents = []
        for child in block.child("contents").array():
            self.count += 1
            if self.count > 100000:
                child.source.fail("DECLARATION_LIMIT", "블록 개수 제한을 넘었습니다.")
            if not isinstance(child.value, dict):
                child.source.fail("INVALID_VALUE", "블록 객체가 필요합니다.")
            if child.value.get("type") == "image":
                image = self.blocks(replace(child, value=[child.value]), inherited, "container")[0]
                if image.width > width or image.height > height:
                    child.source.fail("CONTENT_TOO_WIDE", "Container보다 큰 이미지입니다.")
                if image.align != "left":
                    child.source.fail(
                        "UNSUPPORTED_RENDERING", "고정 Container 이미지는 좌측 상단에 배치합니다."
                    )
                contents.append(image)
                continue
            child.obj({"type", "contents", "text_style", "container_style"}, ("contents", "container_style"))
            if child.value["type"] != "container":
                child.source.fail(
                    "UNSUPPORTED_RENDERING", "고정 Container에는 Image와 문단 전용 Container만 배치합니다."
                )
            style = self.ref(child.child("container_style"), "container_style")
            style.obj({"layout", "padding"}, ("layout",))
            position = style.child("layout")
            position.obj(
                {"mode", "relative_to", "x", "y", "width", "height", "wrap", "z_index"},
                ("mode", "x", "y", "width", "height"),
            )
            p = position.value
            if (
                p["mode"] != "floating"
                or p.get("relative_to", "paragraph") != "paragraph"
                or p.get("wrap", "none") != "none"
                or p.get("z_index", 1) != 1
            ):
                position.source.fail(
                    "UNSUPPORTED_RENDERING",
                    "고정 Container의 자식은 paragraph 기준 floating, wrap none, z_index 1만 지원합니다.",
                )
            x, y, w, h = (
                number(p[k], position.child(k).source, k in ("width", "height"))
                for k in ("x", "y", "width", "height")
            )
            if x + w > width or y + h > height:
                position.source.fail("CONTENT_TOO_WIDE", "자식 영역이 Container를 벗어납니다.")
            padding = style.child("padding")
            if padding.value is not None:
                padding.obj({"left", "right", "top", "bottom"})
            pads = tuple(
                number(padding.value.get(k, 0), padding.child(k).source) if padding.value else D(0)
                for k in ("left", "right", "top", "bottom")
            )
            if pads[0] + pads[1] >= w or pads[2] + pads[3] >= h:
                padding.source.fail("INVALID_VALUE", "내부 여백이 Container 영역을 소진합니다.")
            ts = self.text_style(child.child("text_style"), inherited)
            paragraphs = tuple(self.paragraph(p, ts) for p in child.child("contents").array())
            if len(paragraphs) != 1:
                child.source.fail("UNSUPPORTED_RENDERING", "고정 글상자는 현재 한 문단만 지원합니다.")
            # Explicit line breaks and impossible glyph heights cannot fit this subset.
            text = "".join(r.text for r in paragraphs[0].runs)
            if any(r.linebreak for r in paragraphs[0].runs) or "\n" in text:
                child.source.fail("UNSUPPORTED_RENDERING", "고정 글상자의 줄바꿈은 아직 지원하지 않습니다.")
            size = max((r.style.font_size for r in paragraphs[0].runs), default=D(10)) * D("25.4") / 72
            if size > h - pads[2] - pads[3]:
                child.source.fail("CONTENT_TOO_HIGH", "글자 높이가 글상자를 초과합니다.")
            contents.append(TextBox(paragraphs, x, y, w, h, pads))
        return Composition(tuple(contents), width, height, block.source)

    def resource_set(self, resources, override=None):
        """Resolve either resource set with its own path boundary and source locations."""
        resource_root = self.root
        if override is not None:
            resources = self.file(override, "resources", resources.stack, external_resources=True)
            resource_root = resources.source.path.parent
        elif isinstance(resources.value, dict) and "file_path" in resources.value:
            if set(resources.value) != {"file_path"}:
                resources.source.fail("SRC_INLINE_CONFLICT", "file_path와 다른 속성을 혼합할 수 없습니다.")
            resources = self.file(
                resources.value["file_path"],
                "resources",
                resources.stack,
                resources.child("file_path").source,
                external_resources=True,
            )
            resource_root = resources.source.path.parent
        else:
            if resources.value is None:
                resources = replace(resources, value={})
            resources = self.ref(resources, "resources")
        resources.obj({"tables", "images"})
        registries = {}
        for category in ("tables", "images"):
            registry = resources.child(category)
            definitions = {}
            if registry.value is not None:
                if not isinstance(registry.value, dict):
                    registry.source.fail("INVALID_VALUE", "리소스 이름 객체가 필요합니다.")
                for name in registry.value:
                    item = registry.child(name)
                    item.obj(
                        {"path", "sheet", "range", "header"} if category == "tables" else {"path"}, ("path",)
                    )
                    if "header" in item.value and not isinstance(item.value["header"], bool):
                        item.source.fail("INVALID_VALUE", "header는 boolean이어야 합니다.")
                    if "sheet" in item.value and not isinstance(item.value["sheet"], str):
                        item.source.fail("INVALID_VALUE", "sheet는 문자열이어야 합니다.")
                    path = Resources(resource_root, {}, {}).path(item.value, item.source)
                    definitions[name] = ({**item.value, "path": str(path)}, item.source)
            registries[category] = definitions
        return registries

    def load(self):
        root = self.file("document.yaml", "document")
        root.obj(
            {"defaults", "assets", "resources", "variables", "sections", "headers", "footers", "page_number"},
            ("sections",),
        )
        for key in ("headers", "footers", "page_number"):
            if root.value.get(key) not in (None, []):
                root.child(key).source.fail(
                    "UNSUPPORTED_RENDERING", f"{key}는 물리 페이지 검증기 구현 전까지 지원하지 않습니다."
                )
        defaults = root.child("defaults")
        if defaults.value is None:
            defaults = replace(defaults, value={})
        defaults.obj({"page", "text_style", "paragraph_style", "section_style"})
        default_section = self.section_style(defaults.child("section_style"), SectionStyle())
        self.default_paragraph = self.paragraph_style(defaults.child("paragraph_style"), ParagraphStyle())
        text_style = self.text_style(defaults.child("text_style"), TextStyle())
        page = self.page(defaults.child("page"), Page())
        registries = {"tables": {}, "images": {}}
        for definitions in (
            self.resource_set(root.child("assets")),
            self.resource_set(root.child("resources"), self.resources_override),
        ):
            for category, registry in definitions.items():
                for name, (definition, origin) in registry.items():
                    if name in registries[category]:
                        previous = registries[category][name][1]
                        origin.fail(
                            "DUPLICATE_RESOURCE",
                            f"리소스 이름이 중복되었습니다: {category}.{name} (기존 정의: {previous.path}{previous.pointer})",
                        )
                    registries[category][name] = (definition, origin)
        self.resources = Resources(self.root, registries["tables"], registries["images"])
        # All paths were confined within their own source's root before conversion
        # to absolute paths. Protect inputs from both sets, including unused assets.
        for registry in registries.values():
            for definition, origin in registry.values():
                self.resources.path(definition, origin)
        self.variables = {}
        variables = root.child("variables")
        if variables.value is not None:
            if not isinstance(variables.value, dict) or set(variables.value) != {"resource"}:
                variables.source.fail(
                    "VARIABLE_RESOURCE_REQUIRED", "variables는 {resource: 2열표이름} 형식만 허용합니다."
                )
            name = variables.value["resource"]
            if not isinstance(name, str):
                variables.source.fail("INVALID_VALUE", "resource는 문자열이어야 합니다.")
            self.variables = self.resources.variables(name, variables.source)
        sections = []
        for item in root.child("sections").array():
            item = self.ref(item, "section")
            item.obj({"page", "text_style", "layer", "page_number", "section_style"}, ("layer",))
            if item.value.get("page_number") is not None:
                item.source.fail("UNSUPPORTED_RENDERING", "Section page_number는 아직 지원하지 않습니다.")
            section_page = self.page(item.child("page"), page)
            if sections and section_page != sections[0].page:
                item.source.fail(
                    "UNSUPPORTED_RENDERING", "흐름 중 페이지 설정 변경은 아직 지원하지 않습니다."
                )
            layer = self.ref(item.child("layer"), "layer")
            layer.obj({"type", "contents"}, ("type", "contents"))
            if layer.value["type"] != "single_column":
                layer.source.fail("UNSUPPORTED_RENDERING", "single_column Layer만 지원합니다.")
            self.current_page = section_page
            self.current_section = self.section_style(item.child("section_style"), default_section)
            inherited = self.text_style(item.child("text_style"), text_style)
            sections.append(
                Section(
                    tuple(self.blocks(layer.child("contents"), inherited)),
                    section_page,
                    item.source,
                    self.current_section,
                )
            )
        # Empty sections still use the declared document page.
        if not sections:
            sections = [Section((), page, root.source, default_section)]
        return Document(
            self.root,
            tuple(sections),
            tuple(self.variables.items()),
            tuple(self.resources.diagnostics),
            tuple(sorted(self.files | self.resources.inputs)),
        )


def load_project(root: str | Path, *, resources: str | Path | None = None) -> Document:
    """Load a template, optionally replacing resources with a document-root-relative file."""
    return ProjectLoader(root, resources=resources).load()
