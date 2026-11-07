"""Conservative flow checks. Physical pagination remains the renderer's responsibility."""

from dataclasses import replace
from decimal import Decimal

from .model import Composition, Document, Paragraph, Picture, Table, TableData


def table_horizontal_bounds(data: TableData, body_width: Decimal) -> tuple[Decimal, Decimal]:
    """Visible border bounds in the coordinate frame serialized by the table writer."""
    width = data.size[0] if data.size is not None else sum(data.widths)
    if data.inline:
        # An inline object occupies its outer margins inside the host's first line.
        left = data.host_style.indent + data.host_style.first_line_indent
        right = body_width - data.host_style.right_indent
        leading, trailing = data.margin[:2]
    else:
        # Floating resource tables are anchored to COLUMN; paragraph margins and
        # wrap clearance do not move their border within that coordinate frame.
        left, right = Decimal(0), body_width
        leading = trailing = Decimal(0)
    outer_width = width + leading + trailing
    if data.host_style.align == "center":
        left += (right - left - outer_width) / 2
    elif data.host_style.align == "right":
        left = right - outer_width
    left += leading
    return left, left + width


def constrain_table_note(note: Paragraph, data: TableData, body_width: Decimal) -> Paragraph:
    left, right = table_horizontal_bounds(data, body_width)
    style = replace(
        note.style,
        indent=max(note.style.indent, left),
        right_indent=max(note.style.right_indent, body_width - right),
    )
    if style.indent + style.right_indent + style.first_line_indent >= body_width:
        note.source.fail(
            "TABLE_NOTE_NO_SPACE",
            "표 설명의 들여쓰기와 여백이 표 안의 가용 폭을 소진합니다. indent_depth와 paragraph_style을 조정하세요.",
        )
    return replace(note, style=style)


def plan(document: Document) -> Document:
    for section in document.sections:
        for block in section.contents:
            width = (
                sum(block.data.widths)
                if isinstance(block, Table)
                else block.width
                if isinstance(block, (Picture, Composition))
                else None
            )
            if width is not None and width > section.page.body_width:
                block.source.fail(
                    "CONTENT_TOO_WIDE",
                    f"콘텐츠 너비 {width:.2f}mm가 본문 {section.page.body_width:.2f}mm를 초과합니다. 원본/용지 여백을 조정하세요.",
                )
            if (
                isinstance(block, (Picture, Composition))
                and block.height > section.page.height - section.page.top - section.page.bottom
            ):
                block.source.fail("CONTENT_TOO_HIGH", "이미지 높이가 본문 영역을 초과합니다.")
    return document
