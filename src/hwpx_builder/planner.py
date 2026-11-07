"""Conservative flow checks. Physical pagination remains the renderer's responsibility."""

from .model import Composition, Document, Picture, Table


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
