"""Immutable normalized document model; no workbook/XML handles cross this boundary."""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

from .diagnostics import BuildError, Diagnostic


@dataclass(frozen=True)
class Source:
    path: Path = Path("document.yaml")
    pointer: str = ""
    line: int | None = None
    column: int | None = None

    reference_chain: tuple[dict, ...] = ()

    def fail(self, code: str, message: str):
        raise BuildError(
            code,
            message,
            self.path,
            self.pointer,
            line=self.line,
            column=self.column,
            details={"reference_chain": list(self.reference_chain)} if self.reference_chain else None,
        )


@dataclass(frozen=True)
class TextStyle:
    font_family: str = "함초롬바탕"
    font_size: Decimal = Decimal(10)
    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    color: str = "#000000"
    highlight: str | None = None


@dataclass(frozen=True)
class Run:
    text: str
    style: TextStyle
    linebreak: bool = False
    start: int = 0
    end: int = 0
    reference: str | None = None


@dataclass(frozen=True)
class ParagraphStyle:
    align: str = "left"
    indent: Decimal = Decimal(0)
    right_indent: Decimal = Decimal(0)
    first_line_indent: Decimal = Decimal(0)
    space_before: Decimal = Decimal(0)
    space_after: Decimal = Decimal(0)
    line_spacing: int = 160
    keep_with_next: bool = False
    keep_together: bool = False


@dataclass(frozen=True)
class SectionStyle:
    indent_by_depth: tuple[Decimal, ...] = (Decimal(0),)


@dataclass(frozen=True)
class Paragraph:
    runs: tuple[Run, ...] = ()
    style: ParagraphStyle = field(default_factory=ParagraphStyle)
    source: Source = field(default_factory=Source)
    indent_depth: int = 0


@dataclass(frozen=True)
class Picture:
    data: bytes
    format: str
    width: Decimal
    height: Decimal
    align: str = "left"
    source: Source = field(default_factory=Source)


@dataclass(frozen=True)
class Border:
    style: str = "NONE"
    width: str = "0.12 mm"
    color: str = "#000000"


@dataclass(frozen=True)
class Cell:
    row: int
    col: int
    row_span: int
    col_span: int
    paragraph: Paragraph
    vertical_align: str = "top"
    background: str | None = None
    borders: tuple[Border, ...] = (Border(),) * 4  # left, right, top, bottom
    picture: Picture | None = None
    picture_as_background: bool = False
    padding: tuple[Decimal, ...] = (Decimal(0),) * 4
    paragraphs: tuple[Paragraph, ...] = ()
    size: tuple[Decimal, Decimal] | None = None
    header: bool = False


@dataclass(frozen=True)
class TableData:
    widths: tuple[Decimal, ...]
    heights: tuple[Decimal, ...]
    cells: tuple[Cell, ...]
    source: Source
    host_style: ParagraphStyle = field(default_factory=lambda: ParagraphStyle(line_spacing=100))
    page_break: str = "CELL"
    repeat_header: bool = False
    inline: bool = True
    padding: tuple[Decimal, ...] = (Decimal(0),) * 4
    margin: tuple[Decimal, ...] = (Decimal(0),) * 4
    size: tuple[Decimal, Decimal] | None = None


@dataclass(frozen=True)
class Table:
    data: TableData
    source: Source


@dataclass(frozen=True)
class PageBreak:
    source: Source = field(default_factory=Source)


@dataclass(frozen=True)
class TextBox:
    paragraphs: tuple[Paragraph, ...]
    x: Decimal
    y: Decimal
    width: Decimal
    height: Decimal
    padding: tuple[Decimal, ...] = (Decimal(0),) * 4


@dataclass(frozen=True)
class Composition:
    contents: tuple[Picture | TextBox, ...]
    width: Decimal
    height: Decimal
    source: Source


Block = Paragraph | Picture | Table | PageBreak | Composition


@dataclass(frozen=True)
class Page:
    width: Decimal = Decimal(210)
    height: Decimal = Decimal(297)
    left: Decimal = Decimal(30)
    right: Decimal = Decimal(30)
    top: Decimal = Decimal(20)
    bottom: Decimal = Decimal(15)
    header: Decimal = Decimal(15)
    footer: Decimal = Decimal(15)
    gutter: Decimal = Decimal(0)
    gutter_position: str = "left"

    @property
    def body_width(self) -> Decimal:
        return self.width - self.left - self.right - (self.gutter if self.gutter_position != "top" else 0)


@dataclass(frozen=True)
class Section:
    contents: tuple[Block, ...]
    page: Page
    source: Source
    style: SectionStyle = field(default_factory=SectionStyle)


@dataclass(frozen=True)
class Document:
    root: Path
    sections: tuple[Section, ...]
    variables: tuple[tuple[str, str], ...] = ()
    diagnostics: tuple[Diagnostic, ...] = ()
    inputs: tuple[Path, ...] = ()
