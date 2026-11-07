"""Read-only resource normalization. No formula evaluation or report calculations."""

from __future__ import annotations

import re
from dataclasses import replace
from datetime import date, datetime, time
from decimal import ROUND_HALF_UP, Decimal
from io import BytesIO
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from defusedxml import ElementTree as ET
from openpyxl import load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import get_column_letter, range_boundaries
from PIL import Image, UnidentifiedImageError

from .diagnostics import BuildError, Diagnostic
from .model import Border, Cell, Paragraph, ParagraphStyle, Picture, Source, TableData, TextStyle
from .presentation import border_signature, decode_style, mm, read_note, split_paragraphs
from .text import literal_runs

D = Decimal


def _display(value, fmt: str, source: Source) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (CellRichText, str)):
        return str(value)
    if isinstance(value, (datetime, date, time)):
        patterns = {
            "yyyy-mm-dd": "%Y-%m-%d",
            "yyyy.mm.dd": "%Y.%m.%d",
            "yyyy.mm.dd.": "%Y.%m.%d.",
            "mm/dd/yyyy": "%m/%d/%Y",
            "hh:mm": "%H:%M",
            "h:mm": "%H:%M",
            "yyyy-mm-dd h:mm:ss": "%Y-%m-%d %H:%M:%S",
        }
        if fmt.lower() in patterns:
            return value.strftime(patterns[fmt.lower()])
        source.fail("UNSUPPORTED_XLSX_FORMAT", f"지원하지 않는 날짜 표시 형식: {fmt}")
    if isinstance(value, (int, float, Decimal)):
        if fmt in ("General", "@"):
            return str(value)
        if re.fullmatch(r"(#,##0|0)(\.0+)?%?", fmt):
            percent = fmt.endswith("%")
            clean = fmt.rstrip("%")
            places = len(clean.split(".")[1]) if "." in clean else 0
            number = D(str(value)) * (100 if percent else 1)
            number = number.quantize(D(1).scaleb(-places), rounding=ROUND_HALF_UP)
            return format(number, f"{',' if ',' in clean else ''}.{places}f") + ("%" if percent else "")
        source.fail("UNSUPPORTED_XLSX_FORMAT", f"지원하지 않는 숫자 표시 형식: {fmt}")
    source.fail("INVALID_RESOURCE_VALUE", f"셀 값의 타입을 지원하지 않습니다: {type(value).__name__}")


class WorkbookReader:
    def __init__(self, path: Path, definition: dict, source: Source):
        self.path, self.definition, self.source = path, definition, source
        self.book = None
        self.values = None
        self.diagnostics = []
        self.geometry_exact = False

    def __enter__(self):
        try:
            self.book = load_workbook(self.path, data_only=False, rich_text=True)
            self.values = load_workbook(self.path, data_only=True, rich_text=True)
            sheet_name = self.definition.get("sheet")
            if sheet_name is None:
                visible = [s.title for s in self.book if s.sheet_state == "visible"]
                if not visible:
                    self.source.fail("INVALID_SHEET", "표시 가능한 시트가 없습니다.")
                sheet_name = visible[0]
            if sheet_name not in self.book:
                self.source.fail("INVALID_SHEET", f"시트가 없습니다: {sheet_name}")
            self.sheet = self.book[sheet_name]
            self.cached = self.values[sheet_name]
            if "range" in self.definition:
                value = self.definition["range"]
                if not isinstance(value, str) or not re.fullmatch(
                    r"\$?[A-Z]+\$?[1-9]\d*:\$?[A-Z]+\$?[1-9]\d*", value
                ):
                    self.source.fail("INVALID_RANGE", f"범위는 A1:B2 형식이어야 합니다: {value}")
                self.bounds = range_boundaries(value)
            else:
                used = [c for c in self.sheet._cells.values() if c.value is not None or c.has_style]
                self.bounds = (
                    min((c.column for c in used), default=1),
                    min((c.row for c in used), default=1),
                    max((c.column for c in used), default=1),
                    max((c.row for c in used), default=1),
                )
            c1, r1, c2, r2 = self.bounds
            if c1 > c2 or r1 > r2 or (c2 - c1 + 1) * (r2 - r1 + 1) > 200000:
                self.source.fail("INVALID_RANGE", "범위가 잘못되었거나 200000셀 제한을 넘었습니다.")
            for merge in self.sheet.merged_cells.ranges:
                intersects = not (
                    merge.max_col < c1 or merge.min_col > c2 or merge.max_row < r1 or merge.min_row > r2
                )
                if intersects and not (
                    c1 <= merge.min_col <= merge.max_col <= c2 and r1 <= merge.min_row <= merge.max_row <= r2
                ):
                    self.source.fail("INVALID_CELL_SPAN", f"선택 범위에 걸친 병합: {merge}")
            return self
        except BuildError:
            self.__exit__(None, None, None)
            raise
        except (OSError, ValueError, KeyError, BadZipFile) as exc:
            self.__exit__(None, None, None)
            self.source.fail("RESOURCE_READ_FAILED", str(exc))

    def __exit__(self, *args):
        if self.book is not None:
            self.book.close()
        if self.values is not None:
            self.values.close()

    def cell_value(self, row, col):
        cell = self.sheet.cell(row, col)
        value = cell.value
        if cell.data_type == "f":
            value = self.cached.cell(row, col).value
            if value is None:
                self.cell_source(row, col).fail(
                    "FORMULA_CACHE_MISSING",
                    "수식은 계산하지 않습니다. Excel에서 계산 후 저장하거나 값으로 붙여넣으세요.",
                )
        if cell.data_type == "e":
            self.cell_source(row, col).fail("XLSX_CELL_ERROR", f"Excel 오류 값: {value}")
        return value

    def cell_source(self, row, col):
        return Source(self.path, f"/{self.sheet.title}!{get_column_letter(col)}{row}")

    def theme_colors(self):
        if not self.book.loaded_theme:
            return []
        root = ET.fromstring(self.book.loaded_theme)
        scheme = root.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme")
        if scheme is None:
            return []
        colors = [(next(iter(n)).get("lastClr") or next(iter(n)).get("val")) for n in scheme]
        # Spreadsheet theme indexes are lt1, dk1, lt2, dk2, then accents/hyperlinks.
        return [colors[i] for i in (1, 0, 3, 2, *range(4, len(colors)))]

    def rgb(self, c, default="#000000"):
        if c is None:
            return default
        if c.type == "rgb":
            base = c.rgb[-6:]
        elif c.type == "indexed":
            if c.indexed >= 64:
                return default
            base = COLOR_INDEX[c.indexed][-6:]
        elif c.type == "theme":
            palette = self.theme_colors()
            if c.theme >= len(palette):
                self.source.fail("UNSUPPORTED_XLSX_COLOR", "테마 색상을 해석할 수 없습니다.")
            base = palette[c.theme][-6:]
        elif c.type == "auto":
            return default
        else:
            self.source.fail("UNSUPPORTED_XLSX_COLOR", f"색상 타입을 지원하지 않습니다: {c.type}")
        if c.tint:
            self.source.fail("UNSUPPORTED_XLSX_COLOR", "테마 tint 색상은 아직 지원하지 않습니다.")
        return "#" + base.upper()

    def font(self, font, source):
        if font.u not in (None, "single") or font.vertAlign:
            source.fail(
                "UNSUPPORTED_XLSX_FORMAT", "이중 밑줄 또는 위/아래 첨자 글꼴은 아직 지원하지 않습니다."
            )
        name = getattr(font, "rFont", None) or getattr(font, "name", None)
        if not isinstance(name, str) or not name:
            name = "Calibri"
        return TextStyle(
            font_family=name,
            font_size=D(str(font.sz or 11)),
            bold=bool(font.b),
            italic=bool(font.i),
            underline=bool(font.u),
            strike=bool(font.strike),
            color=self.rgb(font.color),
        )

    def paragraph(self, row, col):
        cell = self.sheet.cell(row, col)
        source = self.cell_source(row, col)
        value = self.cell_value(row, col)
        base = self.font(cell.font, source)
        if isinstance(value, CellRichText):
            runs = []
            for part in value:
                style = self.font(part.font, source) if isinstance(part, TextBlock) else base
                runs.extend(literal_runs(part.text if isinstance(part, TextBlock) else part, style))
        else:
            runs = literal_runs(_display(value, cell.number_format, source), base)
        alignment = cell.alignment
        if alignment.textRotation or alignment.indent or alignment.shrinkToFit:
            source.fail("UNSUPPORTED_XLSX_FORMAT", "회전/들여쓰기/축소 맞춤 셀은 아직 지원하지 않습니다.")
        align = alignment.horizontal or ("right" if isinstance(value, (int, float)) else "left")
        if align == "general":
            align = "left"
        if align == "distributed":
            align = "distribute"
        if align not in ("left", "right", "center", "justify", "distribute"):
            source.fail("UNSUPPORTED_XLSX_FORMAT", f"지원하지 않는 정렬: {align}")
        return Paragraph(tuple(runs), ParagraphStyle(align=align, line_spacing=100), source)

    def border(self, side):
        if side is None or not side.style:
            return Border()
        kinds = {
            "thin": ("SOLID", "0.12 mm"),
            "hair": ("SOLID", "0.1 mm"),
            "medium": ("SOLID", "0.3 mm"),
            "thick": ("SOLID", "0.5 mm"),
            "dashed": ("DASH", "0.12 mm"),
            "dotted": ("DOT", "0.12 mm"),
            "double": ("DOUBLE_SLIM", "0.3 mm"),
            "dashDot": ("DASH_DOT", "0.12 mm"),
            "mediumDashed": ("DASH", "0.3 mm"),
            "mediumDashDot": ("DASH_DOT", "0.3 mm"),
        }
        if side.style not in kinds:
            self.source.fail("UNSUPPORTED_XLSX_FORMAT", f"테두리 유형: {side.style}")
        kind, width = kinds[side.style]
        return Border(kind, width, self.rgb(side.color))

    def hint(self, cell):
        try:
            return read_note(cell)
        except (ValueError, TypeError) as exc:
            self.cell_source(cell.row, cell.column).fail("INVALID_XLSX_PRESENTATION", str(exc))

    def lengths(self, values, count):
        try:
            result = tuple(D(v) for v in values)
            if len(result) != count or any(not v.is_finite() or v < 0 for v in result):
                raise ValueError("Invalid presentation dimensions")
            return result
        except (ValueError, TypeError, ArithmeticError) as exc:
            self.source.fail("INVALID_XLSX_PRESENTATION", str(exc))

    def table(self) -> TableData:
        c1, r1, c2, r2 = self.bounds
        s = self.sheet
        # openpyxl may discard unsupported shapes with a Python warning. Inspect the
        # selected sheet's original drawing parts so a text box never vanishes silently.
        with ZipFile(self.path) as archive:
            for relationship in s._rels:
                if not relationship.Type.endswith("/drawing"):
                    continue
                if relationship.TargetMode == "External":
                    self.source.fail("UNSUPPORTED_XLSX_DRAWING", "외부 Drawing 참조는 지원하지 않습니다.")
                drawing = ET.fromstring(archive.read(relationship.Target.lstrip("/")))
                if any(
                    n.tag.rsplit("}", 1)[-1] in ("sp", "grpSp", "cxnSp", "graphicFrame")
                    for n in drawing.iter()
                ):
                    self.source.fail(
                        "UNSUPPORTED_XLSX_DRAWING",
                        "텍스트박스/도형/차트를 조용히 제거할 수 없습니다. 지원하지 않는 Drawing입니다.",
                    )
        if s.conditional_formatting or s._charts:
            self.source.fail("UNSUPPORTED_XLSX_FORMAT", "조건부 서식/차트는 아직 지원하지 않습니다.")
        if any(d.hidden for d in s.column_dimensions.values()) or any(
            d.hidden for d in s.row_dimensions.values()
        ):
            self.source.fail("UNSUPPORTED_XLSX_FORMAT", "숨김 행/열은 아직 지원하지 않습니다.")

        def column_width(col):
            dim = next((d for d in s.column_dimensions.values() if (d.min or 0) <= col <= (d.max or 0)), None)
            chars = D(str(dim.width if dim is not None else s.sheet_format.defaultColWidth or 8.43))
            pixels = chars * 7 + 5 if chars >= 1 else chars * 12
            return pixels * D("25.4") / 96

        widths = tuple(column_width(c) for c in range(c1, c2 + 1))
        heights = tuple(
            D(str(s.row_dimensions[r].height or s.sheet_format.defaultRowHeight or 15)) * D("25.4") / 72
            for r in range(r1, r2 + 1)
        )
        table_hint = self.hint(s.cell(r1, c1)).get("table", {})
        if not isinstance(table_hint, dict):
            self.source.fail("INVALID_XLSX_PRESENTATION", "table 서식 정보는 객체여야 합니다.")
        if not (c1 == r1 == 1 and table_hint.get("rows") == r2 and table_hint.get("columns") == c2):
            table_hint = {}
        geometry_unchanged = False
        if table_hint:
            try:
                if len(table_hint["excel_widths"]) != c2 or len(table_hint["excel_heights"]) != r2:
                    raise ValueError("Invalid geometry snapshot")
                geometry_unchanged = all(
                    abs(D(str(s.column_dimensions[get_column_letter(i)].width)) - D(str(v))) < D("1e-10")
                    for i, v in enumerate(table_hint["excel_widths"], 1)
                ) and all(
                    abs(D(str(s.row_dimensions[i].height)) - D(str(v))) < D("1e-10")
                    for i, v in enumerate(table_hint["excel_heights"], 1)
                )
                if geometry_unchanged:
                    widths = tuple(mm(v) for v in self.lengths(table_hint["widths"], c2))
                    heights = tuple(mm(v) for v in self.lengths(table_hint["heights"], r2))
                    if any(v <= 0 for v in (*widths, *heights)):
                        raise ValueError("Table dimensions must be positive")
            except (KeyError, ValueError, TypeError, ArithmeticError) as exc:
                self.source.fail("INVALID_XLSX_PRESENTATION", str(exc))
        self.geometry_exact = geometry_unchanged
        anchors = {}
        covered = set()
        for merge in s.merged_cells.ranges:
            if c1 <= merge.min_col <= merge.max_col <= c2 and r1 <= merge.min_row <= merge.max_row <= r2:
                anchors[(merge.min_row, merge.min_col)] = (
                    merge.max_row - merge.min_row + 1,
                    merge.max_col - merge.min_col + 1,
                )
                covered.update(
                    (r, c)
                    for r in range(merge.min_row, merge.max_row + 1)
                    for c in range(merge.min_col, merge.max_col + 1)
                    if (r, c) != (merge.min_row, merge.min_col)
                )
        cells = []
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                if (r, c) in covered:
                    continue
                cell = s.cell(r, c)
                rs, cs = anchors.get((r, c), (1, 1))
                if cell.fill.patternType not in (None, "solid"):
                    self.cell_source(r, c).fail(
                        "UNSUPPORTED_XLSX_FORMAT", "패턴/그라데이션 채우기는 지원하지 않습니다."
                    )
                if cell.border.diagonalUp or cell.border.diagonalDown:
                    self.cell_source(r, c).fail(
                        "UNSUPPORTED_XLSX_FORMAT", "대각선 테두리는 지원하지 않습니다."
                    )
                vertical = cell.alignment.vertical or "top"
                if vertical not in ("top", "center", "bottom"):
                    self.cell_source(r, c).fail("UNSUPPORTED_XLSX_FORMAT", f"세로 정렬: {vertical}")
                cells.append(
                    Cell(
                        r - r1,
                        c - c1,
                        rs,
                        cs,
                        self.paragraph(r, c),
                        vertical,
                        self.rgb(cell.fill.fgColor, "#FFFFFF") if cell.fill.patternType == "solid" else None,
                        tuple(
                            self.border(getattr(cell.border, k)) for k in ("left", "right", "top", "bottom")
                        ),
                    )
                )
                hint = self.hint(cell)
                if hint:
                    current = cells[-1]
                    try:
                        paragraphs = split_paragraphs(current.paragraph, hint)
                        if len(hint.get("paragraphs", [])) > 1 and not paragraphs:
                            origin = self.cell_source(r, c)
                            self.diagnostics.append(
                                Diagnostic(
                                    "XLSX_PARAGRAPH_BOUNDARIES_CHANGED",
                                    "셀 내용이 변경되어 원본 문단 경계를 복원하지 않고 현재 Excel 줄바꿈을 사용합니다.",
                                    str(origin.path),
                                    origin.pointer,
                                    "warning",
                                )
                            )
                        base = current.paragraph
                        if hint.get("paragraphs"):
                            base = replace(
                                base,
                                style=replace(
                                    decode_style(hint["paragraphs"][0]["style"]), align=base.style.align
                                ),
                            )
                        size = (
                            self.lengths(hint["size"], 2)
                            if geometry_unchanged and hint.get("span") == [rs, cs]
                            else None
                        )
                        borders = list(current.borders)
                        if "borders" in hint:
                            if len(hint["borders"]) != 4 or len(hint["excel_borders"]) != 4:
                                raise ValueError("Invalid border count")
                            for i, side in enumerate(("left", "right", "top", "bottom")):
                                if border_signature(getattr(cell.border, side)) == hint["excel_borders"][i]:
                                    b = hint["borders"][i]
                                    if (
                                        b["type"]
                                        not in (
                                            "NONE",
                                            "SOLID",
                                            "DASH",
                                            "DOT",
                                            "DASH_DOT",
                                            "DASH_DOT_DOT",
                                            "LONG_DASH",
                                            "DOUBLE_SLIM",
                                            "SLIM_THICK",
                                            "THICK_SLIM",
                                            "SLIM_THICK_SLIM",
                                            "CIRCLE",
                                        )
                                        or not re.fullmatch(r"\d+(?:\.\d+)? mm", b["width"])
                                        or not re.fullmatch(r"#[0-9A-Fa-f]{6}", b["color"])
                                    ):
                                        raise ValueError("Invalid border presentation")
                                    borders[i] = Border(b["type"], b["width"], b["color"])
                        if type(hint.get("header", False)) is not bool:
                            raise ValueError("Invalid header flag")
                        cells[-1] = replace(
                            current,
                            paragraph=base,
                            paragraphs=paragraphs,
                            size=size,
                            padding=self.lengths(hint.get("padding", [0] * 4), 4),
                            borders=tuple(borders),
                            header=hint.get("header", False),
                        )
                    except (KeyError, ValueError, TypeError, ArithmeticError) as exc:
                        self.cell_source(r, c).fail("INVALID_XLSX_PRESENTATION", str(exc))
        for image in s._images:
            anchor = image.anchor
            if not hasattr(anchor, "_from") or not hasattr(anchor, "ext"):
                self.source.fail("UNSUPPORTED_XLSX_DRAWING", "셀 하나에 고정된 그림만 지원합니다.")
            r, c = anchor._from.row - r1 + 1, anchor._from.col - c1 + 1
            if r < 0 or c < 0 or r > r2 - r1 or c > c2 - c1:
                continue
            matches = [i for i, cell in enumerate(cells) if cell.row == r and cell.col == c]
            if len(matches) != 1:
                self.source.fail("UNSUPPORTED_XLSX_DRAWING", "그림은 병합 셀의 시작 셀에 고정해야 합니다.")
            index = matches[0]
            cell = cells[index]
            if cell.picture is not None:
                self.source.fail("UNSUPPORTED_XLSX_DRAWING", "셀당 그림 하나만 지원합니다.")
            pw = D(anchor.ext.cx) * D("25.4") / 914400
            ph = D(anchor.ext.cy) * D("25.4") / 914400
            cw = sum(widths[c : c + cell.col_span])
            ch = sum(heights[r : r + cell.row_span])
            x = D(anchor._from.colOff) * D("25.4") / 914400
            y = D(anchor._from.rowOff) * D("25.4") / 914400
            background = abs(pw - cw) < D("0.1") and abs(ph - ch) < D("0.1") and x == 0 and y == 0
            if not background and (abs(x - (cw - pw) / 2) > D("0.2") or abs(y - (ch - ph) / 2) > D("0.2")):
                self.source.fail("UNSUPPORTED_XLSX_DRAWING", "셀 전체 배경 또는 셀 가운데 그림만 지원합니다.")
            picture = Picture(image._data(), image.format, pw, ph, "center", cell.paragraph.source)
            cells[index] = replace(cell, picture=picture, picture_as_background=background)
        data = TableData(widths, heights, tuple(cells), self.source)
        if table_hint:
            try:
                if table_hint["page_break"] not in ("CELL", "TABLE", "NONE") or any(
                    type(table_hint[k]) is not bool for k in ("repeat_header", "inline")
                ):
                    raise ValueError("Invalid table presentation")
                data = replace(
                    data,
                    host_style=decode_style(table_hint["host_style"]),
                    page_break=table_hint["page_break"],
                    repeat_header=table_hint["repeat_header"],
                    inline=table_hint["inline"],
                    padding=self.lengths(table_hint["padding"], 4),
                    margin=self.lengths(table_hint["margin"], 4),
                    size=self.lengths(table_hint["size"], 2) if geometry_unchanged else None,
                )
            except (KeyError, ValueError, TypeError, ArithmeticError) as exc:
                self.source.fail("INVALID_XLSX_PRESENTATION", str(exc))
        return data


class Resources:
    def __init__(self, root: Path, tables: dict, images: dict):
        self.root, self.tables, self.images = root, tables, images
        self.inputs = set()
        self.diagnostics = []
        self.cache = {}

    def path(self, definition: dict, source: Source) -> Path:
        raw = definition.get("path")
        if not isinstance(raw, str) or not raw:
            source.fail("MISSING_REQUIRED", "리소스 path 문자열이 필요합니다.")
        path = Path(raw)
        resolved = (path if path.is_absolute() else self.root / path).resolve()
        if not path.is_absolute() and not resolved.is_relative_to(self.root):
            source.fail(
                "RESOURCE_OUTSIDE_ROOT",
                "상대 리소스 경로는 기준 디렉터리를 벗어날 수 없습니다. 외부 파일은 절대 경로로 지정하세요.",
            )
        if not resolved.is_file():
            source.fail("RESOURCE_NOT_FOUND", f"리소스 파일이 없습니다: {resolved}")
        self.inputs.add(resolved)
        return resolved

    def workbook(self, name, source):
        if name not in self.tables:
            source.fail("RESOURCE_NOT_FOUND", f"표 리소스가 없습니다: {name}")
        definition, origin = self.tables[name]
        path = self.path(definition, origin)
        if path.suffix.lower() != ".xlsx":
            source.fail("UNSUPPORTED_RESOURCE", "완성 표는 XLSX만 지원합니다.")
        return WorkbookReader(path, definition, origin)

    def variables(self, name, source):
        with self.workbook(name, source) as reader:
            c1, r1, c2, r2 = reader.bounds
            if c2 - c1 != 1:
                source.fail("VARIABLE_TABLE_COLUMNS", "변수 표는 정확히 2열이어야 합니다.")
            if reader.sheet.merged_cells.ranges:
                source.fail("VARIABLE_TABLE_MERGED", "변수 표의 병합 셀은 허용하지 않습니다.")
            start = r1 + 1 if reader.definition.get("header", False) else r1
            values = {}
            for r in range(start, r2 + 1):
                key = reader.cell_value(r, c1)
                value = reader.cell_value(r, c2)
                if key is None and value is None:
                    continue
                if not isinstance(key, str) or not key.strip():
                    reader.cell_source(r, c1).fail("EMPTY_VARIABLE", "빈 변수명은 허용하지 않습니다.")
                if key in values:
                    reader.cell_source(r, c1).fail("DUPLICATE_VARIABLE", f"중복 변수: {key}")
                values[key] = _display(
                    value, reader.sheet.cell(r, c2).number_format, reader.cell_source(r, c2)
                )
            self.variable_path = reader.path
            return values

    def table(self, name, source):
        if name in self.cache:
            return self.cache[name]
        with self.workbook(name, source) as reader:
            if getattr(self, "variable_path", None) == reader.path:
                source.fail("VARIABLE_TABLE_NOT_INDEPENDENT", "문서 변수와 완성 표는 별도 파일이어야 합니다.")
            data = reader.table()
            self.diagnostics.extend(reader.diagnostics)
            if not reader.geometry_exact:
                self.diagnostics.append(
                    Diagnostic(
                        "XLSX_COLUMN_WIDTH_APPROXIMATION",
                        "Excel 열 너비는 기본 글꼴 기준 픽셀→mm 근사입니다. 추출본의 서식 메모가 있으면 원본 치수·셀 여백을 복원합니다.",
                        str(reader.path),
                        source.pointer,
                        "warning",
                    )
                )
            self.cache[name] = data
            return data

    def picture(self, name, style, source):
        if name not in self.images:
            source.fail("RESOURCE_NOT_FOUND", f"이미지 리소스가 없습니다: {name}")
        definition, origin = self.images[name]
        path = self.path(definition, origin)
        try:
            raw = path.read_bytes()
            with Image.open(BytesIO(raw)) as image:
                w, h = image.size
                dpi = image.info.get("dpi", (96, 96))
                fmt = image.format.lower()
        except (OSError, UnidentifiedImageError) as exc:
            source.fail("INVALID_IMAGE", str(exc))
        width = style.get("width")
        height = style.get("height")
        native_w = D(w) * D("25.4") / D(str(dpi[0] or 96))
        native_h = D(h) * D("25.4") / D(str(dpi[1] or 96))
        if width is None and height is None:
            width, height = native_w, native_h
        elif width is None:
            width = height * D(w) / h
        elif height is None:
            height = width * D(h) / w
        fit = style.get("fit", "contain")
        if fit == "contain":
            scale = min(width / D(w), height / D(h))
            width, height = D(w) * scale, D(h) * scale
        elif fit == "none":
            width, height = native_w, native_h
        elif fit != "fill":
            source.fail("UNSUPPORTED_RENDERING", "cover 잘라내기는 아직 지원하지 않습니다.")
        if fmt not in ("png", "jpeg", "bmp"):
            source.fail("UNSUPPORTED_RESOURCE", f"이미지 형식을 지원하지 않습니다: {fmt}")
        return Picture(raw, fmt, width, height, style.get("align", "left"), source)
