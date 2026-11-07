"""One-time migration of finished HWPX tables to editable Excel sheets.

This module knows HWPX, not report-domain models or calculations.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import tempfile
from contextlib import ExitStack
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from defusedxml import ElementTree as ET
from openpyxl import Workbook, load_workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PillowImage

from hwpx_builder.diagnostics import BuildError
from hwpx_builder.presentation import border_signature, digest, mm

from .reference_styles import margins, paragraph_style, write_note
from .repair_xlsx import repair_workbook

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"


def _text(t) -> str:
    if t is None:
        return ""
    result = t.text or ""
    for child in t:
        result += "\n" if child.tag == HP + "lineBreak" else "".join(child.itertext())
        result += child.tail or ""
    return result


def _rgb(value, default="000000"):
    if not value or value.lower() in ("none", "auto"):
        return default
    return value.lstrip("#")[-6:]


def _font(node, fonts):
    ref = node.find(HH + "fontRef")
    under = node.find(HH + "underline")
    strike = node.find(HH + "strikeout")
    return Font(
        name=fonts.get(ref.get("hangul") if ref is not None else "0", "함초롬바탕"),
        size=int(node.get("height", "1000")) / 100,
        bold=node.find(HH + "bold") is not None,
        italic=node.find(HH + "italic") is not None,
        underline="single" if under is not None and under.get("type") != "NONE" else None,
        strike=strike is not None and strike.get("shape", "NONE") != "NONE",
        color=_rgb(node.get("textColor")),
    )


def _side(node):
    if node is None or node.get("type") == "NONE":
        return Side(style=None)
    kind = node.get("type", "SOLID")
    width = float(node.get("width", "0.12 mm").split()[0])
    style = {
        "DASH": "dashed",
        "DOT": "dotted",
        "DOUBLE_SLIM": "double",
        "DOUBLE": "double",
        "DASH_DOT": "dashDot",
    }.get(kind)
    if style is None:
        style = "thin" if width <= 0.2 else "medium" if width <= 0.5 else "thick"
    return Side(style=style, color=_rgb(node.get("color")))


def _dimensions(count, constraints):
    """Solve cumulative grid boundaries, including columns with no unmerged cells."""
    edges = {i: [] for i in range(count + 1)}
    for start, span, length in constraints:
        edges[start].append((start + span, length))
        edges[start + span].append((start, -length))
    positions = {0: 0.0}
    queue = [0]
    for start in queue:
        for end, distance in edges[start]:
            if end not in positions:
                positions[end] = positions[start] + distance
                queue.append(end)
    # Underconstrained boundaries are interpolated between known enclosing edges.
    positions.setdefault(count, max(positions.values(), default=0) or 2000 * count)
    for i in range(1, count):
        if i not in positions:
            left = max(j for j in positions if j < i)
            right = min(j for j in positions if j > i)
            positions[i] = positions[left] + (positions[right] - positions[left]) * (i - left) / (
                right - left
            )
    return [max(1.0, positions[i + 1] - positions[i]) for i in range(count)]


def _add_image(sheet, archive, binaries, image_node, row, col, width, height, x=0, y=0):
    reference = image_node.get("binaryItemIDRef")
    if reference not in binaries:
        raise BuildError("RESOURCE_NOT_FOUND", f"이미지 참조를 찾을 수 없습니다: {reference}")
    # Convert BMP to PNG in memory; Excel drawings cannot reliably embed BMP.
    data = BytesIO()
    with PillowImage.open(BytesIO(archive.read(binaries[reference]))) as picture:
        picture.save(data, format="PNG")
    data.seek(0)
    image = Image(data)
    # 1 HWPUNIT = 127 EMU (914400 EMU / 7200).
    image.anchor = OneCellAnchor(
        _from=AnchorMarker(col=col, row=row, colOff=round(x * 127), rowOff=round(y * 127)),
        ext=XDRPositiveSize2D(cx=round(width * 127), cy=round(height * 127)),
    )
    sheet.add_image(image)
    return {
        "cell": f"{get_column_letter(col + 1)}{row + 1}",
        "source": binaries[reference],
        "width_hwpunit": width,
        "height_hwpunit": height,
    }


def extract_tables(
    source: Path | str,
    output: Path | str,
    *,
    overwrite: bool = False,
    variables_profile: Path | str | None = None,
    variables_output: Path | str | None = None,
) -> dict:
    """Export every hp:tbl in document order; never evaluate or infer formulas."""
    source, output = Path(source).resolve(), Path(output).resolve()
    manifest_path = output.with_suffix(".manifest.json")
    variables_output = Path(variables_output).resolve() if variables_output is not None else None
    if (variables_profile is None) != (variables_output is None):
        raise BuildError("INVALID_PROFILE", "변수 프로필과 별도 변수 XLSX 경로를 함께 지정하세요.")
    outputs = [output, manifest_path] + ([variables_output] if variables_output is not None else [])
    if len(set(outputs + [source])) != len(outputs) + 1:
        raise BuildError("OUTPUT_CONFLICT", "원본/표/변수/manifest 경로는 서로 달라야 합니다.")
    if not overwrite and any(p.exists() for p in outputs):
        raise BuildError("OUTPUT_EXISTS", "출력이 이미 있습니다. --overwrite로 명시하세요.", output)
    with ZipFile(source) as archive:
        package = ET.fromstring(archive.read("Contents/content.hpf"))
        binaries = {
            n.get("id"): n.get("href")
            for n in package.iter()
            if n.tag.endswith("}item") and n.get("media-type", "").startswith("image/")
        }
        header = ET.fromstring(archive.read("Contents/header.xml"))
        fonts_node = next(n for n in header.iter(HH + "fontface") if n.get("lang") == "HANGUL")
        fonts = {n.get("id"): n.get("face") for n in fonts_node}
        chars = {n.get("id"): _font(n, fonts) for n in header.iter(HH + "charPr")}
        paras = {n.get("id"): n for n in header.iter(HH + "paraPr")}
        fills = {n.get("id"): n for n in header.iter(HH + "borderFill")}
        sections = sorted(
            (n for n in archive.namelist() if n.startswith("Contents/section") and n.endswith(".xml")),
            key=lambda n: int(n.split("section")[-1].split(".")[0]),
        )
        book = Workbook()
        book.remove(book.active)
        manifest = {
            "source": source.name,
            "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "tables": [],
            "notes": [
                "표시된 문자열만 추출했습니다. 수식·계산·데이터 형식 추론을 하지 않습니다.",
                "열 너비는 Excel의 기본 글꼴 기준 픽셀 근사값입니다. 행 높이는 pt로 환산했습니다.",
                "Excel에 없는 셀 여백·문단 간격·표 배치는 XLSX 셀 메모의 presentation:1 서식 정보로 보관합니다.",
            ],
            "variables": [],
            "source_paragraphs": [],
        }
        for section in sections:
            root = ET.fromstring(archive.read(section))

            # Full text outside tables is reusable material, not inferred business data.
            # Include paragraphs inside text boxes while excluding table-cell paragraphs.
            def outside_table_text(node):
                if node.tag == HP + "tbl":
                    return
                if node.tag == HP + "p":
                    value = "".join(
                        _text(t) for run in node.findall(HP + "run") for t in run.findall(HP + "t")
                    )
                    if value.strip():
                        yield node, value
                for child in node:
                    yield from outside_table_text(child)

            for paragraph, value in outside_table_text(root):
                manifest["source_paragraphs"].append(
                    {"text": value, "section": section, "paragraph_id": paragraph.get("id")}
                )
            for table in root.iter(HP + "tbl"):
                number = len(manifest["tables"]) + 1
                sheet = book.create_sheet(f"table_{number:02d}")
                nr, nc = int(table.get("rowCnt")), int(table.get("colCnt"))
                constraints_x, constraints_y, cells_meta, merges = [], [], [], []
                warnings, images_meta = [], []
                hints = {}
                for row in table.findall(HP + "tr"):
                    for cell in row.findall(HP + "tc"):
                        address, span, size = (
                            cell.find(HP + name) for name in ("cellAddr", "cellSpan", "cellSz")
                        )
                        r, c = int(address.get("rowAddr")), int(address.get("colAddr"))
                        rs, cs = int(span.get("rowSpan")), int(span.get("colSpan"))
                        width, height = int(size.get("width")), int(size.get("height"))
                        constraints_x.append((c, cs, width))
                        constraints_y.append((r, rs, height))
                        target = sheet.cell(r + 1, c + 1)
                        sub = cell.find(HP + "subList")
                        rich, plain = CellRichText(), []
                        for pi, para in enumerate(sub.findall(HP + "p")):
                            if pi:
                                rich.append("\n")
                                plain.append("\n")
                            for run in para.findall(HP + "run"):
                                value = "".join(_text(t) for t in run.findall(HP + "t"))
                                style = chars.get(run.get("charPrIDRef"), Font(name="함초롬바탕", size=10))
                                if value:
                                    rich.append(
                                        TextBlock(
                                            InlineFont(
                                                rFont=style.name,
                                                sz=style.sz,
                                                b=style.b,
                                                i=style.i,
                                                u=style.u,
                                                strike=style.strike,
                                                color=style.color,
                                            ),
                                            value,
                                        )
                                    )
                                plain.append(value)
                        value = "".join(plain)
                        target.value = rich if len(rich) else value
                        # Explicit string semantics even if text begins with '='.
                        target.data_type = "s"
                        first_p = sub.find(HP + "p")
                        run = first_p.find(HP + "run") if first_p is not None else None
                        target.font = chars.get(
                            run.get("charPrIDRef") if run is not None else "0", Font(size=10)
                        )
                        paragraph_node = (
                            paras.get(first_p.get("paraPrIDRef")) if first_p is not None else None
                        )
                        align = paragraph_node.find(HH + "align") if paragraph_node is not None else None
                        horizontal = align.get("horizontal", "LEFT").lower() if align is not None else "left"
                        horizontal = {"distribute": "distributed"}.get(horizontal, horizontal)
                        target.alignment = Alignment(
                            horizontal=horizontal,
                            vertical={"CENTER": "center", "BOTTOM": "bottom"}.get(
                                sub.get("vertAlign"), "top"
                            ),
                            wrap_text=True,
                        )
                        fill = fills.get(cell.get("borderFillIDRef"))
                        if fill is not None:
                            target.border = Border(
                                **{
                                    k: _side(fill.find(HH + k + "Border"))
                                    for k in ("top", "bottom", "left", "right")
                                }
                            )
                            brush = fill.find(".//" + HC + "winBrush")
                            if brush is not None and brush.get("faceColor", "none") != "none":
                                target.fill = PatternFill("solid", fgColor=_rgb(brush.get("faceColor")))
                            brush_image = fill.find(".//" + HC + "imgBrush/" + HC + "img")
                            if brush_image is not None:
                                images_meta.append(
                                    _add_image(sheet, archive, binaries, brush_image, r, c, width, height)
                                )
                                warnings.append(f"{target.coordinate}: 셀 배경 이미지를 셀 위 그림으로 변환")
                        for picture in sub.iter(HP + "pic"):
                            picture_size = picture.find(HP + "sz")
                            picture_image = picture.find(HC + "img")
                            pw, ph = int(picture_size.get("width")), int(picture_size.get("height"))
                            # Center inline images within the original cell rectangle.
                            images_meta.append(
                                _add_image(
                                    sheet,
                                    archive,
                                    binaries,
                                    picture_image,
                                    r,
                                    c,
                                    pw,
                                    ph,
                                    max(0, (width - pw) / 2),
                                    max(0, (height - ph) / 2),
                                )
                            )
                        if any(n.tag in (HP + "tbl", HP + "rect", HP + "container") for n in sub.iter()):
                            warnings.append(f"{target.coordinate}: 셀 내부 도형/중첩표 확인 필요")
                        margin = cell.find(HP + "cellMargin")
                        effective_margin = (
                            margin if cell.get("hasMargin") == "1" else table.find(HP + "inMargin")
                        )
                        hints[target.coordinate] = {
                            "padding": [str(v) for v in margins(effective_margin)],
                            "text_sha256": digest(value),
                            "align": "distribute" if horizontal == "distributed" else horizontal,
                            "paragraphs": [
                                {
                                    "length": len(
                                        "".join(
                                            _text(t)
                                            for run in p.findall(HP + "run")
                                            for t in run.findall(HP + "t")
                                        )
                                    ),
                                    "style": paragraph_style(paras[p.get("paraPrIDRef")]),
                                }
                                for p in sub.findall(HP + "p")
                            ],
                            "size": [str(mm(width)), str(mm(height))],
                            "span": [rs, cs],
                            "header": cell.get("header") == "1",
                            "borders": [
                                dict(fill.find(HH + k + "Border").attrib)
                                if fill is not None and fill.find(HH + k + "Border") is not None
                                else {"type": "NONE", "width": "0.12 mm", "color": "#000000"}
                                for k in ("left", "right", "top", "bottom")
                            ],
                        }
                        cells_meta.append(
                            {
                                "cell": target.coordinate,
                                "text": value,
                                "row_span": rs,
                                "col_span": cs,
                                "width_hwpunit": width,
                                "height_hwpunit": height,
                                "margin_hwpunit": dict(margin.attrib) if margin is not None else {},
                            }
                        )
                        if rs > 1 or cs > 1:
                            merges.append((r + 1, c + 1, r + rs, c + cs))
                widths = _dimensions(nc, constraints_x)
                heights = _dimensions(nr, constraints_y)
                for i, width in enumerate(widths, 1):
                    pixels = width * 96 / 7200
                    sheet.column_dimensions[get_column_letter(i)].width = max(0.1, (pixels - 5) / 7)
                for i, height in enumerate(heights, 1):
                    sheet.row_dimensions[i].height = height / 100
                for r, c, rr, cc in merges:
                    sheet.merge_cells(start_row=r, start_column=c, end_row=rr, end_column=cc)
                host = next(p for p in root.iter(HP + "p") if table in list(p.iter(HP + "tbl")))
                hints["A1"]["table"] = {
                    "rows": nr,
                    "columns": nc,
                    "host_style": paragraph_style(paras[host.get("paraPrIDRef")]),
                    "page_break": table.get("pageBreak", "CELL"),
                    "repeat_header": table.get("repeatHeader") == "1",
                    "inline": table.find(HP + "pos").get("treatAsChar") == "1",
                    "size": [str(mm(table.find(HP + "sz").get(k))) for k in ("width", "height")],
                    "padding": [str(v) for v in margins(table.find(HP + "inMargin"))],
                    "margin": [str(v) for v in margins(table.find(HP + "outMargin"))],
                    "widths": [str(v) for v in widths],
                    "heights": [str(v) for v in heights],
                    "excel_widths": [
                        sheet.column_dimensions[get_column_letter(i)].width for i in range(1, nc + 1)
                    ],
                    "excel_heights": [sheet.row_dimensions[i].height for i in range(1, nr + 1)],
                }
                for coordinate, hint in hints.items():
                    hint["excel_borders"] = [
                        border_signature(getattr(sheet[coordinate].border, k))
                        for k in ("left", "right", "top", "bottom")
                    ]
                    write_note(sheet[coordinate], hint)
                sheet.sheet_view.showGridLines = False
                sheet.print_options.horizontalCentered = True
                sheet.sheet_properties.pageSetUpPr.fitToPage = True
                sheet.page_setup.orientation = "landscape" if sum(widths) > 53575 else "portrait"
                sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
                sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
                sheet.print_area = f"A1:{get_column_letter(nc)}{nr}"
                manifest["tables"].append(
                    {
                        "resource": sheet.title,
                        "sheet": sheet.title,
                        "section": section,
                        "source_id": table.get("id"),
                        "rows": nr,
                        "columns": nc,
                        "range": sheet.print_area,
                        "column_widths_hwpunit": widths,
                        "row_heights_hwpunit": heights,
                        "cells": cells_meta,
                        "images": images_meta,
                        "warnings": warnings,
                    }
                )
        if not book.sheetnames:
            raise BuildError("NO_TABLES", "표가 없는 HWPX입니다.", source)
        variables_book = None
        if variables_profile is not None:
            manifest["variables"] = _resolve_variables(Path(variables_profile), manifest)
            variables_book = Workbook()
            variables_book.remove(variables_book.active)
            _variables_sheet(variables_book, manifest)
            manifest["variables_resource"]["path"] = variables_output.name
        output.parent.mkdir(parents=True, exist_ok=True)
        with ExitStack() as stack:
            temporary = Path(stack.enter_context(tempfile.TemporaryDirectory(dir=output.parent)))
            table_file = temporary / "tables.xlsx"
            book.save(table_file)
            repair_workbook(Path(table_file))
            _verify(table_file, manifest)
            if variables_book is not None:
                variables_output.parent.mkdir(parents=True, exist_ok=True)
                variables_temp = Path(
                    stack.enter_context(tempfile.TemporaryDirectory(dir=variables_output.parent))
                )
                variable_file = variables_temp / "variables.xlsx"
                variables_book.save(variable_file)
                repair_workbook(Path(variable_file))
                _verify_variables(variable_file, manifest)
            staged_manifest = temporary / "manifest.json"
            staged_manifest.write_text(
                json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
            )
            # Publish only after every workbook has been reopened and checked.
            os.replace(table_file, output)
            if variables_book is not None:
                os.replace(variable_file, variables_output)
            os.replace(staged_manifest, manifest_path)
        return manifest


def _resolve_variables(profile_path, manifest):
    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    if profile.get("schema_version") != "1.0" or profile.get("type") != "extraction_profile":
        raise BuildError("INVALID_PROFILE", "추출 프로필 버전/타입이 올바르지 않습니다.", profile_path)
    result, seen = [], set()
    for entry in profile["variables"]:
        name = entry["name"]
        if not name or name in seen:
            raise BuildError("DUPLICATE_VARIABLE", f"비어 있거나 중복된 변수명: {name}", profile_path)
        seen.add(name)
        selector = entry["source"]
        if set(selector) != {"paragraph"}:
            raise BuildError("INVALID_PROFILE", "문서 변수는 표 밖 문단에서만 추출합니다.", profile_path)
        pattern = selector["paragraph"]
        candidates = [{"value": p["text"], "source": p} for p in manifest["source_paragraphs"]]
        matches = [(re.fullmatch(pattern, c["value"]), c["source"]) for c in candidates]
        matches = [(m, s) for m, s in matches if m is not None]
        if len(matches) != 1:
            raise BuildError(
                "VARIABLE_EXTRACTION_FAILED",
                f"{name}: 원문에서 정확히 하나의 값이 필요합니다. 발견: {len(matches)}",
                profile_path,
            )
        match, origin = matches[0]
        result.append(
            {"name": name, "value": match.group("value"), "legacy": entry["legacy"], "source": origin}
        )
    return result


def _variables_sheet(book, manifest):
    variables = book.create_sheet("variables")
    variables.append(["name", "value"])
    for item in manifest["variables"]:
        variables.append([item["name"], item["value"]])
        variables.cell(variables.max_row, 2).data_type = "s"
    variables.column_dimensions["A"].width = 60
    variables.column_dimensions["B"].width = 65
    variables.freeze_panes = "B2"
    variables.sheet_view.showGridLines = False
    for cell in variables[1]:
        cell.font = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="334155")
    for row in variables.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="맑은 고딕", size=11)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        variables.row_dimensions[row[0].row].height = 32 if len(str(row[1].value)) > 65 else 20
    manifest["variables_resource"] = {
        "sheet": "variables",
        "range": f"A1:B{variables.max_row}",
        "header": True,
        "naming": "기존 Variables 필드의 이름을 사용한 평면 키. 점(.)은 키 이름의 일부.",
    }


def _verify(output, manifest):
    # Reopen the actual artifact and compare every anchor, span and dimension.
    check = load_workbook(output, rich_text=True)
    for item in manifest["tables"]:
        sheet = check[item["sheet"]]
        for i, width in enumerate(item["column_widths_hwpunit"], 1):
            expected = max(0.1, (width * 96 / 7200 - 5) / 7)
            if abs(sheet.column_dimensions[get_column_letter(i)].width - expected) > 1e-8:
                raise BuildError("EXTRACTION_MISMATCH", f"열 너비 불일치: {item['sheet']}", output)
        for i, height in enumerate(item["row_heights_hwpunit"], 1):
            if abs(sheet.row_dimensions[i].height - height / 100) > 1e-8:
                raise BuildError("EXTRACTION_MISMATCH", f"행 높이 불일치: {item['sheet']}", output)
        for cell in item["cells"]:
            actual = sheet[cell["cell"]].value
            if str(actual if actual is not None else "") != cell["text"]:
                raise BuildError(
                    "EXTRACTION_MISMATCH", f"셀 값 불일치: {item['sheet']}!{cell['cell']}", output
                )
            if (cell["row_span"] > 1 or cell["col_span"] > 1) and not any(
                m.start_cell.coordinate == cell["cell"]
                and m.size == {"rows": cell["row_span"], "columns": cell["col_span"]}
                for m in sheet.merged_cells.ranges
            ):
                raise BuildError("EXTRACTION_MISMATCH", f"병합 불일치: {cell['cell']}", output)
        if len(sheet._images) != len(item["images"]):
            raise BuildError("EXTRACTION_MISMATCH", f"그림 개수 불일치: {item['sheet']}", output)
    check.close()
    manifest["validation"] = {
        "reopened": True,
        "table_count": len(manifest["tables"]),
        "anchor_cells_checked": sum(len(t["cells"]) for t in manifest["tables"]),
        "images_checked": sum(len(t["images"]) for t in manifest["tables"]),
        "variables_checked": len(manifest["variables"]),
    }


def _verify_variables(output, manifest):
    book = load_workbook(output)
    try:
        sheet = book["variables"]
        if sheet.max_column != 2:
            raise BuildError("EXTRACTION_MISMATCH", "변수 표는 반드시 2열이어야 합니다.", output)
        for row, item in enumerate(manifest["variables"], 2):
            actual = sheet.cell(row, 2).value
            if actual is None:
                actual = ""
            if sheet.cell(row, 1).value != item["name"] or actual != item["value"]:
                raise BuildError("EXTRACTION_MISMATCH", f"변수 불일치: {item['name']}", output)
    finally:
        book.close()
