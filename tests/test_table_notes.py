"""Table descriptions use the visible table's horizontal frame, not the whole page."""

from dataclasses import replace
from decimal import Decimal as D
from zipfile import ZipFile

import pytest
import yaml
from defusedxml import ElementTree as ET
from openpyxl import Workbook

from hwpx_builder import BuildError, build, load_project
from hwpx_builder.model import Cell, Paragraph, ParagraphStyle, Source, TableData
from hwpx_builder.resources import Resources
from hwpx_builder.units import mm_to_hwpunit

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"


def project(tmp_path, monkeypatch, align="left", inline=True, header=None, footer=None):
    book = Workbook()
    book.active["A1"] = "가상 표"
    book.save(tmp_path / "table.xlsx")
    data = TableData(
        (D(60),),
        (D(10),),
        (Cell(0, 0, 1, 1, Paragraph()),),
        Source(),
        host_style=ParagraphStyle(align=align, indent=D(10), right_indent=D(20), first_line_indent=D(2)),
        margin=(D(3), D(7), D(0), D(0)),
        inline=inline,
    )
    monkeypatch.setattr(Resources, "table", lambda *args: data)
    table = {
        "type": "table",
        "resource": "t",
        "headernote": header or {"text": "위 설명", "paragraph_style": {"align": "left"}},
        "footernote": footer or {"text": "아래 설명", "paragraph_style": {"align": "right"}},
    }
    config = {
        "schema_version": "1.0",
        "type": "document",
        "defaults": {"section_style": {"indent_by_depth": [0, 5, 40, 80]}},
        "resources": {"tables": {"t": {"path": "table.xlsx"}}},
        "sections": [{"layer": {"type": "single_column", "contents": [table]}}],
    }
    (tmp_path / "document.yaml").write_text(yaml.safe_dump(config, allow_unicode=True), encoding="utf-8")
    return data


@pytest.mark.parametrize(
    "align,inline,left,right",
    [
        ("left", True, 15, 75),
        ("center", True, 39, 99),
        ("right", True, 63, 123),
        ("left", False, 0, 60),
        ("center", False, 45, 105),
        ("right", False, 90, 150),
    ],
)
def test_note_frame_tracks_table_alignment_and_visible_edges(
    tmp_path, monkeypatch, align, inline, left, right
):
    project(tmp_path, monkeypatch, align, inline)
    blocks = load_project(tmp_path).sections[0].contents
    for note in (blocks[0], blocks[2]):
        assert note.style.indent == left
        assert 150 - note.style.right_indent == right
    result = build(tmp_path, "out.hwpx")
    with ZipFile(result.output) as archive:
        header = ET.fromstring(archive.read("Contents/header.xml"))
        section = ET.fromstring(archive.read("Contents/section0.xml"))
    styles = {p.get("id"): p for p in header.iter(HH + "paraPr")}
    for paragraph in (section.findall(HP + "p")[0], section.findall(HP + "p")[-1]):
        margin = styles[paragraph.get("paraPrIDRef")].find(".//" + HH + "margin")
        assert int(margin.find(HC + "left").get("value")) == mm_to_hwpunit(left)
        assert int(margin.find(HC + "right").get("value")) == mm_to_hwpunit(150 - right)


def test_existing_inner_indents_and_note_format_remain_configurable(tmp_path, monkeypatch):
    project(
        tmp_path,
        monkeypatch,
        header={
            "text": "위 설명",
            "indent_depth": 2,
            "paragraph_style": {"right_indent": 100, "line_spacing": "125%", "space_before": 3},
        },
    )
    note = load_project(tmp_path).sections[0].contents[0]
    assert note.style.indent == 40 and note.style.right_indent == 100
    assert note.style.line_spacing == 125 and note.style.space_before == 3


@pytest.mark.parametrize(
    "note",
    [
        {"text": "너무 깊은 설명", "indent_depth": 3},
        {"text": "첫 줄 넘침", "paragraph_style": {"first_line_indent": 60}},
        {"text": "내용 폭 없음", "paragraph_style": {"right_indent": 140}},
    ],
)
@pytest.mark.parametrize("role", ["headernote", "footernote"])
def test_impossible_note_indentation_reports_its_yaml_location(tmp_path, monkeypatch, note, role):
    project(tmp_path, monkeypatch, **{"header" if role == "headernote" else "footer": note})
    with pytest.raises(BuildError) as caught:
        load_project(tmp_path)
    error = caught.value.diagnostic
    assert error.code == "TABLE_NOTE_NO_SPACE"
    assert error.pointer.endswith("/" + role)
    assert error.source_path == str(tmp_path / "document.yaml")
    assert error.line and error.column


def test_declared_table_size_is_used_instead_of_grid_width(tmp_path, monkeypatch):
    data = project(tmp_path, monkeypatch, "center")
    monkeypatch.setattr(Resources, "table", lambda *args: replace(data, size=(D(40), D(10))))
    notes = load_project(tmp_path).sections[0].contents
    assert notes[0].style.indent == notes[2].style.indent == 49
    assert notes[0].style.right_indent == notes[2].style.right_indent == 61


def test_include_container_and_ordinary_paragraph_keep_their_own_scope(tmp_path, monkeypatch):
    project(tmp_path, monkeypatch, "center")
    path = tmp_path / "document.yaml"
    config = yaml.safe_load(path.read_text(encoding="utf-8"))
    table = config["sections"][0]["layer"]["contents"][0]
    included = {
        "schema_version": "1.0",
        "type": "blocks",
        "contents": [{"type": "container", "contents": [table]}],
    }
    (tmp_path / "part.yaml").write_text(yaml.safe_dump(included, allow_unicode=True), encoding="utf-8")
    config["sections"][0]["layer"]["contents"] = [
        {"type": "include", "src": "part.yaml"},
        {"type": "paragraph", "text": "일반 문단", "indent_depth": 1},
    ]
    path.write_text(yaml.safe_dump(config, allow_unicode=True), encoding="utf-8")
    before, table, after, ordinary = load_project(tmp_path).sections[0].contents
    assert before.style.indent == after.style.indent == 39
    assert before.source.path == after.source.path == tmp_path / "part.yaml"
    assert ordinary.style.indent == 5 and ordinary.style.right_indent == 0
