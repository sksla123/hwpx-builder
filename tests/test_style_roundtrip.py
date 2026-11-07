"""Public regression fixtures contain only synthetic factory data."""

from pathlib import Path
from zipfile import ZipFile

import pytest
from defusedxml import ElementTree as ET
from openpyxl import load_workbook
from openpyxl.styles import Font

from hwpx_builder import build, load_project
from hwpx_builder.model import Paragraph

ROOT = Path(__file__).resolve().parents[1]
HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"


def test_public_template_inputs_and_notes():
    for root in (ROOT / "template", ROOT / "examples/factory_report"):
        document = load_project(root)
        assert len(document.variables) == 6
        assert dict(document.variables)["factory_name"] == "농Shim 제 1공장"
        notes = [
            b
            for section in document.sections
            for b in section.contents
            if isinstance(b, Paragraph) and b.source.pointer.endswith(("/headernote", "/footernote"))
        ]
        assert len(notes) == 4
        assert all(p.style.line_spacing == 100 for p in notes)
        assert all(p.style.space_before == p.style.space_after == 0 for p in notes)
        assert notes[1].style.indent == 5
        assert len(document.sections) == 3


def test_public_report_normal_debug_and_determinism(tmp_path):
    a = build(ROOT / "template", tmp_path / "normal.hwpx")
    b = build(ROOT / "template", tmp_path / "debug.hwpx", debug=True)
    again = build(ROOT / "template", tmp_path / "again.hwpx")
    assert a.output.read_bytes() == again.output.read_bytes()
    assert a.tables == b.tables == 4
    assert a.images == b.images == 1
    for result in (a, b):
        with ZipFile(result.output) as archive:
            section = ET.fromstring(archive.read("Contents/section0.xml"))
            text = "".join(section.itertext())
            assert "농Shim 제 1공장 운영보고서" in text
            assert "구름면" in text and "2030. 05." in text
            assert list(section.iter(HP + "lineBreak"))
            header = ET.fromstring(archive.read("Contents/header.xml"))
            shades = {p.get("id"): p.get("shadeColor") for p in header.iter(HH + "charPr")}
            assert all(
                shades[r.get("charPrIDRef")] != "#FFFF00"
                for table in section.iter(HP + "tbl")
                for r in table.iter(HP + "run")
            )
            assert any(v == "#FFFF00" for v in shades.values()) == (result == b)


def test_excel_edit_and_font_remain_authoritative(tmp_path):
    import shutil

    import yaml

    shutil.copytree(ROOT / "template", tmp_path / "project")
    root = tmp_path / "project"
    book = load_workbook(root / "data/tables.xlsx")
    sheet = book["table_03"]
    sheet["A2"] = "합성 제품"
    sheet["A2"].font = Font(name="맑은 고딕", size=13, bold=True)
    book.save(root / "data/tables.xlsx")
    book.close()
    config = yaml.safe_load((root / "sections/01-production.yaml").read_text(encoding="utf-8"))
    note = config["layer"]["contents"][2]["headernote"]
    note["paragraph_style"] = {"line_spacing": "125%", "space_before": 3}
    (root / "sections/01-production.yaml").write_text(
        yaml.safe_dump(config, allow_unicode=True), encoding="utf-8"
    )
    document = load_project(root)
    blocks = document.sections[1].contents
    assert blocks[2].style.line_spacing == 125
    assert blocks[2].style.space_before == 3
    cell = blocks[3].data.cells[3]
    assert cell.paragraph.runs[0].text == "합성 제품"
    assert cell.paragraph.runs[0].style.font_family == "맑은 고딕"
    assert cell.paragraph.runs[0].style.font_size == 13


@pytest.mark.parametrize("project", ["template", "examples/factory_report", "examples/variables"])
def test_public_projects_validate(project):
    assert load_project(ROOT / project).sections
