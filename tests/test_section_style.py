"""Absolute paragraph depth is a section policy, independent of parent/sibling paragraphs."""

from decimal import Decimal
from zipfile import ZipFile

import pytest
import yaml
from defusedxml import ElementTree as ET

from hwpx_builder import BuildError, build, load_project
from hwpx_builder.units import mm_to_hwpunit

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"


def save(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")


def project(root, contents=(), defaults=None, sections=None):
    save(
        root / "document.yaml",
        {
            "schema_version": "1.0",
            "type": "document",
            "defaults": defaults or {},
            "sections": sections
            if sections is not None
            else [{"layer": {"type": "single_column", "contents": list(contents)}}],
        },
    )


def test_absolute_depth_without_parent_and_with_skipped_levels(tmp_path):
    project(
        tmp_path,
        [{"type": "paragraph", "text": str(depth), "indent_depth": depth} for depth in (2, 0, 3, 1)],
        {"section_style": {"indent_by_depth": [0, 4, 8, 12]}},
    )
    document = load_project(tmp_path)
    paragraphs = document.sections[0].contents
    assert [p.indent_depth for p in paragraphs] == [2, 0, 3, 1]
    assert [p.style.indent for p in paragraphs] == [8, 0, 12, 4]
    assert document.sections[0].style.indent_by_depth == tuple(map(Decimal, (0, 4, 8, 12)))
    result = build(tmp_path, "out.hwpx")
    with ZipFile(result.output) as archive:
        header = ET.fromstring(archive.read("Contents/header.xml"))
        section = ET.fromstring(archive.read("Contents/section0.xml"))
    styles = {p.get("id"): p for p in header.iter(HH + "paraPr")}
    assert [
        int(styles[p.get("paraPrIDRef")].find(".//" + HH + "margin/" + HC + "left").get("value"))
        for p in section.findall(HP + "p")
    ] == [mm_to_hwpunit(x) for x in (8, 0, 12, 4)]


def test_section_override_replaces_list_and_does_not_leak(tmp_path):
    save(
        tmp_path / "styles/depth.yaml",
        {"schema_version": "1.0", "type": "section_style", "indent_by_depth": [0, 5, 10, 15]},
    )
    save(
        tmp_path / "sections/first.yaml",
        {
            "schema_version": "1.0",
            "type": "section",
            "section_style": {"indent_by_depth": [0, 2]},
            "layer": {
                "type": "single_column",
                "contents": [{"type": "paragraph", "text": "첫째", "indent_depth": 1}],
            },
        },
    )
    project(
        tmp_path,
        defaults={"section_style": {"src": "styles/depth.yaml"}},
        sections=[
            {"src": "sections/first.yaml"},
            {
                "layer": {
                    "type": "single_column",
                    "contents": [{"type": "paragraph", "text": "둘째", "indent_depth": 2}],
                }
            },
        ],
    )
    document = load_project(tmp_path)
    assert document.sections[0].style.indent_by_depth == (0, 2)
    assert document.sections[0].contents[0].style.indent == 2
    assert document.sections[1].contents[0].style.indent == 10
    first = yaml.safe_load((tmp_path / "sections/first.yaml").read_text(encoding="utf-8"))
    first["layer"]["contents"][0]["indent_depth"] = 2
    save(tmp_path / "sections/first.yaml", first)
    with pytest.raises(BuildError) as caught:
        load_project(tmp_path)
    diagnostic = caught.value.diagnostic
    assert diagnostic.code == "INVALID_INDENT_DEPTH"
    assert diagnostic.source_path == str(tmp_path / "sections/first.yaml")
    assert diagnostic.pointer.endswith("/indent_depth")
    assert diagnostic.line and diagnostic.column
    assert diagnostic.details["reference_chain"][0]["source_path"] == str(tmp_path / "document.yaml")


def test_depth_inside_include_container_and_table_notes(tmp_path):
    from openpyxl import Workbook

    book = Workbook()
    book.active["A1"] = "표"
    book.save(tmp_path / "table.xlsx")
    save(
        tmp_path / "part.yaml",
        {
            "schema_version": "1.0",
            "type": "blocks",
            "contents": [{"type": "paragraph", "indent_depth": 2, "text": "포함"}],
        },
    )
    project(
        tmp_path,
        [
            {"type": "container", "contents": [{"type": "include", "src": "part.yaml"}]},
            {
                "type": "table",
                "resource": "t",
                "headernote": {"text": "위", "indent_depth": 1},
                "footernote": {"text": "아래", "indent_depth": 3},
            },
        ],
        {"section_style": {"indent_by_depth": [0, 3, 6, 9]}},
    )
    data = yaml.safe_load((tmp_path / "document.yaml").read_text(encoding="utf-8"))
    data["resources"] = {"tables": {"t": {"path": "table.xlsx"}}}
    save(tmp_path / "document.yaml", data)
    blocks = load_project(tmp_path).sections[0].contents
    assert [blocks[i].style.indent for i in (0, 1, 3)] == [6, 3, 9]
    assert blocks[2].data.cells[0].paragraph.style.indent == 0


@pytest.mark.parametrize("depth", [-1, True, 1.5, "1", None, 4])
def test_invalid_depth_is_reported_at_the_value(tmp_path, depth):
    project(
        tmp_path,
        [{"type": "paragraph", "text": "문단", "indent_depth": depth}],
        {"section_style": {"indent_by_depth": [0, 5]}},
    )
    with pytest.raises(BuildError) as caught:
        load_project(tmp_path)
    d = caught.value.diagnostic
    assert d.code == "INVALID_INDENT_DEPTH"
    lines = (tmp_path / "document.yaml").read_text(encoding="utf-8").splitlines()
    line = next(i for i, s in enumerate(lines, 1) if "indent_depth:" in s)
    assert d.line == line
    assert d.column == lines[line - 1].index(":") + 3


@pytest.mark.parametrize("values", [[], [0, -1], [0, True], [0, "5"], [0, 5, 2]])
def test_invalid_section_style_has_actual_external_file_location(tmp_path, values):
    save(
        tmp_path / "styles/depth.yaml",
        {"schema_version": "1.0", "type": "section_style", "indent_by_depth": values},
    )
    project(tmp_path, defaults={"section_style": {"src": "styles/depth.yaml"}})
    with pytest.raises(BuildError) as caught:
        load_project(tmp_path)
    d = caught.value.diagnostic
    assert d.code == "INVALID_VALUE"
    assert d.source_path == str(tmp_path / "styles/depth.yaml")
    assert d.line and d.column


def test_no_implicit_indent_steps_and_no_legacy_paragraph_indent(tmp_path):
    project(tmp_path, [{"type": "paragraph", "text": "문단", "indent_depth": 1}])
    with pytest.raises(BuildError, match="INVALID_INDENT_DEPTH"):
        load_project(tmp_path)
    project(tmp_path, [{"type": "paragraph", "text": "문단", "paragraph_style": {"indent": 5}}])
    with pytest.raises(BuildError, match="UNKNOWN_FIELD"):
        load_project(tmp_path)


def test_yaml_syntax_error_and_missing_src_locations(tmp_path):
    project(tmp_path, defaults={"section_style": {"src": "styles/depth.yaml"}})
    path = tmp_path / "styles/depth.yaml"
    path.parent.mkdir()
    path.write_text('schema_version: "1.0"\ntype: section_style\nindent_by_depth: [0, 5\n', encoding="utf-8")
    with pytest.raises(BuildError) as caught:
        load_project(tmp_path)
    d = caught.value.diagnostic
    assert d.source_path == str(path) and d.line == 4 and d.column == 1
    assert d.details["reference_chain"][0]["pointer"] == "/defaults/section_style/src"
    path.unlink()
    with pytest.raises(BuildError) as caught:
        load_project(tmp_path)
    d = caught.value.diagnostic
    assert d.code == "DECLARATION_READ_FAILED"
    assert d.source_path == str(tmp_path / "document.yaml")
    assert d.pointer == "/defaults/section_style/src"
    assert d.line and d.column


def test_unknown_key_points_to_key_in_external_style(tmp_path):
    save(tmp_path / "styles/p.yaml", {"schema_version": "1.0", "type": "paragraph_style", "algin": "left"})
    project(tmp_path, [{"type": "paragraph", "text": "문단", "paragraph_style": {"src": "styles/p.yaml"}}])
    with pytest.raises(BuildError) as caught:
        load_project(tmp_path)
    d = caught.value.diagnostic
    assert d.code == "UNKNOWN_FIELD"
    assert d.source_path == str(tmp_path / "styles/p.yaml") and d.line == 3 and d.column == 1


def test_fixed_textbox_uses_section_depth(tmp_path):
    project(
        tmp_path,
        [
            {
                "type": "container",
                "container_style": {"layout": {"mode": "inline", "width": 60, "height": 20}},
                "contents": [
                    {
                        "type": "container",
                        "container_style": {
                            "layout": {
                                "mode": "floating",
                                "relative_to": "paragraph",
                                "x": 0,
                                "y": 0,
                                "width": 60,
                                "height": 20,
                            }
                        },
                        "contents": [{"type": "paragraph", "text": "글상자", "indent_depth": 2}],
                    }
                ],
            }
        ],
        {"section_style": {"indent_by_depth": [0, 3, 6]}},
    )
    document = load_project(tmp_path)
    paragraph = document.sections[0].contents[0].contents[0].paragraphs[0]
    assert paragraph.indent_depth == 2 and paragraph.style.indent == 6
    build(tmp_path, "textbox.hwpx")


def test_null_section_policy_inherits_and_depth_zero_is_explicit_value(tmp_path):
    project(
        tmp_path,
        defaults={"section_style": {"indent_by_depth": [2, 4]}},
        sections=[
            {
                "section_style": {"indent_by_depth": None},
                "layer": {"type": "single_column", "contents": [{"type": "paragraph", "text": "상속"}]},
            },
            {
                "section_style": None,
                "layer": {"type": "single_column", "contents": [{"type": "paragraph", "text": "상속"}]},
            },
        ],
    )
    assert [s.contents[0].style.indent for s in load_project(tmp_path).sections] == [2, 2]
