from pathlib import Path
from zipfile import ZipFile

import pytest
import yaml
from defusedxml import ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from hwpx_builder import BuildError, build, load_project
from hwpx_builder.model import TextStyle
from hwpx_builder.package import validate_package
from hwpx_builder.text import parse_text
from hwpx_builder.units import mm_to_hwpunit

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
REF = Path(__file__).resolve().parents[1] / "template/data"


def project(root, blocks=(), **extra):
    d = dict(
        schema_version="1.0",
        type="document",
        sections=[{"layer": {"type": "single_column", "contents": list(blocks)}}],
        **extra,
    )
    (root / "document.yaml").write_text(yaml.safe_dump(d, allow_unicode=True), encoding="utf-8")
    return root


def vars_book(root, rows=(("name", "value"), ("회사", "<bold>리터럴</bold>")), **kw):
    w = Workbook()
    w.active.title = "vars"
    for row in rows:
        w.active.append(row)
    w.save(root / "vars.xlsx")
    return {"tables": {"vars": {"path": "vars.xlsx", "sheet": "vars", "header": True, **kw}}}


def error(root, code):
    with pytest.raises(BuildError) as caught:
        load_project(root)
    assert caught.value.diagnostic.code == code
    return caught.value.diagnostic


def test_empty_and_deterministic(tmp_path):
    project(tmp_path)
    a = build(tmp_path, "a.hwpx")
    b = build(tmp_path, "b.hwpx")
    assert a.output.read_bytes() == b.output.read_bytes()
    validate_package(a.output)


def test_variables_notes_and_tags(tmp_path):
    resources = vars_book(tmp_path)
    w = Workbook()
    s = w.active
    s.title = "완성 표"
    s.append(["제목", None, "금액"])
    s.append(["내용", "B", 1234])
    s.merge_cells("A1:B1")
    s["A1"].font = Font(name="맑은 고딕", size=13, bold=True)
    s["A1"].fill = PatternFill("solid", fgColor="FFF2CC")
    s["A1"].alignment = Alignment(horizontal="center")
    s["C2"].number_format = "#,##0"
    w.save(tmp_path / "table.xlsx")
    resources["tables"]["t"] = {"path": "table.xlsx", "sheet": "완성 표"}
    project(
        tmp_path,
        [
            {
                "type": "table",
                "resource": "t",
                "headernote": {"text": "앞 <ref:회사>"},
                "footernote": {"text": "뒤 <bold>끝</bold>"},
            }
        ],
        variables={"resource": "vars"},
        resources=resources,
    )
    d = load_project(tmp_path)
    table = d.sections[0].contents[1]
    assert table.data.cells[0].col_span == 2
    assert table.data.cells[0].paragraph.runs[0].style.bold
    assert table.data.cells[0].paragraph.runs[0].style.font_family == "맑은 고딕"
    a = build(tmp_path, "out.hwpx")
    b = build(tmp_path, "again.hwpx")
    assert a.output.read_bytes() == b.output.read_bytes()
    with ZipFile(a.output) as z:
        s = ET.fromstring(z.read("Contents/section0.xml"))
        assert "1,234" in "".join(s.itertext())
        assert "<bold>리터럴</bold>" in "".join(s.itertext())
        assert len(list(s.iter(HP + "tbl"))) == 1
        assert "".join(s[0].itertext()).startswith("앞 ")
        assert "".join(s[-1].itertext()) == "뒤 끝"


@pytest.mark.parametrize(
    "text,code",
    [
        ("<bold>", "INVALID_TEXT_TAG"),
        ("<foo>", "UNKNOWN_TEXT_TAG"),
        ("<bold><italic>x</bold></italic>", "INVALID_TEXT_TAG"),
        ("<size:0>x</size>", "INVALID_VALUE"),
        ("<color:invalid>x</color>", "INVALID_VALUE"),
        ("<ref:missing>", "UNRESOLVED_REFERENCE"),
        ("<nat>abc", "INVALID_TEXT_TAG"),
    ],
)
def test_invalid_tags(text, code):
    with pytest.raises(BuildError) as e:
        parse_text(text, TextStyle(), {})
    assert e.value.diagnostic.code == code


def test_text_literal_and_linebreak():
    runs = parse_text("<bold>A<nat><ref:x>&lt;<\\nat>B</bold><linebreak>&lt;!", TextStyle(), {})
    assert "".join(r.text for r in runs) == "A<ref:x>&lt;B<!"
    assert sum(r.linebreak for r in runs) == 1
    assert runs[0].style.bold
    assert not runs[-1].style.bold


def test_units():
    assert mm_to_hwpunit(25.4) == 7200
    assert mm_to_hwpunit("0.001763888888888889") == 1


def test_include_and_toml_styles(tmp_path):
    (tmp_path / "s.toml").write_text(
        'schema_version="1.0"\ntype="text_style"\nbold=true\nfont_size=14\n', encoding="utf-8"
    )
    (tmp_path / "part.yaml").write_text(
        'schema_version: "1.0"\ntype: blocks\ncontents:\n  - type: paragraph\n    text: 인용\n    text_style: {src: s.toml}\n',
        encoding="utf-8",
    )
    project(tmp_path, [{"type": "include", "src": "part.yaml"}])
    d = load_project(tmp_path)
    assert d.sections[0].contents[0].runs[0].style.font_size == 14
    assert d.sections[0].contents[0].source.path == tmp_path / "part.yaml"


def test_text_only_inheritance(tmp_path):
    project(
        tmp_path,
        [
            {
                "type": "container",
                "text_style": {"bold": False},
                "contents": [{"type": "paragraph", "text": "자식"}],
            }
        ],
        defaults={"text_style": {"bold": True}},
    )
    d = load_project(tmp_path)
    assert not d.sections[0].contents[0].runs[0].style.bold


@pytest.mark.parametrize(
    "block,code",
    [
        (
            {"type": "table", "resource": "x", "text_style": {"bold": True}},
            "RESOURCE_TABLE_OVERRIDE_FORBIDDEN",
        ),
        ({"type": "table", "rows": []}, "TABLE_RESOURCE_REQUIRED"),
        ({"type": "image", "resource": "x"}, "IMAGE_REQUIRES_CONTAINER"),
        (
            {"type": "container", "contents": [], "container_style": {"layout": {"mode": "floating"}}},
            "UNSUPPORTED_RENDERING",
        ),
        ({"type": "paragraph", "text": "x", "paragraph_style": {"algin": "left"}}, "UNKNOWN_FIELD"),
        ({"type": "paragraph", "text": "x", "text_style": {"bold": "false"}}, "INVALID_VALUE"),
        ({"type": "paragraph", "text": "x", "runs": []}, "INVALID_VALUE"),
        ({"type": "condition", "when": "x", "contents": []}, "UNSUPPORTED_RENDERING"),
    ],
)
def test_invalid_blocks(tmp_path, block, code):
    project(tmp_path, [block])
    error(tmp_path, code)


def test_source_diagnostic(tmp_path):
    (tmp_path / "document.yaml").write_text(
        'schema_version: "1.0"\ntype: document\nsections: []\nwrong: 1\n', encoding="utf-8"
    )
    d = error(tmp_path, "UNKNOWN_FIELD")
    assert d.pointer == "/wrong" and d.line == 4


def test_duplicate_and_inline_variables(tmp_path):
    resources = vars_book(tmp_path, rows=(("name", "value"), ("x", "a"), ("x", "b")))
    project(tmp_path, resources=resources, variables={"resource": "vars"})
    error(tmp_path, "DUPLICATE_VARIABLE")
    project(tmp_path, variables={"x": "inline"})
    error(tmp_path, "VARIABLE_RESOURCE_REQUIRED")
    (tmp_path / "document.yaml").write_text("type: document\ntype: document\n", encoding="utf-8")
    error(tmp_path, "DUPLICATE_KEY")


def test_two_column_variables(tmp_path):
    resources = vars_book(tmp_path, rows=(("name", "value", "extra"), ("x", "a", "b")))
    project(tmp_path, resources=resources, variables={"resource": "vars"})
    error(tmp_path, "VARIABLE_TABLE_COLUMNS")


def test_escape_cycle_symlink(tmp_path):
    project(tmp_path, [{"type": "include", "src": "../outside.yaml"}])
    error(tmp_path, "DECLARATION_OUTSIDE_ROOT")
    (tmp_path / "cycle.yaml").write_text(
        'schema_version: "1.0"\ntype: blocks\ncontents: [{type: include, src: cycle.yaml}]', encoding="utf-8"
    )
    project(tmp_path, [{"type": "include", "src": "cycle.yaml"}])
    error(tmp_path, "CYCLIC_INCLUDE")
    link = tmp_path / "outside"
    try:
        link.symlink_to(tmp_path.parent, target_is_directory=True)
    except OSError:
        pytest.skip("symlinks unavailable")
    project(tmp_path, [{"type": "include", "src": "outside/file.yaml"}])
    error(tmp_path, "DECLARATION_OUTSIDE_ROOT")


def test_formula_and_atomic_output(tmp_path):
    w = Workbook()
    w.active["A1"] = "=1+1"
    w.save(tmp_path / "table.xlsx")
    project(
        tmp_path, [{"type": "table", "resource": "t"}], resources={"tables": {"t": {"path": "table.xlsx"}}}
    )
    out = tmp_path / "out.hwpx"
    out.write_bytes(b"previous")
    with pytest.raises(BuildError, match="FORMULA_CACHE_MISSING"):
        build(tmp_path, out)
    assert out.read_bytes() == b"previous"


def test_pagebreak_and_image(tmp_path):
    project(
        tmp_path,
        [
            {"type": "paragraph", "text": "첫 쪽"},
            {"type": "pagebreak"},
            {
                "type": "container",
                "contents": [{"type": "image", "resource": "logo", "image_style": {"width": 20}}],
            },
        ],
        resources={"images": {"logo": {"path": str(REF / "demo-logo.png")}}},
    )
    result = build(tmp_path, "image.hwpx")
    with ZipFile(result.output) as z:
        xml = ET.fromstring(z.read("Contents/section0.xml"))
        pics = list(xml.iter(HP + "pic"))
        assert len(pics) == 1
        assert any(p.get("pageBreak") == "1" for p in xml.findall(HP + "p"))


def test_warnings_as_errors(tmp_path):
    w = Workbook()
    w.active["A1"] = "표"
    w.save(tmp_path / "t.xlsx")
    project(tmp_path, [{"type": "table", "resource": "t"}], resources={"tables": {"t": {"path": "t.xlsx"}}})
    with pytest.raises(BuildError, match="XLSX_COLUMN_WIDTH_APPROXIMATION"):
        build(tmp_path, "out.hwpx", warnings_as_errors=True)
    assert not (tmp_path / "out.hwpx").exists()


def test_debug_highlights_references_only(tmp_path):
    resources = vars_book(tmp_path, rows=(("name", "value"), ("회사", "주입값")))
    project(
        tmp_path,
        [{"type": "paragraph", "text": "고정 <bold><ref:회사></bold> 뒤 <ref:회사>"}],
        variables={"resource": "vars"},
        resources=resources,
    )
    normal = build(tmp_path, "normal.hwpx")
    debug = build(tmp_path, "debug.hwpx", debug=True)

    def runs(path):
        with ZipFile(path) as archive:
            header = ET.fromstring(archive.read("Contents/header.xml"))
            shades = {
                p.get("id"): p.get("shadeColor")
                for p in header.iter("{http://www.hancom.co.kr/hwpml/2011/head}charPr")
            }
            section = ET.fromstring(archive.read("Contents/section0.xml"))
            return [
                ("".join(r.itertext()), shades[r.get("charPrIDRef")])
                for r in section.iter(HP + "run")
                if r.find(HP + "t") is not None
            ]

    assert not any(shade == "#FFFF00" for _, shade in runs(normal.output))
    debug_runs = runs(debug.output)
    assert [text for text, shade in debug_runs if shade == "#FFFF00"] == ["주입값", "주입값"]
    assert all(shade != "#FFFF00" for text, shade in debug_runs if text != "주입값")


def test_default_excel_font_uses_dark_theme(tmp_path):
    from hwpx_builder.model import Source
    from hwpx_builder.resources import Resources

    w = Workbook()
    w.active["A1"] = "내용"
    w.save(tmp_path / "t.xlsx")
    resources = Resources(tmp_path, {"t": ({"path": "t.xlsx"}, Source(tmp_path / "document.yaml"))}, {})
    data = resources.table("t", Source())
    assert data.cells[0].paragraph.runs[0].style.color == "#000000"


def test_output_is_project_relative_from_other_cwd(tmp_path, monkeypatch):
    root = tmp_path / "프로젝트"
    root.mkdir()
    project(root, [{"type": "paragraph", "text": "본문"}])
    other = tmp_path / "다른 곳"
    other.mkdir()
    monkeypatch.chdir(other)
    result = build(root, "result.hwpx")
    assert result.output == root / "result.hwpx"
    assert not (other / "result.hwpx").exists()


def test_preview_and_atomic_write_failure(tmp_path, monkeypatch):
    from hwpx_builder import compiler

    project(tmp_path, [{"type": "paragraph", "text": "미리보기 본문"}])
    result = build(tmp_path, "out.hwpx")
    with ZipFile(result.output) as archive:
        assert "미리보기 본문" in archive.read("Preview/PrvText.txt").decode("utf-8")
        assert "Preview/PrvImage.png" not in archive.namelist()  # no stale blank template thumbnail
    previous = result.output.read_bytes()

    def fail(path):
        raise BuildError("INVALID_PACKAGE", "injected validation failure", path)

    monkeypatch.setattr(compiler, "validate_package", fail)
    with pytest.raises(BuildError, match="INVALID_PACKAGE"):
        build(tmp_path, "out.hwpx")
    assert result.output.read_bytes() == previous


def test_partial_border_and_no_table_variable_sharing(tmp_path):
    from openpyxl.styles import Border, Side

    resources = vars_book(tmp_path)
    w = Workbook()
    w.active["A1"] = "내용"
    w.active["A1"].border = Border(bottom=Side(style="thin"))
    w.save(tmp_path / "t.xlsx")
    resources["tables"]["t"] = {"path": "t.xlsx"}
    project(
        tmp_path, [{"type": "table", "resource": "t"}], resources=resources, variables={"resource": "vars"}
    )
    assert load_project(tmp_path).sections[0].contents[0].data.cells[0].borders[3].style == "SOLID"
    resources["tables"]["t"] = {"path": "vars.xlsx"}
    project(
        tmp_path, [{"type": "table", "resource": "t"}], resources=resources, variables={"resource": "vars"}
    )
    error(tmp_path, "VARIABLE_TABLE_NOT_INDEPENDENT")


def test_output_cannot_overwrite_unused_registered_resource(tmp_path):
    from shutil import copyfile

    logo = tmp_path / "unused.png"
    copyfile(REF / "demo-logo.png", logo)
    previous = logo.read_bytes()
    project(tmp_path, resources={"images": {"unused": {"path": "unused.png"}}})
    with pytest.raises(BuildError, match="OUTPUT_CONFLICT"):
        build(tmp_path, "unused.png")
    assert logo.read_bytes() == previous


def test_installed_builder_excludes_migration_tools(tmp_path):
    import subprocess
    import sys

    # Isolated Python from outside the repository must import the builder without
    # exposing or importing repository-only migration utilities.
    subprocess.run(
        [
            sys.executable,
            "-I",
            "-c",
            """import importlib.util, sys, hwpx_builder
assert not hasattr(hwpx_builder, 'extract_tables')
assert importlib.util.find_spec('hwpx_builder.extract') is None
assert importlib.util.find_spec('hwpx_builder.assets') is None
assert not any(n == 'scripts' or n.startswith('scripts.') for n in sys.modules)
""",
        ],
        cwd=tmp_path,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    help_result = subprocess.run(
        [sys.executable, "-I", "-m", "hwpx_builder", "--help"],
        cwd=tmp_path,
        check=True,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert "{build,validate}" in help_result.stdout
    assert "extract" not in help_result.stdout


def test_cli_output_is_working_directory_relative(tmp_path, monkeypatch, capsys):
    from hwpx_builder.cli import main

    root = tmp_path / "문서"
    root.mkdir()
    project(root, [{"type": "paragraph", "text": "가상 예제"}])
    monkeypatch.chdir(tmp_path)
    assert main(["build", str(root), "--output", "결과 폴더/report.hwpx"]) == 0
    assert (tmp_path / "결과 폴더/report.hwpx").is_file()
    assert '"output"' in capsys.readouterr().out
    assert main(["build", str(root)]) == 0
    assert (tmp_path / "output.hwpx").is_file()
    absolute = tmp_path / "absolute.hwpx"
    assert main(["build", str(root), "--output", str(absolute)]) == 0
    assert absolute.is_file()
    with pytest.raises(SystemExit) as caught:
        main(["build", str(root), "--output-dir", "removed-option"])
    assert caught.value.code == 2


def test_cli_version_matches_package_metadata(capsys):
    from importlib.metadata import version

    from hwpx_builder.cli import main

    with pytest.raises(SystemExit) as caught:
        main(["--version"])
    assert caught.value.code == 0
    assert capsys.readouterr().out.strip() == f"hwpx-builder {version('hwpx-builder')}"
