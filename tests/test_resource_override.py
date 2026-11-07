"""Select report materials independently of the template's declaration paths."""

import json
from pathlib import Path
from zipfile import ZipFile

import pytest
import yaml
from openpyxl import Workbook
from PIL import Image

from hwpx_builder import BuildError, build, load_project
from hwpx_builder.cli import main


def save(path, kind, **data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        yaml.safe_dump({"schema_version": "1.0", "type": kind, **data}, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )


def materials(root, value):
    (root / "data").mkdir(parents=True)
    (root / "assets").mkdir()
    book = Workbook()
    book.active.append(["name", "value"])
    book.active.append(["month", value])
    book.save(root / "data/variables.xlsx")
    book.close()
    book = Workbook()
    book.active["A1"] = value
    book.save(root / "data/tables.xlsx")
    book.close()
    Image.new("RGB", (2, 2), "blue").save(root / "assets/logo.png")
    path = root / "resources.yaml"
    save(
        path,
        "resources",
        tables={
            "vars": {"path": "data/variables.xlsx", "header": True},
            "table": {"path": "data/tables.xlsx"},
        },
        images={"logo": {"path": "assets/logo.png"}},
    )
    return path


def template(root, resources=None):
    save(root / "styles/body.yaml", "text_style", font_family="휴먼명조", font_size=12)
    save(
        root / "sections/body.yaml",
        "section",
        layer={
            "type": "single_column",
            "contents": [
                {"type": "paragraph", "text": "<ref:month>"},
                {"type": "table", "resource": "table"},
                {"type": "container", "contents": [{"type": "image", "resource": "logo"}]},
            ],
        },
    )
    save(
        root / "document.yaml",
        "document",
        defaults={"text_style": {"src": "styles/body.yaml"}},
        resources=resources or {"src": "missing-default.yaml"},
        variables={"resource": "vars"},
        sections=[{"src": "sections/body.yaml"}],
    )
    return root


def test_override_builds_two_reports_without_copying_template(tmp_path, monkeypatch):
    root = template(tmp_path / "template")
    paths = [materials(tmp_path / "reports" / name, name) for name in ("8월 자료", "9월 자료")]
    monkeypatch.chdir(tmp_path / "reports")
    original = (root / "document.yaml").read_bytes()
    for path in paths:
        # Python API file selection is document-root relative, independent of cwd.
        selected = Path("..") / "reports" / path.parent.name / path.name
        document = load_project(root, resources=selected)
        assert document.root == root
        assert dict(document.variables) == {"month": path.parent.name}
        assert document.sections[0].contents[0].runs[0].style.font_family == "휴먼명조"
        assert document.sections[0].contents[1].data.cells[0].paragraph.runs[0].text == path.parent.name
        assert path in document.inputs
        assert path.parent / "data/tables.xlsx" in document.inputs
        result = build(root, path.parent / "report.hwpx", resources=selected)
        with ZipFile(result.output) as archive:
            assert path.parent.name in archive.read("Contents/section0.xml").decode()
    assert (root / "document.yaml").read_bytes() == original
    assert not (root / "missing-default.yaml").exists()
    assert build(root, "relative-output.hwpx", resources=paths[0]).output == root / "relative-output.hwpx"


def test_yaml_file_path_uses_resource_directory_and_absolute_shared_data(tmp_path):
    path = materials(tmp_path / "runs/august", "8월")
    shared = tmp_path / "runs/shared.xlsx"
    shared.write_bytes((path.parent / "data/tables.xlsx").read_bytes())
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    data["tables"]["table"]["path"] = str(shared)
    path.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    root = template(tmp_path / "template", {"file_path": "../runs/august/resources.yaml"})
    document = load_project(root)
    assert dict(document.variables) == {"month": "8월"}
    assert shared in document.inputs
    assert document.sections[0].source.path == root / "sections/body.yaml"


@pytest.mark.parametrize("category,name", [("tables", "table"), ("images", "logo")])
def test_file_path_relative_assets_cannot_escape_resource_directory(tmp_path, category, name):
    path = materials(tmp_path / "runs/august", "8월")
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    data[category][name]["path"] = "../outside.xlsx"
    path.write_text(yaml.safe_dump(data), encoding="utf-8")
    root = template(tmp_path / "template", {"file_path": "../runs/august/resources.yaml"})
    with pytest.raises(BuildError, match="RESOURCE_OUTSIDE_ROOT"):
        load_project(root)


def test_file_path_relative_symlinks_cannot_escape_resource_directory(tmp_path):
    path = materials(tmp_path / "run", "8월")
    external = tmp_path / "outside.xlsx"
    external.write_bytes((path.parent / "data/tables.xlsx").read_bytes())
    link = path.parent / "data/link.xlsx"
    try:
        link.symlink_to(external)
    except OSError:
        pytest.skip("symlinks unavailable")
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    data["tables"]["table"]["path"] = "data/link.xlsx"
    path.write_text(yaml.safe_dump(data), encoding="utf-8")
    root = template(tmp_path / "template")
    with pytest.raises(BuildError, match="RESOURCE_OUTSIDE_ROOT"):
        load_project(root, resources=path)


def test_src_keeps_document_root_paths_and_confinement(tmp_path):
    root = tmp_path / "template"
    path = materials(root, "기본")
    config = root / "config/resources.yaml"
    config.parent.mkdir()
    config.write_bytes(path.read_bytes())
    template(root, {"src": "config/resources.yaml"})
    assert dict(load_project(root).variables) == {"month": "기본"}
    template(root, {"src": "../outside.yaml"})
    with pytest.raises(BuildError, match="DECLARATION_OUTSIDE_ROOT"):
        load_project(root)
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    data["tables"]["table"]["path"] = "../outside.xlsx"
    inline = {k: v for k, v in data.items() if k not in ("schema_version", "type")}
    save(root / "document.yaml", "document", resources=inline, sections=[])
    with pytest.raises(BuildError, match="RESOURCE_OUTSIDE_ROOT"):
        load_project(root)


@pytest.mark.parametrize("field", ["defaults", "sections"])
def test_override_does_not_allow_external_styles_or_sections(tmp_path, field):
    path = materials(tmp_path / "run", "8월")
    root = template(tmp_path / "template")
    data = yaml.safe_load((root / "document.yaml").read_text(encoding="utf-8"))
    if field == "defaults":
        data["defaults"]["text_style"] = {"src": "../run/style.yaml"}
    else:
        data["sections"] = [{"src": "../run/section.yaml"}]
    (root / "document.yaml").write_text(yaml.safe_dump(data), encoding="utf-8")
    with pytest.raises(BuildError, match="DECLARATION_OUTSIDE_ROOT"):
        load_project(root, resources=path)


def test_external_resources_errors_and_output_protection(tmp_path):
    path = materials(tmp_path / "run", "8월")
    root = template(tmp_path / "template", {"file_path": "../run/resources.yaml"})
    original = path.read_bytes()
    with pytest.raises(BuildError, match="OUTPUT_CONFLICT"):
        build(root, path)
    assert path.read_bytes() == original
    with pytest.raises(BuildError, match="OUTPUT_CONFLICT"):
        build(root, path.parent / "data/tables.xlsx")
    path.write_text('schema_version: "1.0"\ntype: resources\ntables: [\n', encoding="utf-8")
    with pytest.raises(BuildError) as caught:
        load_project(root)
    d = caught.value.diagnostic
    assert d.code == "INVALID_DECLARATION" and d.source_path == str(path)
    assert d.line == 4 and d.column == 1
    assert d.details["reference_chain"][0]["pointer"] == "/resources/file_path"
    path.unlink()
    with pytest.raises(BuildError) as caught:
        load_project(root)
    d = caught.value.diagnostic
    assert d.code == "DECLARATION_READ_FAILED" and d.source_path == str(root / "document.yaml")
    assert d.pointer == "/resources/file_path" and d.line and d.column


@pytest.mark.parametrize("extra", [{"tables": {}}, {"src": "old.yaml"}])
def test_file_path_cannot_mix_with_other_resource_fields(tmp_path, extra):
    root = template(tmp_path, {"file_path": "resources.yaml", **extra})
    with pytest.raises(BuildError, match="SRC_INLINE_CONFLICT"):
        load_project(root)


@pytest.mark.parametrize("value", [None, "", 2, True])
def test_invalid_file_path_reports_its_yaml_location(tmp_path, value):
    root = template(tmp_path, {"file_path": value})
    with pytest.raises(BuildError) as caught:
        load_project(root)
    diagnostic = caught.value.diagnostic
    assert diagnostic.code == "INVALID_VALUE"
    assert diagnostic.pointer == "/resources/file_path"
    assert diagnostic.line and diagnostic.column


def test_cli_build_and_validate_use_the_same_override(tmp_path, monkeypatch, capsys):
    template(tmp_path / "template")
    path = materials(tmp_path / "run", "8월")
    monkeypatch.chdir(tmp_path)
    selected = str(path.relative_to(tmp_path))
    assert main(["validate", "template", "--resources", selected]) == 0
    assert json.loads(capsys.readouterr().out)["variables"] == 1
    output = tmp_path / "report.hwpx"
    assert main(["build", "template", "--resources", selected, "--output", "report.hwpx"]) == 0
    assert json.loads(capsys.readouterr().out)["tables"] == 1
    assert output.exists()
