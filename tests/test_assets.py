"""Fixed assets and replaceable materials share one resource format, with isolated roots."""

import json
from zipfile import ZipFile

import pytest
import yaml

from hwpx_builder import BuildError, build, load_project
from hwpx_builder.cli import main
from tests.test_resource_override import materials, template


def setup_assets(tmp_path):
    root = template(tmp_path / "template")
    static = materials(tmp_path / "fixed", "공통 표지")
    data = yaml.safe_load(static.read_text(encoding="utf-8"))
    data["tables"] = {"cover": data["tables"]["table"]}
    data["images"] = {"brand": data["images"]["logo"]}
    static.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
    document = yaml.safe_load((root / "document.yaml").read_text(encoding="utf-8"))
    document["assets"] = {"file_path": "../fixed/resources.yaml"}
    (root / "document.yaml").write_text(yaml.safe_dump(document, allow_unicode=True), encoding="utf-8")
    section = root / "sections/body.yaml"
    data = yaml.safe_load(section.read_text(encoding="utf-8"))
    data["layer"]["contents"].insert(0, {"type": "table", "resource": "cover"})
    data["layer"]["contents"].append(
        {"type": "container", "contents": [{"type": "image", "resource": "brand"}]}
    )
    section.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    report = materials(tmp_path / "run", "보고서 입력")
    return root, static, report


def test_override_retains_assets_and_each_file_resolves_its_own_data(tmp_path, monkeypatch, capsys):
    root, static, report = setup_assets(tmp_path)
    result = build(root, tmp_path / "report.hwpx", resources=report)
    assert result.tables == 2 and result.images == 2
    with ZipFile(result.output) as archive:
        text = archive.read("Contents/section0.xml").decode()
        assert "공통 표지" in text and "보고서 입력" in text
    document = load_project(root, resources=report)
    assert static in document.inputs and report in document.inputs
    assert (static.parent / "data/tables.xlsx") in document.inputs
    assert (report.parent / "data/tables.xlsx") in document.inputs
    monkeypatch.chdir(tmp_path)
    assert main(["validate", "template", "--resources", "run/resources.yaml"]) == 0
    assert json.loads(capsys.readouterr().out)["valid"]


@pytest.mark.parametrize("category,name", [("tables", "cover"), ("images", "brand")])
def test_duplicate_names_fail_at_replacing_file_without_silent_override(tmp_path, category, name):
    root, static, report = setup_assets(tmp_path)
    data = yaml.safe_load(report.read_text(encoding="utf-8"))
    entry = next(iter(data[category].values()))
    data[category][name] = dict(entry)
    report.write_text(yaml.safe_dump(data, allow_unicode=True), encoding="utf-8")
    with pytest.raises(BuildError, match="DUPLICATE_RESOURCE") as caught:
        load_project(root, resources=report)
    error = caught.value.diagnostic
    assert error.source_path == str(report)
    assert error.pointer == f"/{category}/{name}"
    assert error.line and error.column
    assert str(static) in error.message


@pytest.mark.parametrize("path", ["../run/data/tables.xlsx", "../run/assets/logo.png"])
def test_asset_relative_escape_is_rejected_even_when_other_source_contains_target(tmp_path, path):
    root, static, report = setup_assets(tmp_path)
    data = yaml.safe_load(static.read_text(encoding="utf-8"))
    data["tables"]["cover"]["path"] = path
    static.write_text(yaml.safe_dump(data), encoding="utf-8")
    with pytest.raises(BuildError, match="RESOURCE_OUTSIDE_ROOT"):
        load_project(root, resources=report)


@pytest.mark.parametrize("target", ["declaration", "unused-data"])
def test_assets_are_protected_from_output_overwrite(tmp_path, target):
    root, static, report = setup_assets(tmp_path)
    output = static
    if target == "unused-data":
        output = static.parent / "data/unused.xlsx"
        output.write_bytes((static.parent / "data/tables.xlsx").read_bytes())
        config = yaml.safe_load(static.read_text(encoding="utf-8"))
        config["tables"]["unused"] = {"path": "data/unused.xlsx"}
        static.write_text(yaml.safe_dump(config), encoding="utf-8")
    previous = output.read_bytes()
    with pytest.raises(BuildError, match="OUTPUT_CONFLICT"):
        build(root, output, resources=report)
    assert output.read_bytes() == previous


def test_inline_assets_share_resource_format_and_categories_are_independent(tmp_path):
    root, static, report = setup_assets(tmp_path)
    data = yaml.safe_load((root / "document.yaml").read_text(encoding="utf-8"))
    data["assets"] = {
        "tables": {"cover": {"path": str(static.parent / "data/tables.xlsx")}},
        "images": {"brand": {"path": str(static.parent / "assets/logo.png")}},
    }
    data["assets"]["images"]["cover"] = dict(data["assets"]["images"]["brand"])
    (root / "document.yaml").write_text(yaml.safe_dump(data), encoding="utf-8")
    assert len(load_project(root, resources=report).sections) == 1


def test_assets_file_diagnostics_and_missing_file_reference(tmp_path):
    root, static, report = setup_assets(tmp_path)
    static.write_text("schema_version: '1.0'\ntype: resources\ntables: [\n", encoding="utf-8")
    with pytest.raises(BuildError) as caught:
        load_project(root, resources=report)
    assert caught.value.diagnostic.source_path == str(static)
    assert caught.value.diagnostic.line
    static.unlink()
    with pytest.raises(BuildError) as caught:
        load_project(root, resources=report)
    assert caught.value.diagnostic.pointer == "/assets/file_path"


def test_asset_file_path_cannot_mix_inline_and_src(tmp_path):
    root, _static, report = setup_assets(tmp_path)
    data = yaml.safe_load((root / "document.yaml").read_text(encoding="utf-8"))
    data["assets"]["images"] = {}
    (root / "document.yaml").write_text(yaml.safe_dump(data), encoding="utf-8")
    with pytest.raises(BuildError, match="SRC_INLINE_CONFLICT"):
        load_project(root, resources=report)


def test_asset_symlink_cannot_escape_its_source_directory(tmp_path):
    root, static, report = setup_assets(tmp_path)
    link = static.parent / "data/link.xlsx"
    try:
        link.symlink_to(report.parent / "data/tables.xlsx")
    except OSError:
        pytest.skip("symlink creation is unavailable")
    data = yaml.safe_load(static.read_text(encoding="utf-8"))
    data["tables"]["cover"]["path"] = "data/link.xlsx"
    static.write_text(yaml.safe_dump(data), encoding="utf-8")
    with pytest.raises(BuildError, match="RESOURCE_OUTSIDE_ROOT"):
        load_project(root, resources=report)


def test_assets_alone_can_supply_all_resource_roles(tmp_path):
    root, static, report = setup_assets(tmp_path)
    config = yaml.safe_load((root / "document.yaml").read_text(encoding="utf-8"))
    config.pop("resources")
    (root / "document.yaml").write_text(yaml.safe_dump(config), encoding="utf-8")
    assets = yaml.safe_load(static.read_text(encoding="utf-8"))
    assets["tables"]["vars"] = {"path": str(report.parent / "data/variables.xlsx"), "header": True}
    assets["tables"]["table"] = dict(assets["tables"]["cover"])
    assets["images"]["logo"] = dict(assets["images"]["brand"])
    static.write_text(yaml.safe_dump(assets, allow_unicode=True), encoding="utf-8")
    assert dict(load_project(root).variables) == {"month": "보고서 입력"}
