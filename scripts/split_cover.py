"""Split prepared report cover sheets without re-exporting cells, notes, or drawings."""

from pathlib import Path, PurePosixPath
from posixpath import dirname, join, normpath
from zipfile import ZipFile

import yaml
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]


def read(path):
    return yaml.safe_load(path.read_text(encoding="utf-8"))


def save(path, value):
    path.write_text(yaml.safe_dump(value, allow_unicode=True, sort_keys=False), encoding="utf-8")


MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
COVER = ("table_01", "table_02")


def xml(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def serialize(node):
    return etree.tostring(node, encoding="UTF-8", xml_declaration=True)


def target_part(owner, target):
    return target.lstrip("/") if target.startswith("/") else normpath(join(dirname(owner), target))


def select_sheets(source: Path, selected: tuple[str, ...]):
    """Retain selected sheets and their complete OPC dependency closure."""
    with ZipFile(source) as archive:
        infos = {info.filename: info for info in archive.infolist()}
        parts = {name: archive.read(name) for name in infos}
    workbook = xml(parts["xl/workbook.xml"])
    sheets = workbook.find(MAIN + "sheets")
    names = [s.get("name") for s in sheets]
    if not selected or not set(selected) <= set(names):
        raise ValueError(f"Missing or empty sheet selection: {selected}")
    kept_indices = {
        i: len([n for n in names[:i] if n in selected]) for i, n in enumerate(names) if n in selected
    }
    discarded = set()
    for sheet in list(sheets):
        if sheet.get("name") not in selected:
            discarded.add(sheet.get(REL + "id"))
            sheets.remove(sheet)
    for view in workbook.findall(MAIN + "bookViews/" + MAIN + "workbookView"):
        view.set("activeTab", "0")
        view.set("firstSheet", "0")
    for definitions in workbook.findall(MAIN + "definedNames"):
        for definition in list(definitions):
            index = definition.get("localSheetId")
            if index is None:
                continue
            if int(index) not in kept_indices:
                definitions.remove(definition)
            else:
                definition.set("localSheetId", str(kept_indices[int(index)]))
    relationships = xml(parts["xl/_rels/workbook.xml.rels"])
    for relation in list(relationships):
        if relation.get("Id") in discarded:
            relationships.remove(relation)
    parts["xl/workbook.xml"] = serialize(workbook)
    parts["xl/_rels/workbook.xml.rels"] = serialize(relationships)
    kept = {"[Content_Types].xml"}

    def visit(owner):
        if owner:
            if owner in kept:
                return
            if owner not in parts:
                raise ValueError(f"Missing XLSX relationship target: {owner}")
            kept.add(owner)
        item = PurePosixPath(owner)
        rel_path = str(item.parent / "_rels" / (item.name + ".rels")) if owner else "_rels/.rels"
        if rel_path in parts:
            kept.add(rel_path)
            for relation in xml(parts[rel_path]):
                if relation.get("TargetMode") != "External":
                    visit(target_part(owner, relation.get("Target")))

    visit("")
    types = xml(parts["[Content_Types].xml"])
    for node in list(types):
        if node.get("PartName") and node.get("PartName").lstrip("/") not in kept:
            types.remove(node)
    parts["[Content_Types].xml"] = serialize(types)
    return [(infos[name], parts[name]) for name in infos if name in kept]


def write_parts(path, parts):
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".split.tmp.xlsx")
    try:
        with ZipFile(temporary, "w") as archive:
            for info, data in parts:
                archive.writestr(info, data)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def sheet_names(path):
    with ZipFile(path) as archive:
        return [s.get("name") for s in xml(archive.read("xl/workbook.xml")).find(MAIN + "sheets")]


def split_report_covers():
    cover_path = ROOT / "template/data/cover.xlsx"
    for project in (ROOT / "template", ROOT / "examples/factory_report"):
        source = project / "data/tables.xlsx"
        names = sheet_names(source)
        if any(name in names for name in COVER):
            selected = select_sheets(source, COVER)
            if cover_path.exists():
                with ZipFile(cover_path) as archive:
                    current = {
                        n: archive.read(n) for n in archive.namelist() if not n.startswith("docProps/")
                    }
                candidate = {
                    info.filename: data
                    for info, data in selected
                    if not info.filename.startswith("docProps/")
                }
                if current != candidate:
                    raise ValueError("Cover parts differ; review before sharing the cover workbook")
            else:
                write_parts(cover_path, selected)
            write_parts(source, select_sheets(source, tuple(n for n in names if n not in COVER)))
        elif not cover_path.is_file():
            raise ValueError("Cover sheets and shared cover workbook are missing")
        reference = read(project / "document.yaml")["assets"]
        path = project / reference["file_path"]
        assets = read(path)
        for name in COVER:
            assets["tables"][name]["path"] = "data/cover.xlsx"
        save(path, assets)


if __name__ == "__main__":
    split_report_covers()
