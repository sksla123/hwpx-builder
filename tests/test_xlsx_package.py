"""Input preparation must keep Excel's native SpreadsheetML namespace layout."""

from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from scripts.repair_xlsx import repair_workbook as prepare_workbook
from scripts.repair_xlsx import spreadsheet_xml

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ROOT = Path(__file__).resolve().parents[1]
SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def test_whitespace_only_rich_text_run_keeps_its_linebreak():
    data = (
        f'<worksheet xmlns="{MAIN}"><sheetData><row r="1"><c r="B1" t="inlineStr">'
        "<is><r><t>공사</t></r><r><t>\n</t></r><r><t>부서</t></r></is>"
        "</c></row></sheetData></worksheet>"
    ).encode()
    result = spreadsheet_xml(data)
    root = etree.fromstring(result)
    text = list(root.iter(f"{{{MAIN}}}t"))
    assert [t.text for t in text] == ["공사", "\n", "부서"]
    assert text[1].get(SPACE) == "preserve"
    assert spreadsheet_xml(result) == result


def legacy_workbook(path):
    # Reproduce the old serializer even after the committed fixture is repaired.
    with ZipFile(ROOT / "template/data/tables.xlsx") as source, ZipFile(path, "w") as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename.endswith(".xml"):
                root = etree.fromstring(data)
                if etree.QName(root).namespace == MAIN:
                    prefixed = etree.Element(root.tag, nsmap={"broken": MAIN})
                    prefixed.attrib.update(root.attrib)
                    prefixed.text = root.text
                    prefixed.extend(root)
                    data = etree.tostring(prefixed)
            target.writestr(info, data)


def test_preparation_uses_default_spreadsheet_namespace(tmp_path):
    path = tmp_path / "tables.xlsx"
    legacy_workbook(path)
    prepare_workbook(path)
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.endswith(".xml"):
                continue
            root = etree.fromstring(archive.read(name))
            if etree.QName(root).namespace == MAIN:
                assert root.prefix is None, name
                assert root.nsmap.get(None) == MAIN, name


def test_repair_preserves_all_xml_meaning_and_other_parts(tmp_path):
    from scripts.repair_xlsx import repair_workbook

    path = tmp_path / "tables.xlsx"
    legacy_workbook(path)
    with ZipFile(path) as archive:
        before = {name: archive.read(name) for name in archive.namelist()}
    assert repair_workbook(path) > 0
    with ZipFile(path) as archive:
        assert archive.testzip() is None
        assert set(archive.namelist()) == set(before)
        for name, data in before.items():
            after = archive.read(name)
            if data == after:
                continue
            from xml.etree.ElementTree import canonicalize

            expected = etree.fromstring(data)
            for text in expected.iter(f"{{{MAIN}}}t"):
                if text.text and text.text != text.text.strip():
                    text.set(SPACE, "preserve")
            assert canonicalize(etree.tostring(expected), rewrite_prefixes=True) == canonicalize(
                after, rewrite_prefixes=True
            )
    saved = path.read_bytes()
    assert repair_workbook(path) == 0
    assert path.read_bytes() == saved
