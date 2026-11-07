"""Repair generated XLSX namespaces and whitespace preservation metadata.

Repository maintenance only; never invoked by the document builder.
"""

import argparse
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def spreadsheet_xml(data: bytes) -> bytes:
    """Preserve text runs, including whitespace-only runs rejected by Mac Excel."""
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    root = etree.fromstring(data, parser)
    if etree.QName(root).namespace != MAIN:
        return data
    changed = False
    if root.prefix is not None:
        namespaces = {None: MAIN, "r": RELATIONSHIPS}
        namespaces.update({k: v for k, v in root.nsmap.items() if v not in (MAIN, RELATIONSHIPS)})
        replacement = etree.Element(root.tag, nsmap=namespaces)
        replacement.attrib.update(root.attrib)
        replacement.text = root.text
        replacement.tail = root.tail
        replacement.extend(root)
        etree.cleanup_namespaces(replacement)
        root = replacement
        changed = True
    for text in root.iter(f"{{{MAIN}}}t"):
        if text.text and text.text != text.text.strip() and text.get(SPACE) != "preserve":
            # openpyxl's whitespace() misses strings consisting entirely of whitespace.
            text.set(SPACE, "preserve")
            changed = True
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True) if changed else data


def repair_workbook(path: Path) -> int:
    with ZipFile(path) as archive:
        parts = [(info, archive.read(info.filename)) for info in archive.infolist()]
        comment = archive.comment
    rewritten = []
    changed = 0
    for info, data in parts:
        fixed = spreadsheet_xml(data) if info.filename.endswith(".xml") else data
        changed += fixed != data
        rewritten.append((info, fixed))
    if not changed:
        return 0
    temporary = path.with_suffix(".repair.tmp.xlsx")
    try:
        with ZipFile(temporary, "w") as archive:
            archive.comment = comment
            for info, data in rewritten:
                archive.writestr(info, data)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)
    return changed


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbooks", nargs="+", type=Path)
    for path in parser.parse_args().workbooks:
        print(f"{path}: {repair_workbook(path)} XML parts repaired")
