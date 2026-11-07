"""Extraction regression uses a generated document with fictional content only."""

from pathlib import Path
from zipfile import ZipFile

import pytest
from openpyxl import load_workbook

from hwpx_builder import BuildError, build
from scripts.extract_images import extract_images
from scripts.extract_tables import extract_tables

ROOT = Path(__file__).resolve().parents[1]


def test_extract_synthetic_tables_and_images(tmp_path):
    source = build(ROOT / "template", tmp_path / "source.hwpx").output
    before = source.read_bytes()
    target = tmp_path / "tables.xlsx"
    manifest = extract_tables(source, target)
    assert len(manifest["tables"]) == 4
    book = load_workbook(target, rich_text=True)
    assert book.sheetnames == ["table_01", "table_02", "table_03", "table_04"]
    assert str(book["table_01"]["A1"].value) == "농Shim 제 1공장 운영보고서"
    assert str(book["table_03"]["A2"].value) == "구름면"
    book.close()
    images = extract_images(source, tmp_path / "images")
    assert len(images["images"]) == 1
    with ZipFile(source) as archive:
        for item in images["images"]:
            assert (tmp_path / "images" / item["file"]).read_bytes() == archive.read(item["source_entry"])
    assert source.read_bytes() == before
    with pytest.raises(BuildError, match="OUTPUT_EXISTS"):
        extract_tables(source, target)
    with pytest.raises(BuildError, match="OUTPUT_CONFLICT"):
        extract_tables(source, source, overwrite=True)
