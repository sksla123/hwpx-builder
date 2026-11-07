"""Deterministic ZIP serialization and static HWPX consistency checks."""

from pathlib import Path, PurePosixPath
from xml.etree.ElementTree import ParseError
from zipfile import ZIP_DEFLATED, ZIP_STORED, BadZipFile, ZipFile, ZipInfo

from defusedxml import ElementTree as SafeET
from defusedxml.common import DefusedXmlException
from lxml import etree as ET

from .diagnostics import BuildError
from .writer import HH, HP


def canonicalize(source: Path, output: Path) -> None:
    with ZipFile(source) as archive, ZipFile(output, "w") as result:
        preview = []
        for name in archive.namelist():
            if name.startswith("Contents/section") and name.endswith(".xml"):
                section = SafeET.fromstring(archive.read(name))
                for paragraph in section.iter(HP + "p"):
                    value = "".join(
                        "".join(t.itertext())
                        for r in paragraph.findall(HP + "run")
                        for t in r.findall(HP + "t")
                    )
                    if value:
                        preview.append(value)
        for name in sorted(archive.namelist(), key=lambda p: (p != "mimetype", p)):
            if name == "Preview/PrvImage.png":
                continue  # The template's blank thumbnail does not represent the generated document.
            data = archive.read(name)
            if name == "Preview/PrvText.txt":
                data = ("\r\n".join(preview) + "\r\n").encode("utf-8")
            if name.endswith((".xml", ".hpf")):
                root = ET.fromstring(
                    data, ET.XMLParser(resolve_entities=False, no_network=True, load_dtd=False)
                )
                for i, node in enumerate(root.iter()):
                    if node.tag in {HP + "p", HP + "tbl", HP + "pic", HP + "rect", HP + "container"}:
                        for attr in ("id", "instid"):
                            if attr in node.attrib:
                                node.set(attr, str(i + 1))
                    if node.tag.endswith("}meta") and node.get("name") in (
                        "CreatedDate",
                        "ModifiedDate",
                        "date",
                    ):
                        node.text = "2000-01-01T00:00:00Z"
                data = ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
            entry = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            entry.create_system = 0
            entry.compress_type = ZIP_STORED if name == "mimetype" else ZIP_DEFLATED
            result.writestr(entry, data)


def validate_package(path: str | Path) -> dict:
    path = Path(path)

    def fail(message):
        raise BuildError("INVALID_PACKAGE", message, path)

    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            required = {
                "mimetype",
                "version.xml",
                "Contents/header.xml",
                "Contents/content.hpf",
                "Contents/section0.xml",
                "META-INF/container.xml",
            }
            if not required.issubset(names):
                fail(f"필수 ZIP 엔트리 누락: {sorted(required - set(names))}")
            if len(set(names)) != len(names):
                fail("중복 ZIP 엔트리가 있습니다.")
            if names[0] != "mimetype" or archive.getinfo("mimetype").compress_type != ZIP_STORED:
                fail("mimetype은 첫 번째 비압축 엔트리여야 합니다.")
            if archive.read("mimetype") != b"application/hwp+zip":
                fail("올바르지 않은 mimetype입니다.")
            if archive.testzip() is not None:
                fail("ZIP CRC 오류입니다.")
            for name in names:
                if PurePosixPath(name).is_absolute() or ".." in PurePosixPath(name).parts:
                    fail("ZIP 내부 경로 이탈입니다.")
            parsed = {n: SafeET.fromstring(archive.read(n)) for n in names if n.endswith((".xml", ".hpf"))}
            header = parsed["Contents/header.xml"]
            chars = {n.get("id") for n in header.iter(HH + "charPr")}
            paras = {n.get("id") for n in header.iter(HH + "paraPr")}
            fills = {n.get("id") for n in header.iter(HH + "borderFill")}
            binary_ids = set()
            for node in parsed["Contents/content.hpf"].iter():
                if not node.tag.endswith("}item"):
                    continue
                href = node.get("href")
                if href not in names and "Contents/" + str(href) not in names:
                    fail(f"manifest 대상 누락: {href}")
                if node.get("media-type", "").startswith("image/"):
                    binary_ids.add(node.get("id"))
            for name, root in parsed.items():
                for node in root.iter():
                    for attr, ids in [
                        ("charPrIDRef", chars),
                        ("paraPrIDRef", paras),
                        ("borderFillIDRef", fills),
                        ("binaryItemIDRef", binary_ids),
                    ]:
                        if (
                            attr in node.attrib
                            and node.get(attr) not in ids
                            and not (
                                name == "Contents/header.xml"
                                and node.tag == HH + "paraHead"
                                and attr == "charPrIDRef"
                                and node.get(attr) == "4294967295"
                            )
                        ):
                            fail(f"깨진 {attr}: {node.get(attr)} ({name})")
                    if node.tag == HP + "tbl":
                        rows, cols = int(node.get("rowCnt")), int(node.get("colCnt"))
                        covered = set()
                        for cell in node.findall(HP + "tr/" + HP + "tc"):
                            address = cell.find(HP + "cellAddr")
                            span = cell.find(HP + "cellSpan")
                            r, c = int(address.get("rowAddr")), int(address.get("colAddr"))
                            rs, cs = int(span.get("rowSpan")), int(span.get("colSpan"))
                            if min(rs, cs) < 1 or min(r, c) < 0 or r + rs > rows or c + cs > cols:
                                fail("표 병합 범위 오류입니다.")
                            for rr in range(r, r + rs):
                                for cc in range(c, c + cs):
                                    if (rr, cc) in covered:
                                        fail("표 셀이 겹칩니다.")
                                    covered.add((rr, cc))
                        if len(covered) != rows * cols:
                            fail("표 격자에 빈 영역이 있습니다.")
            return {
                "entries": len(names),
                "tables": sum(len(list(r.iter(HP + "tbl"))) for r in parsed.values()),
                "images": len(binary_ids),
            }
    except BuildError:
        raise
    except (BadZipFile, OSError, ValueError, AttributeError, ParseError, DefusedXmlException) as exc:
        fail(str(exc))
