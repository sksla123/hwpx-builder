"""Extract embedded image originals without transforming their bytes."""

from __future__ import annotations

import hashlib
import json
from io import BytesIO
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from defusedxml import ElementTree as ET
from PIL import Image

from hwpx_builder.diagnostics import BuildError


def extract_images(source: Path | str, output: Path | str, *, overwrite: bool = False) -> dict:
    source, output = Path(source).resolve(), Path(output).resolve()
    with ZipFile(source) as archive:
        package = ET.fromstring(archive.read("Contents/content.hpf"))
        items = []
        for node in package.iter():
            if not node.tag.endswith("}item") or not node.get("media-type", "").startswith("image/"):
                continue
            href = node.get("href")
            name = PurePosixPath(href).name
            if name in {i["file"] for i in items} or name == "index.json":
                raise BuildError("DUPLICATE_RESOURCE", f"이미지 파일명 중복: {name}", source)
            data = archive.read(href)
            with Image.open(BytesIO(data)) as image:
                width, height = image.size
            items.append(
                {
                    "id": node.get("id"),
                    "file": name,
                    "source_entry": href,
                    "media_type": node.get("media-type"),
                    "width_px": width,
                    "height_px": height,
                    "sha256": hashlib.sha256(data).hexdigest(),
                    "data": data,
                }
            )
        paths = [output / i["file"] for i in items] + [output / "index.json"]
        if source in paths:
            raise BuildError("OUTPUT_CONFLICT", "원본을 덮어쓸 수 없습니다.", source)
        if not overwrite and any(p.exists() for p in paths):
            raise BuildError("OUTPUT_EXISTS", "이미지 출력이 이미 있습니다.", output)
        output.mkdir(parents=True, exist_ok=True)
        for item in items:
            (output / item["file"]).write_bytes(item.pop("data"))
        result = {"source": source.name, "images": items}
        (output / "index.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        return result
