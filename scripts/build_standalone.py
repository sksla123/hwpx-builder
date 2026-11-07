"""Build and exercise a native CLI archive using only explicit public inputs."""

import json
import os
import platform
import shutil
import stat
import subprocess
import sys
import tempfile
import tomllib
from importlib.metadata import distribution
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from packaging.requirements import Requirement

ROOT = Path(__file__).resolve().parents[1]


def notices():
    pending = ["hwpx-builder", "pyinstaller"]
    seen = set()
    chunks = []
    while pending:
        name = pending.pop()
        dist = distribution(name)
        name = dist.metadata["Name"]
        if name.lower() in seen:
            continue
        seen.add(name.lower())
        chunks.append(
            f"\n{name} {dist.version}\nLicense: {dist.metadata.get('License-Expression') or dist.metadata.get('License', 'See license text')}\n"
        )
        for item in dist.files or ():
            if "license" in item.name.lower() or "copying" in item.name.lower():
                path = Path(dist.locate_file(item))
                if path.is_file():
                    chunks.append(path.read_text(encoding="utf-8", errors="replace"))
        for raw in dist.requires or ():
            requirement = Requirement(raw)
            if requirement.marker is None or requirement.marker.evaluate():
                pending.append(requirement.name)
    return "\n".join(chunks)


def run(executable, arguments, cwd):
    completed = subprocess.run(
        [str(executable), *arguments], cwd=cwd, check=True, capture_output=True, encoding="utf-8"
    )
    return completed.stdout.strip()


def smoke_test(executable, cwd, version):
    from hwpx_builder.package import validate_package

    assert run(executable, ["--version"], cwd) == f"hwpx-builder {version}"
    assert "{build,validate}" in run(executable, ["--help"], cwd)
    # Test inputs are created after extraction; they are never shipped in the ZIP.
    shutil.copytree(ROOT / "template", cwd / "example", ignore=shutil.ignore_patterns(".DS_Store"))
    import yaml

    config = cwd / "example/resources.yaml"
    data = yaml.safe_load(config.read_text(encoding="utf-8"))
    for key in ("table_01", "table_02"):
        data["tables"][key]["path"] = "data/cover.xlsx"
    config.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
    resources = "example/resources.yaml"
    validated = json.loads(run(executable, ["validate", "example", "--resources", resources], cwd))
    assert validated["valid"] and validated["variables"] == 6
    result = json.loads(
        run(
            executable,
            ["build", "example", "--resources", resources, "--output", "result/report.hwpx", "--debug"],
            cwd,
        )
    )
    output = cwd / "result/report.hwpx"
    assert Path(result["output"]) == output.resolve()
    assert result["tables"] == 4 and result["images"] == 1
    validate_package(output)
    with ZipFile(output) as archive:
        assert b"#FFFF00" in archive.read("Contents/header.xml")
    shutil.rmtree(cwd / "result")


def build_standalone():
    version = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    target = {("Darwin", "arm64"): "macos-arm64", ("Windows", "AMD64"): "windows-x64"}.get(
        (platform.system(), platform.machine())
    )
    if target is None:
        raise RuntimeError("Build on macOS arm64 or Windows x64")
    output = ROOT / "dist/release"
    output.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as scratch:
        scratch = Path(scratch)
        entry = scratch / "entry.py"
        entry.write_text("from hwpx_builder.cli import main\nraise SystemExit(main())\n", encoding="utf-8")
        subprocess.run(
            [
                sys.executable,
                "-m",
                "PyInstaller",
                "--clean",
                "--noconfirm",
                "--onefile",
                "--console",
                "--noupx",
                "--name",
                "hwpx-builder",
                "--distpath",
                str(scratch / "bin"),
                "--workpath",
                str(scratch / "build"),
                "--specpath",
                str(scratch),
                "--copy-metadata",
                "hwpx-builder",
                "--collect-data",
                "hwpx",
                "--hidden-import",
                "hwpx.data",
                str(entry),
            ],
            cwd=scratch,
            check=True,
        )
        bundle = scratch / "bundle"
        bundle.mkdir()
        name = "hwpx-builder.exe" if os.name == "nt" else "hwpx-builder"
        executable = bundle / name
        shutil.copy2(scratch / "bin" / name, executable)
        shutil.copy2(ROOT / "LICENSE", bundle / "LICENSE")
        (bundle / "THIRD_PARTY_NOTICES.txt").write_text(notices(), encoding="utf-8")
        archive = output / f"hwpx-builder-{version}-{target}.zip"
        with ZipFile(archive, "w", ZIP_DEFLATED) as zipped:
            for path in sorted(bundle.rglob("*")):
                if path.is_file():
                    zipped.write(path, path.relative_to(bundle).as_posix())
        # Test the actual archive from a separate directory, including non-ASCII paths.
        unpacked = scratch / "unpacked example 한글"
        with ZipFile(archive) as zipped:
            assert set(zipped.namelist()) == {name, "LICENSE", "THIRD_PARTY_NOTICES.txt"}
            zipped.extractall(unpacked)
        executable = unpacked / name
        if os.name != "nt":
            executable.chmod(executable.stat().st_mode | stat.S_IXUSR)
        smoke_test(executable, unpacked, version)
        print(f"Verified native archive: {archive}")


if __name__ == "__main__":
    build_standalone()
