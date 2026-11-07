"""Public compilation API; validation precedes atomic publication."""

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path

from hwpx import HwpxDocument

from .diagnostics import BuildError, Diagnostic
from .package import canonicalize, validate_package
from .planner import plan
from .project import load_project
from .writer import Writer


@dataclass(frozen=True)
class BuildResult:
    output: Path
    diagnostics: tuple[Diagnostic, ...]
    tables: int
    images: int


def build(
    project_root: str | Path,
    output: str | Path = "output.hwpx",
    *,
    debug: bool = False,
    warnings_as_errors: bool = False,
    resources: str | Path | None = None,
) -> BuildResult:
    document = plan(load_project(project_root, resources=resources))
    output = Path(output)
    output = (output if output.is_absolute() else document.root / output).resolve()
    if output in document.inputs:
        raise BuildError("OUTPUT_CONFLICT", "입력 선언/리소스 파일을 출력으로 덮어쓸 수 없습니다.", output)
    if warnings_as_errors and document.diagnostics:
        d = document.diagnostics[0]
        raise BuildError(d.code, d.message, d.source_path, d.pointer)
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        with tempfile.TemporaryDirectory(dir=output.parent) as temporary:
            raw = Path(temporary) / "raw.hwpx"
            ready = Path(temporary) / "ready.hwpx"
            Writer(document, debug=debug).save(raw)
            canonicalize(raw, ready)
            stats = validate_package(ready)
            reopened = HwpxDocument.open(ready)
            if not reopened.sections:
                raise BuildError("INVALID_PACKAGE", "문서를 재열 수 없습니다.", ready)
            os.replace(ready, output)
    except BuildError:
        raise
    except (OSError, ValueError, KeyError, AttributeError, RuntimeError) as exc:
        raise BuildError("WRITE_FAILED", str(exc), output) from exc
    return BuildResult(output, document.diagnostics, stats["tables"], stats["images"])
