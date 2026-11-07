"""Command-line entry point for declaration compilation."""

import argparse
import json
import logging
import sys
from pathlib import Path
from zipfile import BadZipFile

from .compiler import build
from .diagnostics import BuildError
from .planner import plan
from .project import load_project


def main(argv: list[str] | None = None) -> int:
    # JSON and Korean diagnostics use UTF-8 even when redirected on Windows.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")

    class TemplateFallbacks(logging.Filter):
        def filter(self, record):
            return not any(
                record.getMessage().startswith(prefix)
                for prefix in ("manifest에서 masterPage", "manifest에서 history", "manifest에서 version")
            )

    logging.getLogger("hwpx.opc.package").addFilter(TemplateFallbacks())
    parser = argparse.ArgumentParser(description="Declarative, Excel-backed HWPX builder")
    sub = parser.add_subparsers(dest="command", required=True)
    compilation = sub.add_parser("build", help="document.yaml에서 HWPX 생성")
    compilation.add_argument("project", type=Path)
    destination = compilation.add_mutually_exclusive_group()
    destination.add_argument(
        "--output", type=Path, default=Path("output.hwpx"), help="출력 파일 (문서 루트 기준)"
    )
    destination.add_argument(
        "--output-dir",
        type=Path,
        help="output.hwpx를 생성할 폴더 (현재 터미널 기준, --output과 함께 사용 불가)",
    )
    compilation.add_argument("--debug", action="store_true", help="<ref:...> 치환 구간을 노란색으로 강조")
    compilation.add_argument("--warnings-as-errors", action="store_true")
    validation = sub.add_parser("validate", help="선언·리소스·흐름 배치 검증")
    validation.add_argument("project", type=Path)
    for command in (compilation, validation):
        command.add_argument(
            "--resources",
            type=Path,
            metavar="FILE_PATH",
            help="resources 전체를 파일로 교체 (상대 경로: document.yaml 디렉터리)",
        )
    args = parser.parse_args(argv)
    try:
        if args.command == "build":
            result = build(
                args.project,
                args.output_dir.resolve() / "output.hwpx" if args.output_dir is not None else args.output,
                debug=args.debug,
                warnings_as_errors=args.warnings_as_errors,
                resources=args.resources,
            )
            print(
                json.dumps(
                    {
                        "output": str(result.output),
                        "tables": result.tables,
                        "images": result.images,
                        "debug": args.debug,
                        "diagnostics": [d.as_dict() for d in result.diagnostics],
                    },
                    ensure_ascii=False,
                )
            )
            return 0
        if args.command == "validate":
            document = plan(load_project(args.project, resources=args.resources))
            print(
                json.dumps(
                    {
                        "valid": True,
                        "sections": len(document.sections),
                        "variables": len(document.variables),
                        "diagnostics": [d.as_dict() for d in document.diagnostics],
                    },
                    ensure_ascii=False,
                )
            )
            return 0
    except BuildError as exc:
        print(json.dumps(exc.diagnostic.as_dict(), ensure_ascii=False), file=sys.stderr)
        return 1
    except (OSError, BadZipFile, KeyError, ValueError) as exc:
        error = BuildError(
            "BUILD_FAILED",
            str(exc),
            args.project,
        )
        print(json.dumps(error.diagnostic.as_dict(), ensure_ascii=False), file=sys.stderr)
        return 1
    return 0
