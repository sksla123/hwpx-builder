"""Repository-only CLI for preparing reference inputs.

Run from the repository root: uv run python -m scripts.extract_reference --help
"""

import argparse
import json
import sys
from pathlib import Path
from zipfile import BadZipFile

from hwpx_builder.diagnostics import BuildError

from .extract_images import extract_images
from .extract_tables import extract_tables


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Prepare reference inputs for examples and tests")
    sub = parser.add_subparsers(dest="command", required=True)
    tables = sub.add_parser("tables", help="완성 HWPX 표와 문서 변수를 XLSX로 이관")
    tables.add_argument("source", type=Path)
    tables.add_argument("--output", type=Path, required=True)
    tables.add_argument("--variables-profile", type=Path)
    tables.add_argument("--variables-output", type=Path)
    tables.add_argument("--overwrite", action="store_true")
    images = sub.add_parser("images", help="원본 이미지 추출")
    images.add_argument("source", type=Path)
    images.add_argument("--output", type=Path, required=True)
    images.add_argument("--overwrite", action="store_true")
    args = parser.parse_args(argv)
    try:
        if args.command == "images":
            result = extract_images(args.source, args.output, overwrite=args.overwrite)
            summary = {"image_count": len(result["images"])}
        else:
            result = extract_tables(
                args.source,
                args.output,
                overwrite=args.overwrite,
                variables_profile=args.variables_profile,
                variables_output=args.variables_output,
            )
            summary = result["validation"]
    except BuildError as exc:
        print(json.dumps(exc.diagnostic.as_dict(), ensure_ascii=False), file=sys.stderr)
        return 1
    except (OSError, BadZipFile, KeyError, ValueError) as exc:
        error = BuildError("EXTRACTION_FAILED", str(exc), args.source)
        print(json.dumps(error.diagnostic.as_dict(), ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps(summary, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
