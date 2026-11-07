from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Diagnostic:
    code: str
    message: str
    source_path: str = ""
    pointer: str = ""
    severity: str = "error"
    line: int | None = None
    column: int | None = None
    details: Any = None

    def as_dict(self) -> dict:
        return asdict(self)


class BuildError(Exception):
    def __init__(self, code: str, message: str, source: str | Path = "", pointer: str = "", **kw):
        self.diagnostic = Diagnostic(code, message, str(source), pointer, **kw)
        location = str(source)
        if kw.get("line") is not None:
            location += f":{kw['line']}"
            if kw.get("column") is not None:
                location += f":{kw['column']}"
        super().__init__(f"{code}: {location} {pointer}: {message}")
