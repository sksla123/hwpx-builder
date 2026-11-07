"""A bounded stack scanner, deliberately independent of XML and report business logic."""

import re
from dataclasses import replace
from decimal import Decimal, InvalidOperation

from .model import Run, Source, TextStyle

COLORS = {
    "black": "#000000",
    "white": "#FFFFFF",
    "red": "#FF0000",
    "green": "#008000",
    "blue": "#0000FF",
    "yellow": "#FFFF00",
    "gray": "#808080",
}
DEFAULT_SOURCE = Source()


def color(value: str, source: Source = DEFAULT_SOURCE) -> str:
    if not isinstance(value, str):
        source.fail("INVALID_VALUE", "색상은 이름 또는 #RRGGBB 문자열이어야 합니다.")
    if value in COLORS:
        return COLORS[value]
    if re.fullmatch(r"#[0-9a-fA-F]{6}", value):
        return value.upper()
    source.fail("INVALID_VALUE", f"올바르지 않은 색상: {value}")


def parse_text(
    text: str, style: TextStyle, variables: dict[str, str], source: Source = DEFAULT_SOURCE
) -> tuple[Run, ...]:
    if not isinstance(text, str):
        source.fail("INVALID_VALUE", "text는 문자열이어야 합니다.")
    if len(text) > 1024 * 1024:
        source.fail("TEXT_LIMIT", "텍스트 길이 제한을 넘었습니다.")
    runs, stack, current, i = [], [], style, 0
    while i < len(text):
        start = i
        if text[i] != "<":
            end = text.find("<", i)
            if end < 0:
                end = len(text)
            value = text[i:end].replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
            runs.append(Run(value, current, start=i, end=end))
            i = end
            continue
        if text.startswith("<nat>", i):
            end = text.find("<\\nat>", i + 5)
            if end < 0:
                source.fail("INVALID_TEXT_TAG", "nat 종료 태그가 없습니다.")
            runs.append(Run(text[i + 5 : end], current, start=i, end=end + 6))
            i = end + 6
            continue
        end = text.find(">", i)
        if end < 0:
            source.fail("INVALID_TEXT_TAG", "태그가 닫히지 않았습니다.")
        token = text[i + 1 : end]
        i = end + 1
        if token == "linebreak":
            runs.append(Run("", current, True, start, i))
            continue
        if token.startswith("ref:"):
            key = token[4:]
            if key not in variables:
                source.fail("UNRESOLVED_REFERENCE", f"변수가 없습니다: {key}")
            runs.append(Run(variables[key], current, start=start, end=i, reference=key))
            continue
        if token.startswith("/"):
            if not stack or stack[-1][0] != token[1:]:
                source.fail("INVALID_TEXT_TAG", f"맞지 않는 종료 태그: {token}")
            _, current = stack.pop()
            continue
        name, sep, param = token.partition(":")
        if name not in ("bold", "italic", "underline", "strike", "size", "color", "highlight"):
            source.fail("UNKNOWN_TEXT_TAG", f"알 수 없는 태그: {token}")
        if (name in ("size", "color", "highlight")) != bool(sep):
            source.fail("INVALID_TEXT_TAG", f"태그 인수가 올바르지 않습니다: {token}")
        if len(stack) >= 128:
            source.fail("TEXT_LIMIT", "태그 중첩 제한을 넘었습니다.")
        stack.append((name, current))
        if name == "size":
            try:
                size = Decimal(param)
            except InvalidOperation:
                source.fail("INVALID_VALUE", f"올바르지 않은 크기: {param}")
            if not size.is_finite() or size <= 0:
                source.fail("INVALID_VALUE", "글자 크기는 양수여야 합니다.")
            current = replace(current, font_size=size)
        elif name in ("color", "highlight"):
            current = replace(current, **{name: color(param, source)})
        else:
            current = replace(current, **{name: True})
    if stack:
        source.fail("INVALID_TEXT_TAG", "닫히지 않은 서식 태그가 있습니다.")
    return tuple(runs) or (Run("", style),)


def literal_runs(text: str, style: TextStyle) -> tuple[Run, ...]:
    """Excel LF means an explicit cell line break; Excel text never enters the tag parser."""
    out = []
    for i, part in enumerate(text.split("\n")):
        if i:
            out.append(Run("", style, linebreak=True))
        out.append(Run(part, style))
    return tuple(out)
