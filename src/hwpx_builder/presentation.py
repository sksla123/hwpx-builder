"""Portable presentation hints in standard Excel cell notes.

Native Excel values and formatting always win. Notes store only the presentation
Excel cannot express, never a second copy of cell text or HWPX XML.
"""

import hashlib
import json
from dataclasses import fields, replace
from decimal import Decimal

from .model import Paragraph, ParagraphStyle, Run

PREFIX = "hwpx-builder:presentation:1\n"
SIDES = ("left", "right", "top", "bottom")


def mm(value):
    return Decimal(str(value)) * Decimal("25.4") / 7200


def decode_style(data):
    allowed = {f.name for f in fields(ParagraphStyle)}
    if not isinstance(data, dict) or set(data) - allowed:
        raise ValueError("Invalid paragraph presentation hint")
    changes = dict(data)
    for key in ("indent", "right_indent", "first_line_indent", "space_before", "space_after"):
        value = Decimal(changes.get(key, 0))
        if not value.is_finite() or value < 0:
            raise ValueError("Invalid paragraph margin")
        changes[key] = value
    if changes.get("align", "left") not in ("left", "right", "center", "justify", "distribute"):
        raise ValueError("Invalid paragraph alignment")
    spacing = changes.get("line_spacing", 160)
    if type(spacing) is not int or not 1 <= spacing <= 10000:
        raise ValueError("Invalid line spacing")
    for key in ("keep_with_next", "keep_together"):
        if key in changes and type(changes[key]) is not bool:
            raise ValueError("Invalid paragraph break setting")
    return ParagraphStyle(**changes)


def digest(text):
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def border_signature(side):
    if side is None:
        return [None, None]
    c = side.color
    return [side.style, [c.type, c.value, c.tint] if c is not None else None]


def read_note(cell):
    if cell.comment is None or not cell.comment.text.startswith("hwpx-builder:presentation:"):
        return {}
    if not cell.comment.text.startswith(PREFIX):
        raise ValueError("Unsupported presentation note version")
    data = json.loads(cell.comment.text[len(PREFIX) :])
    if not isinstance(data, dict):
        raise TypeError("Invalid cell presentation note")
    return data


def split_paragraphs(paragraph, hint):
    """Split native rich text only when unchanged; never inject stored text."""
    specs = hint.get("paragraphs", [])
    if not specs:
        return ()
    # Preserve an intentional soft break versus a paragraph boundary.
    text = "".join("\n" if r.linebreak else r.text for r in paragraph.runs)
    if digest(text) != hint.get("text_sha256"):
        return ()
    lengths = [s["length"] for s in specs]
    if any(type(n) is not int or n < 0 for n in lengths) or sum(lengths) + len(lengths) - 1 != len(text):
        raise ValueError("Invalid paragraph boundaries")
    result, start = [], 0
    for spec in specs:
        end = start + spec["length"]
        runs, offset = [], 0
        for run in paragraph.runs:
            value = "\n" if run.linebreak else run.text
            lo, hi = max(start, offset), min(end, offset + len(value))
            if hi > lo:
                runs.append(replace(run, text="" if run.linebreak else value[lo - offset : hi - offset]))
            offset += len(value)
        if not runs:
            runs = [Run("", paragraph.runs[0].style)]
        style = decode_style(spec["style"])
        # The Excel alignment remains editable; retain per-paragraph alignment
        # only if the cell-level alignment has not changed since export.
        if paragraph.style.align != hint.get("align"):
            style = replace(style, align=paragraph.style.align)
        result.append(Paragraph(tuple(runs), style, paragraph.source))
        start = end + 1
    return tuple(result)
