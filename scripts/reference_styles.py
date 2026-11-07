"""HWPX reference style extraction and Excel presentation-note authoring."""

import json
from decimal import Decimal

from openpyxl.comments import Comment

from hwpx_builder.presentation import PREFIX, mm

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HH = "{http://www.hancom.co.kr/hwpml/2011/head}"
HC = "{http://www.hancom.co.kr/hwpml/2011/core}"
SIDES = ("left", "right", "top", "bottom")


def margins(node):
    return tuple(mm(node.get(k, "0")) if node is not None else Decimal(0) for k in SIDES)


def paragraph_style(node):
    """Normalize the default HWPX paragraph branch, rather than retaining XML."""
    align = node.find(HH + "align")
    default = node.find(".//" + HP + "default")
    branch = default if default is not None else node
    margin = branch.find(".//" + HH + "margin")
    spacing = branch.find(".//" + HH + "lineSpacing")
    breaking = node.find(HH + "breakSetting")
    result = {"align": align.get("horizontal", "LEFT").lower() if align is not None else "left"}
    for source, target in (
        ("left", "indent"),
        ("right", "right_indent"),
        ("intent", "first_line_indent"),
        ("prev", "space_before"),
        ("next", "space_after"),
    ):
        child = margin.find(HC + source) if margin is not None else None
        result[target] = str(mm(child.get("value", "0"))) if child is not None else "0"
    if spacing is not None and spacing.get("type") != "PERCENT":
        raise ValueError("Only percentage paragraph line spacing is supported")
    result["line_spacing"] = int(spacing.get("value", "160")) if spacing is not None else 160
    for key, attr in (("keep_with_next", "keepWithNext"), ("keep_together", "keepLines")):
        result[key] = breaking is not None and breaking.get(attr) == "1"
    return result


def write_note(cell, data):
    cell.comment = Comment(
        PREFIX + json.dumps(data, ensure_ascii=False, separators=(",", ":")), "hwpx-builder"
    )
