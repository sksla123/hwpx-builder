from decimal import ROUND_HALF_UP, Decimal


def mm_to_hwpunit(value: float | str | Decimal) -> int:
    return int(
        (Decimal(str(value)) * Decimal(7200) / Decimal("25.4")).quantize(Decimal(1), rounding=ROUND_HALF_UP)
    )
