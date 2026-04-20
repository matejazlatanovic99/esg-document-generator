from __future__ import annotations

from collections.abc import Sequence


COMPACT_FUEL_VOLUME_UNITS: tuple[str, ...] = ("L", "kL", "ML", "gal", "kgal")
LONG_FUEL_VOLUME_UNITS: tuple[str, ...] = (
    "Litres",
    "Kilolitres",
    "Megalitres",
    "Gallons",
    "Thousand gallons",
)

DEFAULT_COMPACT_FUEL_VOLUME_UNIT = COMPACT_FUEL_VOLUME_UNITS[0]
DEFAULT_LONG_FUEL_VOLUME_UNIT = LONG_FUEL_VOLUME_UNITS[0]

DOCUMENT_TYPE_FUEL_VOLUME_UNITS: dict[str, tuple[str, ...]] = {
    "fuel_invoice": LONG_FUEL_VOLUME_UNITS,
    "delivery_note": LONG_FUEL_VOLUME_UNITS,
    "generator_log": COMPACT_FUEL_VOLUME_UNITS,
    "fuel_card": COMPACT_FUEL_VOLUME_UNITS,
}

DOCUMENT_TYPE_DEFAULT_FUEL_VOLUME_UNIT: dict[str, str] = {
    "fuel_invoice": DEFAULT_LONG_FUEL_VOLUME_UNIT,
    "delivery_note": DEFAULT_LONG_FUEL_VOLUME_UNIT,
    "generator_log": DEFAULT_COMPACT_FUEL_VOLUME_UNIT,
    "fuel_card": DEFAULT_COMPACT_FUEL_VOLUME_UNIT,
}


def default_fuel_volume_unit(document_type: str | None) -> str:
    return DOCUMENT_TYPE_DEFAULT_FUEL_VOLUME_UNIT.get(
        document_type or "",
        DEFAULT_COMPACT_FUEL_VOLUME_UNIT,
    )


def option_index(options: Sequence[str], selected: str | None) -> int:
    if selected in options:
        return options.index(selected)
    return 0
