from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP

CURRENCY_DISPLAY: dict[str, str] = {
    "GBP": "GBP (£)",
    "EUR": "EUR (€)",
    "USD": "USD ($)",
    "JPY": "JPY (¥)",
    "DKK": "DKK (kr)",
    "HUF": "HUF (Ft)",
}


def currency_options() -> list[str]:
    return list(CURRENCY_DISPLAY.values())


def currency_display(value: str | None, fallback_code: str = "EUR") -> str:
    if value in CURRENCY_DISPLAY:
        return CURRENCY_DISPLAY[str(value)]
    if value in CURRENCY_DISPLAY.values():
        return str(value)

    code = str(value or "").split(" ", 1)[0].upper()
    if code in CURRENCY_DISPLAY:
        return CURRENCY_DISPLAY[code]
    return CURRENCY_DISPLAY.get(fallback_code, CURRENCY_DISPLAY["EUR"])


def currency_index(value: str | None, fallback_code: str = "EUR") -> int:
    options = currency_options()
    selected = currency_display(value, fallback_code=fallback_code)
    return options.index(selected) if selected in options else 0


def currency_code(value: str | None, fallback_code: str = "GBP") -> str:
    display = currency_display(value, fallback_code=fallback_code)
    return display.split(" ", 1)[0]


def currency_symbol(value: str | None) -> str:
    display = currency_display(value, fallback_code="GBP")
    mapping = {
        "(£)": "£",
        "(€)": "€",
        "($)": "$",
        "(¥)": "¥",
        "(kr)": "kr",
        "(Ft)": "Ft",
    }
    for token, symbol in mapping.items():
        if token in display:
            return symbol
    return ""


def format_money(value, currency: str | None) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    symbol = currency_symbol(currency)
    amount = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return f"{symbol}{amount:,.2f}"


def replace_pound_labels(strings: dict[str, str], currency: str | None) -> dict[str, str]:
    code = currency_code(currency)
    if code == "GBP":
        return strings
    return {key: value.replace("£", code) for key, value in strings.items()}
