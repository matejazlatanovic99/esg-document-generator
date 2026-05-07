"""Deterministic bad-data / invalid-field corruption engine.

Usage
-----
1.  Obtain a ``BadDataConfig`` from ``get_bad_data_config(raw_config)``.
2.  Call ``corrupt_records(records, artifact_key, field_types, cfg)`` to
    corrupt a flat list of record dicts in-place (returns a new list).
3.  For nested structures (e.g. statements containing transactions) use the
    lower-level ``corrupt_payload(payload, ...)`` on each leaf dict.

The corruption is *deterministic*: the same ``random_seed + artifact_key``
always produces the same set of corrupted fields and the same corrupted
values.  Field selection uses ``ceil(rate * eligible_count)`` unique picks.

Corrupted values are always *strings*.  Renderers should check whether a
value that is normally numeric/date is now a ``str``, and if so render it
as plain text without applying a number/date format.
"""
from __future__ import annotations

import copy
import math
import random
import string
from dataclasses import dataclass, field


# ── public field-type constants ──────────────────────────────────────────────

FTYPE_TEXT = "text"
FTYPE_IDENTIFIER = "identifier"
FTYPE_NUMERIC = "numeric"
FTYPE_DATE_TIME = "date_time"
FTYPE_CURRENCY_UNIT = "currency_unit"


# ── config ────────────────────────────────────────────────────────────────────

@dataclass
class BadDataConfig:
    enabled: bool = False
    preset: str = "mixed"      # "light" | "mixed" | "aggressive"
    rate: float = 0.15         # 0.0 – 1.0
    random_seed: int = 42


def get_bad_data_config(raw_config: dict) -> BadDataConfig:
    """Extract invalid-data settings from a generator raw_config dict."""
    inv = raw_config.get("document", {}).get("invalid_data", {})
    return BadDataConfig(
        enabled=bool(inv.get("enabled", False)),
        preset=str(inv.get("preset", "mixed")),
        rate=float(inv.get("rate", 0.15)),
        random_seed=int(raw_config.get("random_seed", 42)),
    )


# ── internal helpers ─────────────────────────────────────────────────────────

_ALPHA = string.ascii_letters
_ALPHANUM = string.ascii_letters + string.digits
_DIGITS = string.digits


def _rand_char(rng: random.Random, pool: str) -> str:
    return rng.choice(pool)


def _swap_char(s: str, rng: random.Random) -> str:
    """Swap one random character for another from the same class."""
    if not s:
        return s
    idx = rng.randrange(len(s))
    ch = s[idx]
    if ch.isdigit():
        new_ch = rng.choice([c for c in _DIGITS if c != ch] or list(_DIGITS))
    elif ch.isalpha():
        new_ch = rng.choice([c for c in _ALPHA if c.lower() != ch.lower()] or list(_ALPHA))
    else:
        new_ch = rng.choice(_ALPHANUM)
    return s[:idx] + new_ch + s[idx + 1:]


def _drop_char(s: str, rng: random.Random) -> str:
    if len(s) <= 1:
        return s
    idx = rng.randrange(len(s))
    return s[:idx] + s[idx + 1:]


def _insert_char(s: str, rng: random.Random, pool: str = _ALPHANUM) -> str:
    idx = rng.randrange(len(s) + 1)
    return s[:idx] + rng.choice(pool) + s[idx:]


def _truncate(s: str, rng: random.Random, min_keep: float = 0.4) -> str:
    if len(s) <= 2:
        return s
    keep = max(1, int(len(s) * rng.uniform(min_keep, 0.75)))
    return s[:keep]


def _random_alphanum(rng: random.Random, length: int = 6) -> str:
    return "".join(rng.choice(_ALPHANUM) for _ in range(max(1, length)))


def _damage_delimiter(s: str, rng: random.Random) -> str:
    """Replace or remove a delimiter character (-/.) to damage IDs/dates."""
    delimiters = [c for c in s if c in "-/. :"]
    if not delimiters:
        return _swap_char(s, rng)
    target = rng.choice(delimiters)
    idx = s.index(target)
    action = rng.choice(["remove", "swap"])
    if action == "remove":
        return s[:idx] + s[idx + 1:]
    replacement = rng.choice([c for c in "-/. :" if c != target])
    return s[:idx] + replacement + s[idx + 1:]


def _alpha_in_number(s: str, rng: random.Random) -> str:
    """Insert a letter inside a numeric string."""
    digit_positions = [i for i, c in enumerate(s) if c.isdigit()]
    if not digit_positions:
        return _insert_char(s, rng, _ALPHA)
    idx = rng.choice(digit_positions)
    return s[:idx] + rng.choice(_ALPHA) + s[idx + 1:]


def _corrupt_currency_code(s: str, rng: random.Random) -> str:
    """Corrupt a currency code (e.g. 'EUR' → 'EUX' or 'EBR')."""
    if not s:
        return s
    idx = rng.randrange(len(s))
    ch = s[idx]
    new_ch = rng.choice([c for c in _ALPHA.upper() if c != ch.upper()])
    return s[:idx] + new_ch + s[idx + 1:]


# ── per-type, per-preset corruption rules ────────────────────────────────────

def _corrupt_text_light(s: str, rng: random.Random) -> str:
    action = rng.choice(["swap", "truncate"])
    if action == "swap":
        return _swap_char(s, rng)
    return _truncate(s, rng)


def _corrupt_text_mixed(s: str, rng: random.Random) -> str:
    action = rng.choice(["swap", "truncate", "insert", "replace_short"])
    if action == "swap":
        return _swap_char(s, rng)
    if action == "truncate":
        return _truncate(s, rng)
    if action == "insert":
        return _insert_char(s, rng)
    return _random_alphanum(rng, rng.randint(3, 8))


def _corrupt_text_aggressive(s: str, rng: random.Random) -> str:
    action = rng.choice(["swap", "truncate", "insert", "replace_short", "full_replace"])
    if action == "full_replace":
        return _random_alphanum(rng, rng.randint(4, 12))
    return _corrupt_text_mixed(s, rng)


def _corrupt_identifier_light(s: str, rng: random.Random) -> str:
    action = rng.choice(["truncate", "damage_delimiter", "swap"])
    if action == "truncate":
        return _truncate(s, rng, min_keep=0.5)
    if action == "damage_delimiter":
        return _damage_delimiter(s, rng)
    return _swap_char(s, rng)


def _corrupt_identifier_mixed(s: str, rng: random.Random) -> str:
    action = rng.choice(["truncate", "damage_delimiter", "swap", "alpha_substitution"])
    if action == "alpha_substitution":
        # Replace a digit with a letter or vice-versa
        return _swap_char(s, rng)
    return _corrupt_identifier_light(s, rng)


def _corrupt_identifier_aggressive(s: str, rng: random.Random) -> str:
    action = rng.choice(["truncate", "damage_delimiter", "swap", "full_replace"])
    if action == "full_replace":
        return _random_alphanum(rng, max(3, len(s) // 2))
    return _corrupt_identifier_mixed(s, rng)


def _as_display_str(value) -> str:
    """Convert any record value to its 'display' string representation."""
    from decimal import Decimal
    from datetime import date, datetime
    if isinstance(value, str):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return str(value)
    return str(value)


def _corrupt_numeric_light(s: str, rng: random.Random) -> str:
    action = rng.choice(["drop_digit", "damage_separator"])
    if action == "drop_digit":
        return _drop_char(s, rng)
    return _damage_delimiter(s, rng)


def _corrupt_numeric_mixed(s: str, rng: random.Random) -> str:
    action = rng.choice(["drop_digit", "damage_separator", "alpha_in_number", "insert_digit"])
    if action == "alpha_in_number":
        return _alpha_in_number(s, rng)
    if action == "insert_digit":
        return _insert_char(s, rng, _DIGITS)
    return _corrupt_numeric_light(s, rng)


def _corrupt_numeric_aggressive(s: str, rng: random.Random) -> str:
    action = rng.choice(["drop_digit", "damage_separator", "alpha_in_number", "insert_digit", "full_text"])
    if action == "full_text":
        return _random_alphanum(rng, rng.randint(3, 8))
    return _corrupt_numeric_mixed(s, rng)


def _corrupt_date_time_light(s: str, rng: random.Random) -> str:
    action = rng.choice(["damage_delimiter", "truncate"])
    if action == "damage_delimiter":
        return _damage_delimiter(s, rng)
    return _truncate(s, rng, min_keep=0.5)


def _corrupt_date_time_mixed(s: str, rng: random.Random) -> str:
    action = rng.choice(["damage_delimiter", "truncate", "invalid_day_month"])
    if action == "invalid_day_month":
        # Force day to 32 or month to 13 by swapping leading digit
        parts = s.split("-") if "-" in s else s.split("/")
        if len(parts) == 3:
            # e.g. "2026-01-15" → pick day or month part, replace leading digit
            which = rng.choice([1, 2])  # month or day
            part = list(parts[which])
            if part:
                part[0] = rng.choice(["3", "4", "9"])
                parts[which] = "".join(part)
            sep = "-" if "-" in s else "/"
            return sep.join(parts)
    return _corrupt_date_time_light(s, rng)


def _corrupt_date_time_aggressive(s: str, rng: random.Random) -> str:
    action = rng.choice(["damage_delimiter", "truncate", "invalid_day_month", "full_replace"])
    if action == "full_replace":
        return _random_alphanum(rng, rng.randint(4, 10))
    return _corrupt_date_time_mixed(s, rng)


def _corrupt_currency_unit_light(s: str, rng: random.Random) -> str:
    action = rng.choice(["truncate", "swap"])
    if action == "truncate":
        return _truncate(s, rng, min_keep=0.5)
    return _swap_char(s, rng)


def _corrupt_currency_unit_mixed(s: str, rng: random.Random) -> str:
    action = rng.choice(["truncate", "swap", "code_corrupt"])
    if action == "code_corrupt":
        return _corrupt_currency_code(s, rng)
    return _corrupt_currency_unit_light(s, rng)


def _corrupt_currency_unit_aggressive(s: str, rng: random.Random) -> str:
    action = rng.choice(["truncate", "swap", "code_corrupt", "mixed_letters"])
    if action == "mixed_letters":
        return "".join(rng.choice(_ALPHANUM) for _ in s) if s else s
    return _corrupt_currency_unit_mixed(s, rng)


# ── dispatcher ────────────────────────────────────────────────────────────────

_CORRUPT_FNS: dict[str, dict[str, object]] = {
    FTYPE_TEXT: {
        "light": _corrupt_text_light,
        "mixed": _corrupt_text_mixed,
        "aggressive": _corrupt_text_aggressive,
    },
    FTYPE_IDENTIFIER: {
        "light": _corrupt_identifier_light,
        "mixed": _corrupt_identifier_mixed,
        "aggressive": _corrupt_identifier_aggressive,
    },
    FTYPE_NUMERIC: {
        "light": _corrupt_numeric_light,
        "mixed": _corrupt_numeric_mixed,
        "aggressive": _corrupt_numeric_aggressive,
    },
    FTYPE_DATE_TIME: {
        "light": _corrupt_date_time_light,
        "mixed": _corrupt_date_time_mixed,
        "aggressive": _corrupt_date_time_aggressive,
    },
    FTYPE_CURRENCY_UNIT: {
        "light": _corrupt_currency_unit_light,
        "mixed": _corrupt_currency_unit_mixed,
        "aggressive": _corrupt_currency_unit_aggressive,
    },
}


def _corrupt_one(value, field_type: str, preset: str, rng: random.Random) -> str:
    """Corrupt a single value and return a corrupted *string*."""
    s = _as_display_str(value)
    if not s:
        return s  # Never blank a field; skip empty values
    fn_map = _CORRUPT_FNS.get(field_type, _CORRUPT_FNS[FTYPE_TEXT])
    fn = fn_map.get(preset, fn_map.get("mixed"))
    result = fn(s, rng)
    # Ensure the result is at least minimally different and non-empty
    if not result or result == s:
        result = _swap_char(s, rng) if s else _random_alphanum(rng, 4)
    if not result:
        result = _random_alphanum(rng, 4)
    return result


# ── public API ────────────────────────────────────────────────────────────────

def corrupt_payload(
    payload: dict,
    artifact_key: str,
    field_types: dict[str, str],
    cfg: BadDataConfig,
) -> dict:
    """Return a new dict with ~``cfg.rate`` fraction of eligible fields corrupted.

    Parameters
    ----------
    payload:
        A flat dict of display-ready field values.
    artifact_key:
        A stable string that seeds the RNG together with ``cfg.random_seed``.
        Use something like ``"fuel_invoice:company1:site2:record3"``.
    field_types:
        Mapping of field name → FTYPE_* constant.  Keys absent from this
        mapping are skipped.
    cfg:
        ``BadDataConfig`` instance.
    """
    if not cfg.enabled:
        return payload

    rng = random.Random(f"{cfg.random_seed}:{artifact_key}")

    # Eligible fields: present in field_types map, non-empty, non-None value
    eligible = [
        k for k in field_types
        if k in payload and payload[k] not in (None, "", [])
    ]

    n_corrupt = math.ceil(cfg.rate * len(eligible))
    if n_corrupt <= 0 or not eligible:
        return payload

    n_corrupt = min(n_corrupt, len(eligible))
    to_corrupt = set(rng.sample(eligible, n_corrupt))

    result = dict(payload)
    for field_name in to_corrupt:
        ftype = field_types[field_name]
        value = payload[field_name]
        # Use a per-field sub-seed for stable individual corruption
        field_rng = random.Random(f"{cfg.random_seed}:{artifact_key}:{field_name}")
        result[field_name] = _corrupt_one(value, ftype, cfg.preset, field_rng)

    return result


def corrupt_records(
    records: list[dict],
    artifact_key: str,
    field_types: dict[str, str],
    cfg: BadDataConfig,
) -> list[dict]:
    """Apply ``corrupt_payload`` to each record in *records* independently.

    Each record receives an independent corruption pass; the corruption
    percentage target applies per-record (not across all records combined).
    Returns a new list of new record dicts.
    """
    if not cfg.enabled:
        return records
    return [
        corrupt_payload(rec, f"{artifact_key}:r{idx}", field_types, cfg)
        for idx, rec in enumerate(records)
    ]
