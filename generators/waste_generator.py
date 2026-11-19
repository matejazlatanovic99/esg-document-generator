"""Scope 3 Category 5 – Waste generated in operations.

Generates synthetic waste documents (WTN, contractor invoice / credit note, weighbridge ticket,
waste report, hazardous waste consignment note, skip / waste log) together with ground truth in
the ``Scope3WasteGroundTruth`` shape from the waste-management spike (``scope3Waste.types.ts``).

Every document is first built as a *record* (header fields + ground-truth lines, with display-only
values stored under ``_``-prefixed keys). Renderers read records; ground truth strips the private
keys. Header fields go through the shared bad-data engine before both, so ground truth always
matches what is printed.
"""
from __future__ import annotations

import csv
import json
import random
import re
from calendar import monthrange
from datetime import date, timedelta
from io import BytesIO, StringIO
from typing import Any

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl.styles import Font, PatternFill
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from generators import waste_reference as ref
from generators.stationary_combustion_generator import (
    _augment_stationary_field_ids as _augment_field_ids,
    _currency_symbol,
    _financial_period,
    _months_in_range,
    _multi_document_count,
    _ordered_stationary_field_ids as _ordered_field_ids,
    _set_docx_cell_text,
    _shade_docx_cell,
    _slugify_filename,
    _stationary_distractor_field_map as _distractor_field_map,
    _stationary_document_lines as _document_lines,
    _stationary_row_values as _row_values,
    _style_docx_document,
    _with_special_chars,
    _write_stationary_csv_preamble as _write_csv_preamble,
    _write_stationary_xlsx_preamble as _write_xlsx_preamble,
    _zip_documents,
)
from utils.bad_data import FTYPE_DATE_TIME, FTYPE_IDENTIFIER, FTYPE_TEXT, corrupt_records, get_bad_data_config
from utils.distractor_fields import resolve_distractor_plan
from utils.layouts import resolve_layout_plan

CATEGORY = ref.CATEGORY_WASTE

_HEADER_FIELD_TYPES = {
    "document_number": FTYPE_IDENTIFIER,
    "supplier_name": FTYPE_TEXT,
    "organisation": FTYPE_TEXT,
    "carrier_reg": FTYPE_IDENTIFIER,
    "document_date": FTYPE_DATE_TIME,
}

# ── options ──────────────────────────────────────────────────────────────────

DEFAULT_OPTIONS: dict[str, Any] = {
    "fill_factor": 1.0,
    # WTN
    "wtn_mode": "single",  # single | season_ticket
    "wtn_quantity_style": "approx",  # approx | container
    "wtn_rd_in_destination": True,
    # invoice
    "inv_role": "invoice",  # invoice | credit_note
    "inv_basis": "lift_priced",  # lift_priced | weighed | unweighed | mixed
    "inv_excess_weight": False,
    "inv_billing_period": "calendar_month",  # calendar_month | four_five_week
    "inv_show_routes": True,
    "inv_rebate": False,
    "inv_landfill_tax": False,
    "inv_non_waste_charges": True,
    "inv_unitless_weight": False,
    "inv_eu_format": False,
    "inv_no_waste_lines": False,
    "inv_two_month_period": False,
    "inv_weighbridge_refs": False,
    # weighbridge
    "wb_unit": "kg",  # kg | t
    "wb_mode": "gross_tare_net",  # gross_tare_net | stored_tare | gross_only | net_mismatch
    # report
    "rep_route_split": True,
    "rep_multi_site": False,
    "rep_totals_row": True,
    # HWCN
    "hwcn_part_e_differs": False,
    "hwcn_interim": False,
    "hwcn_non_hazardous_line": False,
    # skip / waste log
    "log_mixed_units": False,
    "log_volume_rows": True,
    "log_blank_rows": False,
    "log_total_row": False,
    "log_hazardous_row": False,
    "log_notes_sheet": False,
}


def _options(raw_config: dict) -> dict:
    return {**DEFAULT_OPTIONS, **(raw_config.get("document", {}).get("waste") or {})}


def _document_type(raw_config: dict) -> str:
    return str(raw_config.get("document_type") or raw_config.get("document", {}).get("type") or ref.DOC_INVOICE)


def _seed(raw_config: dict) -> int:
    return int(raw_config.get("random_seed", 42))


def _vat_rate(raw_config: dict) -> float:
    try:
        return float(raw_config.get("document", {}).get("vat_rate", 20) or 0)
    except (TypeError, ValueError):
        return 20.0


def _layout_context(raw_config: dict) -> dict:
    return {
        "language": raw_config.get("document", {}).get("language", "en"),
        "period_label": raw_config.get("financial_period", {}).get("label", ""),
        "company_labels": [c.get("label", f"Company {i + 1}") for i, c in enumerate(raw_config.get("companies", []))],
    }


def _layout_plan(raw_config: dict, output_format: str) -> dict:
    return resolve_layout_plan(
        raw_config.get("document", {}),
        random_seed=_seed(raw_config),
        category=CATEGORY,
        document_type=_document_type(raw_config),
        output_format=output_format,
        context=_layout_context(raw_config),
    )


def _distractor_plan(raw_config: dict, output_format: str):
    return resolve_distractor_plan(
        raw_config.get("document", {}),
        random_seed=_seed(raw_config),
        category=CATEGORY,
        document_type=_document_type(raw_config),
        output_format=output_format,
        context=_layout_context(raw_config),
    )


# ── company / site / stream helpers ──────────────────────────────────────────


def _company_site_pairs(raw_config: dict) -> list[tuple[dict, dict]]:
    pairs = []
    for company in raw_config.get("companies", []):
        for site in company.get("sites", []) or [{}]:
            pairs.append((company, site))
    return pairs or [({}, {})]


def _site_streams(site: dict, *, hazardous: bool | None = None) -> list[dict]:
    """Return [{key, preset, lifts_per_month}] for a site, optionally filtered by hazard status."""
    streams = []
    for entry in site.get("streams", []) or []:
        key = entry.get("preset") if isinstance(entry, dict) else entry
        preset = ref.STREAM_PRESETS.get(key)
        if not preset:
            continue
        is_haz = preset["ewc"].endswith("*")
        if hazardous is not None and is_haz != hazardous:
            continue
        lifts = entry.get("lifts_per_month", 4) if isinstance(entry, dict) else 4
        try:
            lifts = max(int(lifts), 1)
        except (TypeError, ValueError):
            lifts = 4
        streams.append({"key": key, "preset": preset, "lifts_per_month": lifts})
    return streams


def _fallback_streams(keys: list[str]) -> list[dict]:
    return [{"key": k, "preset": ref.STREAM_PRESETS[k], "lifts_per_month": 4} for k in keys]


def _site_address_text(site: dict) -> str:
    return ", ".join(line for line in site.get("address", []) if line)


def _contractor(company: dict) -> dict:
    contractor = company.get("contractor") or {}
    if contractor.get("name"):
        return {
            "name": contractor.get("name", ""),
            "code": contractor.get("code", "") or "WC",
            "carrier_reg": contractor.get("carrier_reg", ""),
            "address": list(contractor.get("address") or []),
        }
    return dict(ref.WASTE_CONTRACTORS[0])


def _currency(company: dict) -> tuple[str, str]:
    raw = str(company.get("currency") or "GBP (£)")
    code = raw.split()[0] if raw else "GBP"
    return code, _currency_symbol(raw) or ("£" if code == "GBP" else "€" if code == "EUR" else "")


def _month_bounds(year: int, month: int) -> tuple[date, date]:
    return date(year, month, 1), date(year, month, monthrange(year, month)[1])


def _add_months(year: int, month: int, delta: int) -> tuple[int, int]:
    index = year * 12 + (month - 1) + delta
    return index // 12, index % 12 + 1


def _doc_month(raw_config: dict, index: int) -> tuple[date, date]:
    fp = _financial_period(raw_config)
    months = _months_in_range(fp["start_date"], fp["end_date"]) or [(fp["start_date"].year, fp["start_date"].month)]
    year, month = months[index % len(months)]
    return _month_bounds(year, month)


def _random_day(rng: random.Random, start: date, end: date) -> date:
    return start + timedelta(days=rng.randint(0, max((end - start).days, 0)))


def _text(raw_config: dict, value: str) -> str:
    return _with_special_chars(raw_config, value)


# ── formatting ───────────────────────────────────────────────────────────────


def _fmt_number(value, decimals: int = 2, *, eu: bool = False, thousands: bool = True) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    text = f"{value:,.{decimals}f}" if thousands else f"{value:.{decimals}f}"
    if eu:
        text = text.replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return text


def _fmt_money(value, symbol: str, *, eu: bool = False) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    sign = "-" if value < 0 else ""
    body = _fmt_number(abs(value), 2, eu=eu)
    return f"{sign}{symbol} {body}" if eu else f"{sign}{symbol}{body}"


def _fmt_date(value, *, eu: bool = False) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    return value.strftime("%d-%m-%Y" if eu else "%d/%m/%Y")


def _iso(value) -> str | None:
    if value is None:
        return None
    return value.isoformat() if hasattr(value, "isoformat") else str(value)


def _r(value: float | None, digits: int = 4) -> float | None:
    return None if value is None else round(value, digits)


def _dedupe(items: list[str]) -> list[str]:
    out: list[str] = []
    for item in items:
        if item and item not in out:
            out.append(item)
    return out


# ── line derivation (single source of truth for tonnes / route / flags) ──────


def _resolve_route(preset: dict, material: str, evidence: dict) -> tuple[str | None, str | None, list[str], str | None]:
    """Apply the unified precedence: stated → R/D code → landfill tax → stream keyword → default.

    Returns (route, source, flags, treatment_code).
    """
    flags: list[str] = []
    rd_code = evidence.get("rd_code")
    if evidence.get("stated"):
        return evidence["stated"], ref.SRC_STATED, flags, rd_code
    interim = False
    if rd_code:
        route = ref.rd_code_route(rd_code, material, ref.preset_loop_route(preset))
        if route:
            return route, ref.SRC_RD_CODE, flags, rd_code
        interim = ref.rd_code_is_interim(rd_code)
    if interim:
        return None, None, [ref.FLAG_INTERIM], rd_code
    if evidence.get("landfill_tax"):
        return ref.ROUTE_LANDFILL, ref.SRC_LANDFILL_TAX, flags, rd_code
    if evidence.get("keyword"):
        return evidence["keyword"], ref.SRC_KEYWORD, flags, rd_code
    return ref.MATERIAL_DEFAULT_ROUTE.get(material, ref.ROUTE_COMBUSTION), ref.SRC_DEFAULT, [ref.FLAG_ROUTE_DEFAULTED], rd_code


def _quantity_tonnes(preset: dict, material: str, quantity, unit: str | None, *, lifts=None, fill: float = 1.0) -> float | None:
    if quantity is None or unit is None:
        return None
    if unit == ref.UNIT_KG:
        return quantity / 1000.0
    if unit == ref.UNIT_T:
        return float(quantity)
    density = ref.ewc_density(preset.get("ewc"), material)
    if density is None:
        return None
    if unit == ref.UNIT_YD3:
        return quantity * ref.CUBIC_YARD_TO_CUBIC_METRE * density * fill
    if unit == ref.UNIT_M3:
        return quantity * density * fill
    if unit == ref.UNIT_LIFT:
        _, size, size_unit = preset["container"]
        if not size or size_unit != ref.UNIT_L:
            return None
        return quantity * (size / 1000.0) * density * fill
    return None


def _waste_line(
    preset: dict,
    *,
    description: str,
    quantity,
    unit: str | None,
    basis: str | None,
    collection_date: date | None,
    evidence: dict,
    document_type: str,
    fill: float,
    line_kind: str = ref.LINE_WASTE,
    cost: float | None = None,
    lifts: int | None = None,
    show_ewc: bool = True,
    container_override: tuple | None = None,
    gross=None,
    tare=None,
    net=None,
    site: str | None = None,
    tonnes_override: float | None = None,
    omit_tonnes: bool = False,
    extra_flags: list[str] | None = None,
    display: dict | None = None,
) -> dict:
    material = ref.preset_material(preset)
    route, source, route_flags, treatment_code = _resolve_route(preset, material, evidence)
    ewc = preset["ewc"]
    hazardous = ref.ewc_is_hazardous(ewc)
    ctype, csize, cunit = container_override or preset["container"]

    if omit_tonnes:
        tonnes = None
    elif tonnes_override is not None:
        tonnes = tonnes_override
    else:
        tonnes = _quantity_tonnes(preset, material, quantity, unit, lifts=lifts, fill=fill)

    flags: list[str] = []
    if basis == ref.BASIS_ESTIMATED:
        flags.append(ref.FLAG_ESTIMATED)
    if basis == ref.BASIS_VOLUME:
        flags.append(ref.FLAG_VOLUME)
    if basis == ref.BASIS_LIFT:
        flags.append(ref.FLAG_LIFT)
    if isinstance(quantity, (int, float)) and quantity < 0:
        flags.append(ref.FLAG_NEGATIVE)
    if hazardous and document_type != ref.DOC_HWCN:
        flags.append(ref.FLAG_HAZ_NON_HWCN)
    flags += route_flags + list(extra_flags or [])

    line = {
        "lineKind": line_kind,
        "description": description,
        "ewcCode": ewc if show_ewc else None,
        "isHazardous": hazardous if show_ewc else (hazardous or None),
        "material": material,
        "treatmentCode": treatment_code,
        "treatmentRoute": route,
        "treatmentRouteSource": source,
        "quantity": quantity,
        "quantityUnit": unit,
        "grossWeight": gross,
        "tareWeight": tare,
        "netWeight": net,
        "containerType": ctype,
        "containerSize": csize,
        "containerSizeUnit": cunit if csize else None,
        "numberOfLifts": lifts,
        "collectionDate": _iso(collection_date),
        "cost": _r(cost, 2),
        "site": site,
        "quantityBasis": basis,
        "quantityTonnes": _r(tonnes),
        "expectedFlags": _dedupe(flags) or None,
    }
    if display:
        line["_display"] = display
    return line


def _plain_line(kind: str, description: str, *, cost=None, collection_date=None, quantity=None, unit=None, display=None) -> dict:
    line = {
        "lineKind": kind,
        "description": description,
        "cost": _r(cost, 2),
        "collectionDate": _iso(collection_date),
        "quantity": quantity,
        "quantityUnit": unit,
    }
    if display:
        line["_display"] = display
    return line


def _base_record(raw_config: dict, company: dict, site: dict, **fields) -> dict:
    code, symbol = _currency(company)
    record = {
        "document_id": "",
        "document_type": _document_type(raw_config),
        "document_role": ref.ROLE_ORIGINAL,
        "document_number": "",
        "related_document_number": None,
        "referenced_document_numbers": None,
        "supplier_name": "",
        "carrier_reg": None,
        "organisation": _text(raw_config, company.get("name") or company.get("label") or "Organisation"),
        "site_label": site.get("label", ""),
        "site_address": _site_address_text(site),
        "site_address_lines": list(site.get("address", [])),
        "sic_code": company.get("sic_code") or None,
        "document_date": None,
        "period_start": None,
        "period_end": None,
        "currency_code": code,
        "currency_symbol": symbol,
        "total_cost": None,
        "lines": [],
        "doc_flags": [],
        "special_cases": [],
        "notes": "",
    }
    record.update(fields)
    return record


# ── invoice / credit note ────────────────────────────────────────────────────

_SKIP_LIKE = {ref.CONTAINER_SKIP, ref.CONTAINER_COMPACTOR, ref.CONTAINER_RORO}


def _invoice_stream_lines(raw_config, rng, opts, preset, stream, period_start, period_end, *, month_factor, landfill_override, symbol, refs):
    """Lines for one stream on an invoice. Returns (waste_lines, tax_lines, special_cases)."""
    document_type = ref.DOC_INVOICE
    fill = float(opts["fill_factor"])
    eu = bool(opts["inv_eu_format"])
    ctype, csize, cunit = preset["container"]
    route = ref.ROUTE_LANDFILL if landfill_override else preset["route"]
    basis_mode = opts["inv_basis"]
    if basis_mode == "mixed":
        basis_mode = rng.choice(["lift_priced", "weighed", "unweighed"])
    always_weighed = ctype in {ref.CONTAINER_LOOSE, ref.CONTAINER_OTHER, ref.CONTAINER_BAG} or not csize
    if ctype in _SKIP_LIKE:
        # Skip / compactor operators weigh at the receiving site unless the job is priced as volume only.
        weighed = always_weighed or basis_mode in {"weighed", "lift_priced"}
    else:
        # Wheeled bins: real invoices charge per lift; total weight only with on-board weighing.
        weighed = always_weighed or basis_mode == "weighed"
    specials: list[str] = []

    landfill_tax = bool(opts["inv_landfill_tax"]) and route == ref.ROUTE_LANDFILL
    stated = route if (opts["inv_show_routes"] and not landfill_tax) else None
    evidence = {"stated": stated, "landfill_tax": landfill_tax, "keyword": preset.get("keyword_route")}
    route_suffix = f" – {ref.ROUTE_LABEL[route]}" if stated else ""
    if landfill_tax and stated is None:
        route_suffix = " – disposal"

    lifts_total = max(1, round(stream["lifts_per_month"] * month_factor + rng.randint(-1, 1)))
    lo, hi = preset["kg_per_lift"]
    waste_lines: list[dict] = []
    tax_lines: list[dict] = []

    if ctype in _SKIP_LIKE:
        exchanges = min(lifts_total, 3)
        for k in range(exchanges):
            when = _random_day(rng, period_start, period_end)
            kg = rng.randint(lo, hi)
            action = rng.choice(["exchange", "collection", "delivery & exchange"]) if k == 0 else rng.choice(["exchange", "collection"])
            desc = _text(raw_config, f"{preset['description']}{route_suffix} – {action}")
            flags = []
            if weighed:
                rate = round(rng.uniform(32, 48), 2) if ctype == ref.CONTAINER_COMPACTOR else None
                cost = round(kg / 1000 * rate, 2) if rate else round(rng.uniform(210, 300), 2)
                ref_no = None
                if opts["inv_weighbridge_refs"]:
                    ref_no = f"WB-{rng.randint(50000, 69999)}"
                    refs.append(ref_no)
                    flags.append(ref.FLAG_POSSIBLE_DUPLICATE)
                if opts["inv_unitless_weight"]:
                    flags.append(ref.FLAG_AMBIGUOUS_UNIT)
                line = _waste_line(
                    preset, description=desc, quantity=kg, unit=ref.UNIT_KG, basis=ref.BASIS_WEIGHED,
                    collection_date=when, evidence=evidence, document_type=document_type, fill=fill, cost=cost,
                    lifts=1, extra_flags=flags,
                    display={"date": when, "qty": "1", "weight": _fmt_number(kg, 0, eu=eu),
                             "rate": f"{_fmt_money(rate, symbol, eu=eu)}/t" if rate else _fmt_money(cost, symbol, eu=eu)},
                )
            else:
                cost = round(rng.uniform(210, 300), 2)
                specials.append("volume_only")
                line = _waste_line(
                    preset, description=desc, quantity=csize, unit=cunit, basis=ref.BASIS_VOLUME,
                    collection_date=when, evidence=evidence, document_type=document_type, fill=fill, cost=cost,
                    lifts=1, tonnes_override=(round(kg / 1000, 4) if landfill_tax else None),
                    display={"date": when, "qty": "1", "weight": "", "rate": _fmt_money(cost, symbol, eu=eu)},
                )
                if landfill_tax:
                    line["quantityBasis"] = ref.BASIS_TAX
                    line["expectedFlags"] = [f for f in (line["expectedFlags"] or []) if f != ref.FLAG_VOLUME] or None
            waste_lines.append(line)
            if landfill_tax:
                tonnes = round(kg / 1000, 2)
                rate = ref.landfill_tax_rate(when)
                tax_lines.append(_plain_line(
                    ref.LINE_LANDFILL_TAX, _text(raw_config, f"Landfill Tax {_fmt_number(tonnes, 2, eu=eu)} t @ {_fmt_money(rate, symbol, eu=eu)}"),
                    cost=round(tonnes * rate, 2), collection_date=when,
                    display={"date": when, "qty": "", "weight": "", "rate": ""},
                ))
        return waste_lines, tax_lines, specials

    when = period_end
    if ctype == ref.CONTAINER_LOOSE or not csize or ctype in {ref.CONTAINER_OTHER, ref.CONTAINER_BAG}:
        collections = 1 if stream["lifts_per_month"] < 3 else 2
        for _ in range(collections):
            when = _random_day(rng, period_start, period_end)
            tonnes = round(rng.randint(lo, hi) / 1000, 2)
            if preset.get("rebate") and opts["inv_rebate"]:
                rate = -round(rng.uniform(35, 60), 2)
                kind = ref.LINE_REBATE
                route_words = " – closed-loop recycling" if opts["inv_show_routes"] else ""
                desc = _text(raw_config, f"{preset['description']}{route_words} – REBATE")
                ev = {"stated": ref.ROUTE_CLOSED_LOOP if opts["inv_show_routes"] else None, "keyword": preset.get("keyword_route")}
                specials.append("rebate")
            else:
                rate = round(rng.uniform(8, 25), 2)
                kind = ref.LINE_WASTE
                desc = _text(raw_config, f"{preset['description']}{route_suffix}")
                ev = evidence
            cost = round(tonnes * rate, 2)
            flags = [ref.FLAG_AMBIGUOUS_UNIT] if opts["inv_unitless_weight"] else []
            waste_lines.append(_waste_line(
                preset, description=desc, quantity=round(tonnes * 1000), unit=ref.UNIT_KG, basis=ref.BASIS_WEIGHED,
                collection_date=when, evidence=ev, document_type=document_type, fill=fill, cost=cost, lifts=None,
                line_kind=kind, extra_flags=flags,
                display={"date": when, "qty": "1", "weight": _fmt_number(round(tonnes * 1000), 0, eu=eu),
                         "rate": f"{_fmt_money(rate, symbol, eu=eu)}/t"},
            ))
        return waste_lines, tax_lines, specials

    # wheeled bins – one aggregated line per stream
    rate = round(rng.uniform(8.5, 19.5), 2)
    cost = round(lifts_total * rate, 2)
    desc = _text(raw_config, f"{preset['description']}{route_suffix} ({lifts_total} lifts)")
    kg = sum(rng.randint(lo, hi) for _ in range(lifts_total))
    if landfill_tax:
        tonnes = round(kg / 1000, 2)
        tax_rate = ref.landfill_tax_rate(when)
        tax_lines.append(_plain_line(
            ref.LINE_LANDFILL_TAX, _text(raw_config, f"Landfill Tax {_fmt_number(tonnes, 2, eu=eu)} t @ {_fmt_money(tax_rate, symbol, eu=eu)}"),
            cost=round(tonnes * tax_rate, 2), collection_date=when,
            display={"date": when, "qty": "", "weight": "", "rate": ""},
        ))
    if weighed:
        flags = [ref.FLAG_AMBIGUOUS_UNIT] if opts["inv_unitless_weight"] else []
        waste_lines.append(_waste_line(
            preset, description=desc, quantity=kg, unit=ref.UNIT_KG, basis=ref.BASIS_WEIGHED, collection_date=when,
            evidence=evidence, document_type=document_type, fill=fill, cost=cost, lifts=lifts_total, extra_flags=flags,
            display={"date": when, "qty": str(lifts_total), "weight": _fmt_number(kg, 0, eu=eu), "rate": _fmt_money(rate, symbol, eu=eu)},
        ))
    else:
        specials.append("lifts_only")
        if opts["inv_excess_weight"] and not landfill_tax:
            allowance = int(hi * 0.8)
            excess = max(kg - allowance * lifts_total, rng.randint(40, 300))
            excess_rate = round(rng.uniform(120, 180), 2)
            tax_lines.append(_plain_line(
                ref.LINE_NON_WASTE,
                _text(raw_config, f"Excess weight – {preset['description']} ({_fmt_number(excess, 0, eu=eu)} kg over {allowance} kg/lift allowance)"),
                cost=round(excess / 1000 * excess_rate, 2), collection_date=when,
                display={"date": when, "qty": "", "weight": _fmt_number(excess, 0, eu=eu), "rate": f"{_fmt_money(excess_rate, symbol, eu=eu)}/t"},
            ))
            specials.append("excess_weight_line")
        line = _waste_line(
            preset, description=desc, quantity=lifts_total, unit=ref.UNIT_LIFT, basis=ref.BASIS_LIFT, collection_date=when,
            evidence=evidence, document_type=document_type, fill=fill, cost=cost, lifts=lifts_total,
            tonnes_override=(round(kg / 1000, 4) if landfill_tax else None),
            display={"date": when, "qty": str(lifts_total), "weight": "", "rate": _fmt_money(rate, symbol, eu=eu)},
        )
        if landfill_tax:
            line["quantityBasis"] = ref.BASIS_TAX
            line["expectedFlags"] = [f for f in (line["expectedFlags"] or []) if f != ref.FLAG_LIFT] or None
        waste_lines.append(line)
    return waste_lines, tax_lines, specials


def _build_invoices(raw_config: dict) -> list[dict]:
    opts = _options(raw_config)
    seed = _seed(raw_config)
    pairs = _company_site_pairs(raw_config)
    records = []
    for index in range(_multi_document_count(raw_config)):
        rng = random.Random(f"{seed}:waste_invoice:{index}")
        company, site = pairs[index % len(pairs)]
        contractor = _contractor(company)
        code, symbol = _currency(company)
        eu = bool(opts["inv_eu_format"])
        period_start, period_end = _doc_month(raw_config, index)
        month_factor = 1.0
        specials: list[str] = []
        if opts["inv_billing_period"] == "four_five_week":
            # Contractor billing periods of 4 or 5 weeks, starting on a Monday – not calendar months.
            period_start = period_start - timedelta(days=period_start.weekday())
            period_end = period_start + timedelta(days=rng.choice([27, 34]))
            month_factor = (period_end - period_start).days / 30.0
            specials.append("billing_period_not_calendar_month")
        if opts["inv_two_month_period"]:
            y, m = _add_months(period_end.year, period_end.month, 1)
            period_end = _month_bounds(y, m)[1] if opts["inv_billing_period"] == "calendar_month" else period_end + timedelta(days=28)
            month_factor *= 2.0
        doc_date = period_end + timedelta(days=rng.randint(2, 6))
        number = f"{contractor['code']}-INV-{rng.randint(240000, 249999)}"
        record = _base_record(
            raw_config, company, site,
            document_number=number,
            supplier_name=_text(raw_config, contractor["name"]),
            carrier_reg=contractor["carrier_reg"],
            document_date=doc_date,
            period_start=period_start,
            period_end=period_end,
            _contractor=contractor,
            _eu=eu,
            _vat_rate=_vat_rate(raw_config),
        )
        streams = _site_streams(site, hazardous=False) or _fallback_streams(["general_waste_1100", "dmr_1100"])
        if not opts["inv_rebate"]:
            streams = [s for s in streams if not s["preset"].get("rebate")] or streams
        elif not any(s["preset"].get("rebate") for s in streams):
            streams = streams + _fallback_streams(["cardboard_baled"])
        streams = streams[:4] if len(streams) > 4 else streams
        landfill_idx = None
        if opts["inv_landfill_tax"] and not any(s["preset"]["route"] == ref.ROUTE_LANDFILL for s in streams):
            landfill_idx = next((i for i, s in enumerate(streams) if ref.preset_material(s["preset"]) == ref.MAT_CI), 0)

        lines: list[dict] = []
        refs: list[str] = []
        if opts["inv_role"] == "credit_note":
            record["document_role"] = ref.ROLE_CREDIT_NOTE
            record["related_document_number"] = number
            record["document_number"] = f"{contractor['code']}-CN-{rng.randint(240000, 249999)}"
            stream = next((s for s in streams if s["preset"]["container"][0] == ref.CONTAINER_BIN), streams[0])
            preset = stream["preset"]
            lifts = -rng.randint(1, 3)
            rate = round(rng.uniform(8.5, 19.5), 2)
            cost = round(lifts * rate, 2)
            desc = _text(raw_config, f"Credit – {preset['description']} lifts charged in error ({abs(lifts)} lifts)")
            evidence = {"keyword": preset.get("keyword_route")}
            if opts["inv_basis"] == "weighed":  # on-board weighing → credited kg shown
                lo, hi = preset["kg_per_lift"]
                kg = -sum(rng.randint(lo, hi) for _ in range(abs(lifts)))
                line = _waste_line(preset, description=desc, quantity=kg, unit=ref.UNIT_KG, basis=ref.BASIS_WEIGHED,
                                   collection_date=period_end, evidence=evidence, document_type=ref.DOC_INVOICE,
                                   fill=float(opts["fill_factor"]), cost=cost, lifts=lifts,
                                   display={"date": period_end, "qty": str(lifts), "weight": _fmt_number(kg, 0, eu=eu), "rate": _fmt_money(rate, symbol, eu=eu)})
            else:
                line = _waste_line(preset, description=desc, quantity=lifts, unit=ref.UNIT_LIFT, basis=ref.BASIS_LIFT,
                                   collection_date=period_end, evidence=evidence, document_type=ref.DOC_INVOICE,
                                   fill=float(opts["fill_factor"]), cost=cost, lifts=lifts,
                                   display={"date": period_end, "qty": str(lifts), "weight": "", "rate": _fmt_money(rate, symbol, eu=eu)})
            lines.append(line)
            surcharge = round(cost * ref.FUEL_SURCHARGE_PCT / 100, 2)
            lines.append(_plain_line(ref.LINE_NON_WASTE, _text(raw_config, "Credit – fuel surcharge on above"), cost=surcharge,
                                     collection_date=period_end, display={"date": period_end, "qty": "", "weight": "", "rate": ""}))
            specials.append("credit_note")
            record["notes"] = "Credit note – waste quantities are negative; route should be inherited from the credited invoice line."
        elif opts["inv_no_waste_lines"]:
            for desc, cost in rng.sample(ref.NON_WASTE_CHARGES, 3):
                when = _random_day(rng, period_start, period_end)
                lines.append(_plain_line(ref.LINE_NON_WASTE, _text(raw_config, desc), cost=cost, collection_date=when,
                                         display={"date": when, "qty": "1", "weight": "", "rate": _fmt_money(cost, symbol, eu=eu)}))
            record["doc_flags"].append(ref.FLAG_NO_WASTE_LINES)
            specials += ["no_waste_lines", "non_waste_charges"]
            record["notes"] = "Negative test – waste contractor invoice with no waste lines; must produce no emissions record."
        else:
            for i, stream in enumerate(streams):
                waste, tax, sc = _invoice_stream_lines(
                    raw_config, rng, opts, stream["preset"], stream, period_start, period_end,
                    month_factor=month_factor, landfill_override=(i == landfill_idx), symbol=symbol, refs=refs,
                )
                lines += waste + tax
                specials += sc
                if any(l["lineKind"] == ref.LINE_LANDFILL_TAX for l in tax):
                    specials.append("landfill_tax_line")
            if opts["inv_non_waste_charges"]:
                for desc, cost in rng.sample(ref.NON_WASTE_CHARGES, rng.randint(2, 3)):
                    when = _random_day(rng, period_start, period_end)
                    lines.append(_plain_line(ref.LINE_NON_WASTE, _text(raw_config, desc), cost=cost, collection_date=when,
                                             display={"date": when, "qty": "1", "weight": "", "rate": _fmt_money(cost, symbol, eu=eu)}))
                waste_cost = sum(l["cost"] or 0 for l in lines if l["lineKind"] == ref.LINE_WASTE)
                surcharge = round(waste_cost * ref.FUEL_SURCHARGE_PCT / 100, 2)
                lines.append(_plain_line(ref.LINE_NON_WASTE, _text(raw_config, f"Fuel surcharge {ref.FUEL_SURCHARGE_PCT}%"), cost=surcharge,
                                         collection_date=period_end, display={"date": period_end, "qty": "", "weight": "", "rate": ""}))
                specials.append("non_waste_charges")
        if refs:
            record["referenced_document_numbers"] = refs
            specials.append("duplicate_across_docs")
        if opts["inv_unitless_weight"] and any(l.get("quantityUnit") == ref.UNIT_KG for l in lines):
            specials.append("ambiguous_unit")
        if eu:
            specials.append("european_decimal")
        if "excess_weight_line" in specials:
            record["notes"] = ("Bins are charged per lift; the 'Excess weight' line is only the kg above the agreed allowance – "
                               "it is a charge, not the collected quantity, and must not be used as tonnes.")
        subtotal = round(sum(l.get("cost") or 0 for l in lines), 2)
        vat = round(subtotal * record["_vat_rate"] / 100, 2)
        record.update(lines=lines, _subtotal=subtotal, _vat=vat, total_cost=round(subtotal + vat, 2), special_cases=_dedupe(specials))
        records.append(record)
    return records


# ── waste transfer note ──────────────────────────────────────────────────────


def _build_wtns(raw_config: dict) -> list[dict]:
    opts = _options(raw_config)
    seed = _seed(raw_config)
    fill = float(opts["fill_factor"])
    pairs = _company_site_pairs(raw_config)
    fp = _financial_period(raw_config)
    records = []
    for index in range(_multi_document_count(raw_config)):
        rng = random.Random(f"{seed}:waste_wtn:{index}")
        company, site = pairs[index % len(pairs)]
        contractor = _contractor(company)
        streams = _site_streams(site, hazardous=False) or _fallback_streams(["general_waste_1100"])
        season = opts["wtn_mode"] == "season_ticket"
        period_start, period_end = _doc_month(raw_config, index)
        transfer_date = _random_day(rng, period_start, period_end)
        specials: list[str] = []
        lines: list[dict] = []
        quantity_texts: list[str] = []
        destinations: list[str] = []
        if season:
            chosen = [s for s in streams if s["preset"]["container"][0] == ref.CONTAINER_BIN][:3] or streams[:2]
            start = fp["start_date"]
            y, m = _add_months(start.year, start.month, 11)
            end = _month_bounds(y, m)[1]
        else:
            chosen = [rng.choice(streams)]
            start = end = transfer_date
        for stream in chosen:
            preset = stream["preset"]
            ctype, csize, cunit = preset["container"]
            facility = ref.RECEIVING_FACILITIES[preset["route"]]
            rd_code = preset["rd_code"] if opts["wtn_rd_in_destination"] else None
            dest = f"{facility['name']}, {facility['address']}"
            if rd_code:
                dest += f" – {rd_code} ({ref.RD_CODE_TEXT.get(rd_code, 'recovery / disposal operation')})"
            destinations.append(dest)
            evidence = {"rd_code": rd_code, "keyword": preset.get("keyword_route")}
            lo, hi = preset["kg_per_lift"]
            if season:
                bins = rng.randint(1, 3)
                annual_t = round(stream["lifts_per_month"] * 12 * (lo + hi) / 2 / 1000)
                annual_t = max(annual_t, 1)
                qtext = f"{preset['description']}: {bins} x {csize}L bins, weekly – est. {annual_t} tonnes per annum" if csize else f"{preset['description']} – est. {annual_t} tonnes per annum"
                lines.append(_waste_line(preset, description=_text(raw_config, preset["description"]), quantity=annual_t, unit=ref.UNIT_T,
                                         basis=ref.BASIS_ESTIMATED, collection_date=None, evidence=evidence,
                                         document_type=ref.DOC_WTN, fill=fill, extra_flags=[ref.FLAG_POSSIBLE_DUPLICATE]))
            elif opts["wtn_quantity_style"] == "container" and csize and ctype in _SKIP_LIKE:
                qtext = f"1 x {csize}yd {'skip' if ctype == ref.CONTAINER_SKIP else 'compactor'}"
                lines.append(_waste_line(preset, description=_text(raw_config, preset["description"]), quantity=csize, unit=cunit,
                                         basis=ref.BASIS_VOLUME, collection_date=transfer_date, evidence=evidence,
                                         document_type=ref.DOC_WTN, fill=fill, lifts=1))
                specials.append("volume_only")
            elif opts["wtn_quantity_style"] == "container" and csize and cunit == ref.UNIT_L:
                n = rng.randint(2, 6)
                qtext = f"{n} x {csize}L bins"
                lines.append(_waste_line(preset, description=_text(raw_config, preset["description"]), quantity=n, unit=ref.UNIT_LIFT,
                                         basis=ref.BASIS_LIFT, collection_date=transfer_date, evidence=evidence,
                                         document_type=ref.DOC_WTN, fill=fill, lifts=n))
                specials.append("lifts_only")
            else:
                tonnes = round(rng.randint(lo, hi) * rng.randint(1, 3) / 1000, 1) or 0.1
                qtext = f"Approx. {tonnes} tonnes"
                lines.append(_waste_line(preset, description=_text(raw_config, preset["description"]), quantity=tonnes, unit=ref.UNIT_T,
                                         basis=ref.BASIS_ESTIMATED, collection_date=transfer_date, evidence=evidence,
                                         document_type=ref.DOC_WTN, fill=fill))
            quantity_texts.append(qtext)
        if season:
            specials += ["season_ticket", "duplicate_across_docs"]
        record = _base_record(
            raw_config, company, site,
            document_role=ref.ROLE_SEASON_TICKET if season else ref.ROLE_ORIGINAL,
            document_number=f"{contractor['code']}-WTN-{'ST-' if season else ''}{rng.randint(10000, 99999)}",
            supplier_name=_text(raw_config, contractor["name"]),
            carrier_reg=contractor["carrier_reg"],
            sic_code=company.get("sic_code") or rng.choice(ref.SIC_CODES),
            document_date=transfer_date if not season else start,
            period_start=start,
            period_end=end,
            lines=lines,
            special_cases=_dedupe(specials) or ["clean_baseline"],
            notes=("Season ticket – annual estimate; must not be summed with invoices/reports for the same streams."
                   if season else "WTN quantity is an estimate; R/D code (if any) is embedded in the destination text."),
            _contractor=contractor,
            _quantity_texts=quantity_texts,
            _destinations=destinations,
            _council=company.get("council") or rng.choice(ref.COUNCILS),
            _time=f"{rng.randint(7, 16):02d}:{rng.choice(['00', '15', '30', '45'])}",
            _signatory=rng.choice(["J. Okafor (Facilities Manager)", "R. Patel (Site Manager)", "S. Lindqvist (EHS Lead)"]),
            _driver=rng.choice(["D. Marsh", "K. Brennan", "T. Walsh"]),
            _containers={s["preset"]["container"][0] for s in chosen},
        )
        records.append(record)
    return records


# ── weighbridge ticket ───────────────────────────────────────────────────────


def _build_weighbridge_tickets(raw_config: dict) -> list[dict]:
    opts = _options(raw_config)
    seed = _seed(raw_config)
    pairs = _company_site_pairs(raw_config)
    records = []
    for index in range(_multi_document_count(raw_config)):
        rng = random.Random(f"{seed}:waste_wb:{index}")
        company, site = pairs[index % len(pairs)]
        contractor = _contractor(company)
        streams = _site_streams(site, hazardous=False) or _fallback_streams(["compactor_general"])
        heavy = [s for s in streams if s["preset"]["container"][0] in _SKIP_LIKE | {ref.CONTAINER_LOOSE}] or streams
        preset = rng.choice(heavy)["preset"]
        facility = ref.RECEIVING_FACILITIES[preset["route"]]
        period_start, period_end = _doc_month(raw_config, index)
        when = _random_day(rng, period_start, period_end)
        lo, hi = preset["kg_per_lift"]
        net_kg = rng.randint(lo, hi)
        tare_kg = rng.randint(11800, 15200)
        gross_kg = net_kg + tare_kg
        unit = opts["wb_unit"]
        mode = opts["wb_mode"]
        printed_net_kg = net_kg + (rng.choice([-1, 1]) * rng.randint(100, 300) if mode == "net_mismatch" else 0)

        def conv(kg):
            return round(kg / 1000, 2) if unit == ref.UNIT_T else kg

        gross, tare, net = conv(gross_kg), conv(tare_kg), conv(printed_net_kg)
        material_text = f"{preset['description']} {preset['ewc']} – {facility['type_text']}"
        evidence = {"stated": preset["route"]}
        specials = ["gross_tare_net"]
        common = dict(description=_text(raw_config, material_text), collection_date=when, evidence=evidence,
                      document_type=ref.DOC_WEIGHBRIDGE, fill=float(opts["fill_factor"]))
        if mode == "gross_only":
            line = _waste_line(preset, quantity=None, unit=unit, basis=None, gross=gross, omit_tonnes=True,
                               extra_flags=[ref.FLAG_GROSS_ONLY], **common)
            specials = ["gross_only"]
        elif mode == "net_mismatch":
            line = _waste_line(preset, quantity=net, unit=unit, basis=ref.BASIS_WEIGHED, gross=gross, tare=tare, net=net,
                               omit_tonnes=True, extra_flags=[ref.FLAG_NET_MISMATCH], **common)
            specials = ["net_mismatch", "gross_tare_net"]
        else:
            line = _waste_line(preset, quantity=net, unit=unit, basis=ref.BASIS_WEIGHED, gross=gross, tare=tare, net=net, **common)
            if mode == "stored_tare":
                specials = ["gross_tare_net", "stored_tare"]
        line["containerSize"] = None
        line["containerSizeUnit"] = None
        time_in = f"{rng.randint(7, 15):02d}:{rng.randint(0, 59):02d}"
        hh, mm_ = map(int, time_in.split(":"))
        mm_ += rng.randint(8, 25)
        time_out = f"{hh + mm_ // 60:02d}:{mm_ % 60:02d}"
        gate_fee = round(rng.uniform(85, 165), 2) if preset["route"] != ref.ROUTE_LANDFILL else round(ref.landfill_tax_rate(when) + rng.uniform(25, 45), 2)
        code, symbol = _currency(company)
        record = _base_record(
            raw_config, company, site,
            document_number=f"WB-{facility['name'][:2].upper()}-{rng.randint(50000, 69999)}",
            supplier_name=_text(raw_config, facility["name"]),
            document_date=when,
            period_start=when,
            period_end=when,
            lines=[line],
            special_cases=specials,
            notes={"gross_only": "Gross weight only – net cannot be determined; gross must NOT be used as the quantity.",
                   "net_mismatch": "Printed net differs from gross − tare beyond tolerance; quantityTonnes omitted – reviewer decides."}.get(mode, "Quantity is NET, not gross."),
            _facility=facility,
            _unit=unit,
            _mode=mode,
            _gross=gross, _tare=tare, _net=net,
            _time_in=time_in, _time_out=time_out,
            _vehicle=f"{rng.choice(['YX', 'YJ', 'YK', 'YB'])}{rng.randint(15, 74)} {rng.choice(['KLM', 'RTA', 'BDF', 'HJK'])}",
            _haulier=contractor["name"],
            _operator=rng.choice(["K. Brennan", "L. Moss", "A. Hussain"]),
            _ewc=preset["ewc"],
            _carrier_reg=contractor["carrier_reg"],
            _account=f"{contractor['code']}-{rng.randint(1000, 9999)}",
            _job_ref=f"WTN-{rng.randint(100000, 999999)}",
            _charge=(f"{_fmt_money(gate_fee, symbol)}/t × {_fmt_number(printed_net_kg / 1000, 2)} t = "
                     f"{_fmt_money(round(gate_fee * printed_net_kg / 1000, 2), symbol)}") if mode != "gross_only" else "On account",
            _stored_tare_date=when - timedelta(days=rng.randint(20, 160)),
            _not_weighed=rng.choice(["NOT WEIGHED", "———— X ————"]),
        )
        records.append(record)
    return records


# ── waste report ─────────────────────────────────────────────────────────────

_SPLIT_COLUMNS = [ref.ROUTE_OPEN_LOOP, ref.ROUTE_COMBUSTION, ref.ROUTE_AD, ref.ROUTE_LANDFILL]
_SPLIT_LABELS = {ref.ROUTE_OPEN_LOOP: "Recycled %", ref.ROUTE_COMBUSTION: "Energy recovery %", ref.ROUTE_AD: "AD %", ref.ROUTE_LANDFILL: "Landfill %"}


def _route_split(preset: dict, rng: random.Random) -> dict[str, int]:
    route = preset["route"]
    if route in (ref.ROUTE_OPEN_LOOP, ref.ROUTE_CLOSED_LOOP):
        rejects = rng.choice([0, 5, 8, 10])
        return {ref.ROUTE_OPEN_LOOP: 100 - rejects, ref.ROUTE_COMBUSTION: rejects}
    if route == ref.ROUTE_COMBUSTION:
        lf = rng.choice([0, 3, 5])
        return {ref.ROUTE_COMBUSTION: 100 - lf, ref.ROUTE_LANDFILL: lf}
    if route == ref.ROUTE_AD:
        return {ref.ROUTE_AD: 100}
    if route == ref.ROUTE_LANDFILL:
        return {ref.ROUTE_LANDFILL: 100}
    return {ref.ROUTE_OPEN_LOOP: 100}


def _build_report(raw_config: dict) -> list[dict]:
    opts = _options(raw_config)
    seed = _seed(raw_config)
    rng = random.Random(f"{seed}:waste_report")
    fp = _financial_period(raw_config)
    months = max(len(_months_in_range(fp["start_date"], fp["end_date"])), 1)
    company = (raw_config.get("companies") or [{}])[0]
    sites = company.get("sites") or [{}]
    multi = bool(opts["rep_multi_site"]) and len(sites) > 1
    sites = sites if multi else sites[:1]
    contractor = _contractor(company)
    rows: list[dict] = []
    lines: list[dict] = []
    for site in sites:
        for stream in _site_streams(site, hazardous=False) or _fallback_streams(["general_waste_1100", "dmr_1100", "food_240"]):
            preset = stream["preset"]
            lo, hi = preset["kg_per_lift"]
            lifts = stream["lifts_per_month"] * months
            tonnes = round(sum(rng.randint(lo, hi) for _ in range(min(lifts, 400))) * (lifts / min(lifts, 400)) / 1000, 2)
            split = _route_split(preset, rng)
            site_value = _site_address_text(site) if multi else None
            flags = [ref.FLAG_MULTI_SITE] if multi else []
            rows.append({"site": site.get("label", ""), "stream": preset["description"], "ewc": preset["ewc"],
                         "container": _container_text(preset), "lifts": lifts, "tonnes": tonnes, "split": split,
                         "primary": ref.ROUTE_LABEL[preset["route"]]})
            common = dict(document_type=ref.DOC_REPORT, fill=float(opts["fill_factor"]), lifts=lifts,
                          collection_date=fp["end_date"], site=site_value, extra_flags=flags)
            if opts["rep_route_split"]:
                for route in _SPLIT_COLUMNS:
                    pct = split.get(route, 0)
                    if not pct:
                        continue
                    final_route = ref.preset_loop_route(preset) if route == ref.ROUTE_OPEN_LOOP else route
                    qty = round(tonnes * pct / 100, 4)
                    lines.append(_waste_line(preset, description=_text(raw_config, f"{preset['description']} – {pct}% {ref.ROUTE_LABEL[route]}"),
                                             quantity=qty, unit=ref.UNIT_T, basis=ref.BASIS_WEIGHED,
                                             evidence={"stated": final_route}, **common))
            else:
                lines.append(_waste_line(preset, description=_text(raw_config, preset["description"]), quantity=tonnes, unit=ref.UNIT_T,
                                         basis=ref.BASIS_WEIGHED, evidence={"stated": preset["route"]}, **common))
    total = round(sum(r["tonnes"] for r in rows), 2)
    specials = []
    if opts["rep_route_split"]:
        specials.append("route_percentage_split")
    if opts["rep_totals_row"]:
        lines.append(_plain_line(ref.LINE_TOTAL, "TOTAL", quantity=total, unit=ref.UNIT_T))
        specials.append("totals_row")
    if multi:
        specials.append("multi_site")
    record = _base_record(
        raw_config, company, sites[0],
        document_number=f"{contractor['code']}-RPT-{fp['start_date']:%Y%m}",
        supplier_name=_text(raw_config, contractor["name"]),
        document_date=fp["end_date"] + timedelta(days=2),
        period_start=fp["start_date"],
        period_end=fp["end_date"],
        lines=lines,
        doc_flags=[ref.FLAG_MULTI_SITE] if multi else [],
        special_cases=specials or ["clean_baseline"],
        notes=("Each stream is split into one line per route using the % columns." if opts["rep_route_split"] else "Primary route stated per stream.")
        + (" TOTAL row must be ignored." if opts["rep_totals_row"] else "")
        + (" Lines must be split per site." if multi else ""),
        _rows=rows,
        _multi=multi,
        _total=total,
        _contractor=contractor,
        _account=f"{(company.get('code') or 'ACC')[:4].upper()}-{rng.randint(1000, 9999)}",
    )
    if multi:
        record["site_address"] = None
    return [record]


def _container_text(preset: dict) -> str:
    ctype, csize, cunit = preset["container"]
    if preset.get("container_text"):
        return preset["container_text"]
    if csize and cunit == ref.UNIT_L:
        return f"{csize}L"
    if csize and cunit == ref.UNIT_YD3:
        return f"{csize}yd {'skip' if ctype == ref.CONTAINER_SKIP else ctype.replace('_', ' ')}"
    return {ref.CONTAINER_LOOSE: "Loose / baled", ref.CONTAINER_BAG: "Bags"}.get(ctype, "Other")


# ── hazardous waste consignment note ─────────────────────────────────────────

_HWCN_NON_HAZ_PRESET = {
    "description": "Absorbents and wiping cloths (non-hazardous)",
    "ewc": "15 02 03",
    "material": ref.MAT_CI,
    "container": (ref.CONTAINER_DRUM, 205, ref.UNIT_L),
    "route": ref.ROUTE_COMBUSTION,
    "rd_code": "R1",
    "keyword_route": None,
    "kg_per_lift": (30, 80),
    "physical_form": "Solid",
    "hazard_codes": "–",
    "components": "Cotton/polypropylene wipes 90%, water 10%",
}


def _build_hwcns(raw_config: dict) -> list[dict]:
    opts = _options(raw_config)
    seed = _seed(raw_config)
    pairs = _company_site_pairs(raw_config)
    records = []
    for index in range(_multi_document_count(raw_config)):
        rng = random.Random(f"{seed}:waste_hwcn:{index}")
        company, site = pairs[index % len(pairs)]
        contractor = _contractor(company)
        streams = _site_streams(site, hazardous=True) or _fallback_streams(["waste_oil_ibc", "oily_rags_drum"])
        chosen = rng.sample(streams, min(2, len(streams)))
        presets = [s["preset"] for s in chosen]
        if opts["hwcn_non_hazardous_line"]:
            presets.append(_HWCN_NON_HAZ_PRESET)
        period_start, period_end = _doc_month(raw_config, index)
        when = _random_day(rng, period_start, period_end)
        lines, part_b, part_e = [], [], []
        specials: list[str] = []
        for i, preset in enumerate(presets):
            rd_code = preset["rd_code"]
            if opts["hwcn_interim"] and i == 0:
                rd_code = "R13" if not rd_code.startswith("D") else "D15"
            elif not opts["hwcn_interim"] and ref.rd_code_is_interim(rd_code):
                rd_code = "R4"
            if ref.rd_code_is_interim(rd_code):
                specials.append("interim_treatment")
            lo, hi = preset["kg_per_lift"]
            received = rng.randint(lo, hi)
            estimate = received
            flags = []
            if opts["hwcn_part_e_differs"] and i == 0:
                estimate = int(round(received * rng.uniform(1.15, 1.35), -1))
                flags.append(ref.FLAG_CONSIGNEE_DIFFERS)
                specials.append("part_e_differs")
            ctype, csize, cunit = preset["container"]
            container_text = preset.get("container_text") or (
                f"{rng.randint(1, 3)} x {csize}L {'drum' if ctype == ref.CONTAINER_DRUM else 'IBC'}" if csize else "Other")
            lines.append(_waste_line(preset, description=_text(raw_config, preset["description"]), quantity=received, unit=ref.UNIT_KG,
                                     basis=ref.BASIS_WEIGHED, collection_date=when, evidence={"rd_code": rd_code},
                                     document_type=ref.DOC_HWCN, fill=float(opts["fill_factor"]), extra_flags=flags))
            lines[-1]["containerSize"] = None
            lines[-1]["containerSizeUnit"] = None
            part_b.append([preset["description"], preset["ewc"], f"{estimate} (est.)" if estimate != received else str(estimate),
                           preset.get("components", ""), preset.get("physical_form", "Solid"), preset.get("hazard_codes", ""),
                           container_text, preset.get("un_number", "–")])
            part_e.append([preset["ewc"], str(received), "Accepted", rd_code])
        if opts["hwcn_non_hazardous_line"]:
            specials.append("mixed_hazard_status")
        consignee = ref.INTERIM_FACILITY if "interim_treatment" in specials else {
            "name": "Meridian Environmental Ltd – Hull Treatment Centre", "permit": "EPR/MB4417PX", "address": "Dock Lane, Hull HU9 1XX"}
        # A1 code: first 6 alphanumerics of the A2 company name (padded with Q) + "/" + 5 chosen characters.
        prefix = "".join(ch for ch in (company.get("name") or company.get("label") or "").upper() if ch.isalnum())[:6].ljust(6, "Q")
        received_time = f"{rng.randint(9, 17):02d}:{rng.choice(['05', '20', '35', '50'])}"
        record = _base_record(
            raw_config, company, site,
            document_number=f"{prefix}/{rng.randint(1, 99999):05d}",
            supplier_name=_text(raw_config, contractor["name"]),
            carrier_reg=contractor["carrier_reg"],
            sic_code=company.get("sic_code") or rng.choice(ref.SIC_CODES),
            document_date=when,
            period_start=when,
            period_end=when,
            lines=lines,
            special_cases=_dedupe(specials) or ["clean_baseline"],
            notes="Use Part E (consignee) quantities. R12/R13/D13–D15 are interim – final route unknown.",
            _contractor=contractor,
            _part_b=part_b,
            _part_e=part_e,
            _consignee=consignee,
            _process=rng.choice(["Maintenance of production machinery and lighting", "Equipment cleaning and painting",
                                 "Fleet workshop servicing", "Site refurbishment works"]),
            _vehicle=f"{rng.choice(['YD', 'YC', 'YE'])}{rng.randint(15, 74)} {rng.choice(['MEV', 'HZW', 'TRK'])}",
            _signatory=rng.choice(["J. Okafor", "R. Patel", "S. Lindqvist"]),
            _collected_time=f"{rng.randint(7, 8):02d}:{rng.choice(['10', '25', '40'])}",
            _received_time=received_time,
            _consignee_signatory=rng.choice(["P. Hargreaves", "M. Doyle", "C. Ahmed"]),
            _phone=f"0113 {rng.randint(200, 499)} {rng.randint(1000, 9999)}",
            _email=f"facilities@{''.join(ch for ch in (company.get('name') or 'org').lower() if ch.isalnum())[:12]}.example",
        )
        records.append(record)
    return records


# ── skip / waste log ─────────────────────────────────────────────────────────


def _build_waste_log(raw_config: dict) -> list[dict]:
    opts = _options(raw_config)
    seed = _seed(raw_config)
    rng = random.Random(f"{seed}:waste_log")
    fill = float(opts["fill_factor"])
    fp = _financial_period(raw_config)
    company = (raw_config.get("companies") or [{}])[0]
    site = (company.get("sites") or [{}])[0]
    contractor = _contractor(company)
    streams = _site_streams(site, hazardous=False) or _fallback_streams(["cd_skip_8", "plasterboard_skip_4", "food_240"])
    if opts["log_hazardous_row"]:
        streams = streams + _fallback_streams(["asbestos"])
    rows: list[dict] = []
    lines: list[dict] = []
    specials: list[str] = []
    for year, month in _months_in_range(fp["start_date"], fp["end_date"]):
        m_start, m_end = _month_bounds(year, month)
        m_start, m_end = max(m_start, fp["start_date"]), min(m_end, fp["end_date"])
        for stream in streams:
            preset = stream["preset"]
            ctype, csize, cunit = preset["container"]
            lo, hi = preset["kg_per_lift"]
            occurrences = 2 if ctype in _SKIP_LIKE and stream["lifts_per_month"] >= 2 else 1
            for _ in range(occurrences):
                when = _random_day(rng, m_start, m_end)
                kg = rng.randint(lo, hi)
                hazardous = preset["ewc"].endswith("*")
                kind = ref.LINE_REBATE if preset.get("rebate") else ref.LINE_WASTE
                evidence = {"keyword": preset.get("keyword_route")}
                size_text = _container_text(preset)
                ref_no = f"{contractor['code']}-{rng.randint(1000, 9999)}" if ctype in _SKIP_LIKE else ""
                comment = "rebate" if preset.get("rebate") else ""
                if hazardous:
                    comment = "see HWCN?"
                common = dict(description=_text(raw_config, preset["description"]), collection_date=when, evidence=evidence,
                              document_type=ref.DOC_SPREADSHEET, fill=fill, line_kind=kind)
                if ctype in _SKIP_LIKE and opts["log_volume_rows"] and not hazardous and rng.random() < 0.5:
                    line = _waste_line(preset, quantity=csize, unit=cunit, basis=ref.BASIS_VOLUME, lifts=1, **common)
                    qty_cell, unit_cell = None, ""
                    comment = comment or "not weighed"
                    specials.append("volume_only")
                elif ctype == ref.CONTAINER_BIN and opts["log_mixed_units"]:
                    n = stream["lifts_per_month"]
                    line = _waste_line(preset, quantity=n, unit=ref.UNIT_LIFT, basis=ref.BASIS_LIFT, lifts=n, **common)
                    qty_cell, unit_cell = n, "bins"
                    specials.append("mixed_units")
                else:
                    if opts["log_mixed_units"]:
                        unit_cell = rng.choice(["kg", "t", "KG"])
                    else:
                        unit_cell = "t"
                    qty_cell = kg if unit_cell.lower() == "kg" else round(kg / 1000, 2)
                    line = _waste_line(preset, quantity=qty_cell, unit=unit_cell.lower(), basis=ref.BASIS_WEIGHED, **common)
                    if unit_cell != "t":
                        specials.append("mixed_units")
                if hazardous:
                    specials.append("hazardous_on_non_hwcn")
                if kind == ref.LINE_REBATE:
                    specials.append("rebate")
                line["containerSize"] = csize if ctype != ref.CONTAINER_BIN or line["quantityUnit"] == ref.UNIT_LIFT else line["containerSize"]
                lines.append(line)
                rows.append({"date": when, "ref": ref_no, "size": size_text, "type": preset["description"], "ewc": preset["ewc"],
                             "qty": qty_cell, "unit": unit_cell, "contractor": contractor["name"].split()[0], "notes": comment})
    order = sorted(range(len(rows)), key=lambda i: rows[i]["date"])
    rows = [rows[i] for i in order]
    lines = [lines[i] for i in order]
    if opts["log_total_row"]:
        lines.append(_plain_line(ref.LINE_TOTAL, "TOTAL (sum over mixed units – meaningless)" if opts["log_mixed_units"] else "TOTAL"))
        specials.append("totals_row")
    record = _base_record(
        raw_config, company, site,
        document_number=f"LOG-{fp['start_date']:%Y%m}",
        supplier_name=None,
        period_start=fp["start_date"],
        period_end=fp["end_date"],
        lines=lines,
        special_cases=_dedupe(specials) or ["clean_baseline"],
        notes="Customer-maintained log; organisation and site come from the header block above the table. Overlaps contractor documents.",
        _rows=rows,
        _title=f"{site.get('label') or 'Site'} – skip & waste collection log",
    )
    return [record]


# ── record pipeline ──────────────────────────────────────────────────────────

_BUILDERS = {
    ref.DOC_INVOICE: _build_invoices,
    ref.DOC_WTN: _build_wtns,
    ref.DOC_WEIGHBRIDGE: _build_weighbridge_tickets,
    ref.DOC_REPORT: _build_report,
    ref.DOC_HWCN: _build_hwcns,
    ref.DOC_SPREADSHEET: _build_waste_log,
}


def _records(raw_config: dict) -> list[dict]:
    records = _BUILDERS[_document_type(raw_config)](raw_config)
    for record in records:
        record["document_id"] = _slugify_filename(record["document_number"], record["document_type"]).lower()
    return corrupt_records(records, f"waste_{_document_type(raw_config)}", _HEADER_FIELD_TYPES, get_bad_data_config(raw_config))


# ── ground truth ─────────────────────────────────────────────────────────────


def _clean(value):
    if isinstance(value, dict):
        return {k: _clean(v) for k, v in value.items() if not k.startswith("_") and v is not None and v != [] and v != ""}
    if isinstance(value, list):
        return [_clean(v) for v in value]
    return value


def _ground_truth_document(record: dict) -> dict:
    return _clean({
        "Scope": ref.SCOPE_LABEL,
        "Category": CATEGORY,
        "documentId": record["document_id"],
        "file": record["document_id"],
        "documentType": record["document_type"],
        "documentRole": record["document_role"],
        "documentNumber": record["document_number"],
        "relatedDocumentNumber": record["related_document_number"],
        "referencedDocumentNumbers": record["referenced_document_numbers"],
        "supplierName": record["supplier_name"],
        "carrierRegistrationNumber": record["carrier_reg"],
        "organisationName": record["organisation"],
        "siteAddress": record["site_address"],
        "sicCode": record["sic_code"],
        "invoiceDate": _iso(record["document_date"]) if record["document_type"] == ref.DOC_INVOICE else None,
        "periodStart": _iso(record["period_start"]),
        "periodEnd": _iso(record["period_end"]),
        "totalCost": record["total_cost"],
        "currency": record["currency_code"] if record["total_cost"] is not None else None,
        "lines": record["lines"],
        "expectedDocumentFlags": record["doc_flags"],
        "specialCases": record["special_cases"],
        "notes": record["notes"],
    })


def generate_ground_truth_json(raw_config: dict) -> bytes:
    documents = [_ground_truth_document(record) for record in _records(raw_config)]
    return json.dumps(documents, indent=2, ensure_ascii=False).encode("utf-8")


# ── block renderers (PDF + DOCX share one document description) ──────────────
# block types: header, heading, kv, table, paragraph, totals, form


def _invoice_blocks(raw_config: dict, record: dict, distractor_plan, layout_plan) -> list[dict]:
    contractor = record["_contractor"]
    eu = record["_eu"]
    symbol = record["currency_symbol"]
    credit = record["document_role"] == ref.ROLE_CREDIT_NOTE
    title = "CREDIT NOTE" if credit else ("FACTUUR / INVOICE" if eu else "TAX INVOICE")
    meta = [("Document no.", record["document_number"]), ("Date", _fmt_date(record["document_date"], eu=eu)),
            ("Service period", f"{_fmt_date(record['period_start'], eu=eu)} – {_fmt_date(record['period_end'], eu=eu)}"),
            ("Account", f"{(raw_config.get('companies') or [{}])[0].get('code') or 'ACC'}-0042")]
    if record.get("related_document_number"):
        meta.append(("Credits invoice", record["related_document_number"]))
    if record.get("referenced_document_numbers"):
        meta.append(("Weighbridge ref.", ", ".join(record["referenced_document_numbers"])))
    meta += [tuple(line.split(": ", 1)) for line in _document_lines(distractor_plan, placement="meta")]
    weight_header = "Weight" if raw_config_opt(raw_config, "inv_unitless_weight") else "Weight (kg)"
    rows = []
    for line in record["lines"]:
        d = line.get("_display", {})
        rows.append([_fmt_date(d.get("date"), eu=eu), line["description"], line.get("ewcCode") or "", d.get("qty", ""),
                     d.get("weight", ""), d.get("rate", ""), _fmt_money(line.get("cost"), symbol, eu=eu)])
    sections = {
        "addresses": [{"type": "kv2", "left_title": "Invoice to", "left": [record["organisation"], "Accounts Payable", *record["site_address_lines"]],
                       "right_title": "Service address", "right": [record["organisation"], record["site_label"], *record["site_address_lines"]]}],
        "meta": [{"type": "kv", "rows": meta}],
        "line_items": [{"type": "table", "headers": ["Date", "Description", "EWC", "Lifts", weight_header, "Rate", "Amount"],
                        "rows": rows, "widths": [21, 62, 17, 11, 20, 22, 25], "right_cols": [3, 4, 5, 6]}],
        "totals": [{"type": "totals", "rows": [("Subtotal", _fmt_money(record["_subtotal"], symbol, eu=eu), False),
                                                (f"VAT {record['_vat_rate']:g}%", _fmt_money(record["_vat"], symbol, eu=eu), False),
                                                ("Total due" if not credit else "Total credit", _fmt_money(record["total_cost"], symbol, eu=eu), True)]}],
        "footer": [{"type": "paragraph", "small": True, "text": " · ".join(
            ["Payment terms 30 days. Please quote the document number on remittance."] + _document_lines(distractor_plan, placement="summary"))}],
    }
    order = ["addresses", "meta", "line_items", "totals", "footer"]
    if layout_plan.get("enabled") and layout_plan.get("section_order"):
        order = [s for s in layout_plan["section_order"] if s in sections] + [s for s in order if s not in layout_plan["section_order"]]
    blocks = [{"type": "header", "title": record["supplier_name"],
               "lines": [", ".join(contractor["address"]), f"Waste carrier reg. {record['carrier_reg'] or contractor['carrier_reg']} · VAT GB 381 2290 17"],
               "doc_title": title}]
    for name in order:
        blocks += sections[name]
    return blocks


def raw_config_opt(raw_config: dict, key: str):
    return _options(raw_config)[key]


_UK_POSTCODE = re.compile(r"\b([A-Z]{1,2}\d[A-Z\d]? ?\d[A-Z]{2})\b")


def _postcode(lines: list[str]) -> str:
    for line in reversed(lines):
        match = _UK_POSTCODE.search(line.upper())
        if match:
            return match.group(1)
    return ""


def _box(checked: bool) -> str:
    return "[X]" if checked else "[ ]"


def _wtn_blocks(raw_config: dict, record: dict, distractor_plan, layout_plan) -> list[dict]:
    del layout_plan
    contractor = record["_contractor"]
    containers = record["_containers"]
    season = record["document_role"] == ref.ROLE_SEASON_TICKET
    lines = record["lines"]
    a_left = [("A1 Description of the waste being transferred", "\n".join(l["description"] for l in lines)),
              ("List of Waste Regulations code(s)", ", ".join(l["ewcCode"] for l in lines))]
    a_right = [("A2 How is the waste contained?",
                f"Loose {_box(ref.CONTAINER_LOOSE in containers)}   Sacks {_box(ref.CONTAINER_BAG in containers)}   "
                f"Skip {_box(bool(containers & {ref.CONTAINER_SKIP, ref.CONTAINER_COMPACTOR}))}   Drum {_box(ref.CONTAINER_DRUM in containers)}   "
                f"Other {_box(bool(containers & {ref.CONTAINER_BIN, ref.CONTAINER_OTHER, ref.CONTAINER_IBC}))}"
                + ("  (wheeled bins)" if ref.CONTAINER_BIN in containers else "")),
               ("A3 How much waste?", "\n".join(record["_quantity_texts"]))]
    postcode = _postcode(record["site_address_lines"])
    b_left = [("B1 Full name", record["_signatory"].split(" (")[0]),
              ("Company name and address", "\n".join([record["organisation"], *record["site_address_lines"]])),
              ("Postcode", postcode), ("SIC code (2007)", record["sic_code"] or ""),
              ("B2 Name of your unitary authority or council", record["_council"])]
    b_right = [("B3 Are you:", ""), ("The producer of the waste?", _box(True)), ("The importer of the waste?", _box(False)),
               ("The local authority?", _box(False)), ("The holder of an environmental permit?", _box(False)),
               ("Registered waste exemption?", _box(False)), ("A registered waste carrier, broker or dealer?", _box(False))]
    c_left = [("C1 Full name", record["_driver"]),
              ("Company name and address", "\n".join([record["supplier_name"], *contractor["address"]])),
              ("Postcode", _postcode(contractor["address"])),
              ("C2 Are you: The local authority?", _box(False))]
    c_right = [("C3 Are you:", ""), ("The holder of an environmental permit?", _box(False)),
               ("Registered waste exemption?", _box(False)),
               ("A registered waste carrier, broker or dealer?", _box(True)),
               ("Registration number", record["carrier_reg"] or ""),
               ("Details (are you a carrier, broker or dealer?)", "Carrier"),
               ("Waste taken to", "\n".join(record["_destinations"]))]
    d_left = [("D1 Address of transfer or collection point", "\n".join(record["site_address_lines"])),
              ("Date of transfer (DD/MM/YYYY)" if not season else "Period covered (multiple transfers)",
               _fmt_date(record["period_start"]) if not season else
               f"{_fmt_date(record['period_start'])} – {_fmt_date(record['period_end'])} (season ticket, max. 12 months)")]
    d_right = [("D2 Broker or dealer who arranged this transfer (if applicable)", ""), ("Registration number", ""),
               ("Time(s)", record["_time"] if not season else "Weekly collections")]
    blocks = [
        {"type": "header", "title": "Duty of care: waste transfer note",
         "lines": ["Keep this page and copy it for future use. Please write as clearly as possible.", f"Note no. {record['document_number']}"],
         "doc_title": None},
        {"type": "form", "title": "Section A – Description of waste", "left": a_left, "right": a_right},
        {"type": "paragraph", "text": "By signing in Section D below I confirm that I have fulfilled my duty to apply the waste hierarchy "
                                      "as required by Regulation 12 of the Waste (England and Wales) Regulations 2011   Yes [X]", "small": True},
        {"type": "form", "title": "Section B – Current holder of the waste – Transferor", "left": b_left, "right": b_right},
        {"type": "form", "title": "Section C – Person collecting the waste – Transferee", "left": c_left, "right": c_right},
        {"type": "form", "title": "Section D – The transfer", "left": d_left, "right": d_right},
        {"type": "form", "title": "Signatures",
         "left": [("Transferor's signature", record["_signatory"].split(" (")[0]), ("Name", record["_signatory"]), ("Representing", record["organisation"])],
         "right": [("Transferee's signature", record["_driver"]), ("Name", f"{record['_driver']} (Driver)"), ("Representing", record["supplier_name"])]},
    ]
    if season:
        blocks.insert(2, {"type": "paragraph", "small": True,
                          "text": "This note covers repeated transfers of the same waste between the same parties for the period stated."})
    extras = _document_lines(distractor_plan)
    if extras:
        blocks.append({"type": "paragraph", "small": True, "text": " · ".join(extras)})
    return blocks


def _weighbridge_blocks(raw_config: dict, record: dict, distractor_plan, layout_plan) -> list[dict]:
    del raw_config, layout_plan
    facility = record["_facility"]
    unit = record["_unit"]
    mode = record["_mode"]
    blank = record["_not_weighed"]
    gross_only = mode == "gross_only"
    day = _fmt_date(record["document_date"])
    fmt = (lambda v: _fmt_number(v, 2)) if unit == ref.UNIT_T else (lambda v: _fmt_number(v, 0))
    if mode == "stored_tare":
        tare_label, tare_when = f"TARE (stored) ({unit})", f"Stored tare {_fmt_date(record['_stored_tare_date'])}"
    else:
        tare_label, tare_when = f"2nd weight – TARE ({unit})", blank if gross_only else f"{day} {record['_time_out']}"
    rows = [("Ticket no.", record["document_number"]),
            ("Customer account", record["_account"]),
            ("Customer / producer", f"{record['organisation']} – {record['site_label']}"),
            ("Haulier", record["_haulier"]),
            ("Waste carrier reg.", record["_carrier_reg"]),
            ("Vehicle reg.", record["_vehicle"]),
            ("WTN / job ref.", record["_job_ref"]),
            ("Product / material", record["lines"][0]["description"]),
            ("EWC code", record["_ewc"]),
            ("1st weight – date / time", f"{day} {record['_time_in']}"),
            (f"1st weight – GROSS ({unit})", fmt(record["_gross"])),
            ("2nd weight – date / time", tare_when),
            (tare_label, blank if gross_only else fmt(record["_tare"])),
            (f"NET ({unit})", blank if gross_only else fmt(record["_net"])),
            ("Charge", record["_charge"])]
    rows += [tuple(line.split(": ", 1)) for line in _document_lines(distractor_plan)]
    blocks = [{"type": "header", "title": record["supplier_name"], "lines": [f"{facility['address']} · Permit {facility['permit']}"],
               "doc_title": "WEIGHBRIDGE TICKET"},
              {"type": "kv", "rows": rows, "mono_rows": [10, 12, 13], "big": True}]
    if gross_only:
        blocks.append({"type": "heading", "text": "GROSS WEIGHT ONLY"})
    blocks.append({"type": "paragraph", "small": True,
                   "text": "Weighed on a weighbridge certified under the Weights and Measures Act 1985. "
                           f"Operator: {record['_operator']}. Driver signature: ____________"})
    return blocks


def _report_blocks(raw_config: dict, record: dict, distractor_plan, layout_plan) -> list[dict]:
    del layout_plan
    opts = _options(raw_config)
    headers, rows = _report_table(raw_config, record, opts)
    sites = sorted({r["site"] for r in record["_rows"] if r["site"]})
    meta = [("Customer", record["organisation"] + (" (group)" if record["_multi"] else "")),
            ("Site", ", ".join(sites) if record["_multi"] else record["site_address"] or record["site_label"]),
            ("Period", f"{raw_config.get('financial_period', {}).get('label', '')} ({_fmt_date(record['period_start'])} – {_fmt_date(record['period_end'])})"),
            ("Account", record["_account"])]
    meta += [tuple(line.split(": ", 1)) for line in _document_lines(distractor_plan, placement="meta")]
    widths = [180 / len(headers)] * len(headers)
    return [
        {"type": "header", "title": record["supplier_name"],
         "lines": ["Account management – sustainability reporting"],
         "doc_title": "Annual Waste & Diversion Summary" if record["_multi"] else "Monthly Waste & Diversion Report"},
        {"type": "kv", "rows": meta},
        {"type": "table", "headers": headers, "rows": rows, "widths": widths, "bold_last": bool(opts["rep_totals_row"]),
         "right_cols": [i for i, h in enumerate(headers) if h in {"Lifts", "Tonnes"} or h.endswith("%")]},
        {"type": "paragraph", "small": True, "text": "Tonnages from on-board vehicle weighing. Rejects from DMR sorting are sent for energy recovery."},
    ]


def _report_field_ids(opts: dict, multi: bool) -> list[str]:
    ids = (["site"] if multi else []) + ["stream", "ewc", "container", "lifts", "tonnes"]
    ids += [f"pct_{route}" for route in _SPLIT_COLUMNS] if opts["rep_route_split"] else ["primary_route"]
    return ids


_REPORT_HEADERS = {"site": "Site", "stream": "Waste stream", "ewc": "EWC", "container": "Container", "lifts": "Lifts",
                   "tonnes": "Tonnes", "primary_route": "Primary route",
                   **{f"pct_{route}": label for route, label in _SPLIT_LABELS.items()}}


def _report_row_map(row: dict) -> dict:
    out = {"site": row["site"], "stream": row["stream"], "ewc": row["ewc"], "container": row["container"],
           "lifts": row["lifts"], "tonnes": row["tonnes"], "primary_route": row["primary"]}
    for route in _SPLIT_COLUMNS:
        out[f"pct_{route}"] = row["split"].get(route, 0)
    return out


def _report_table(raw_config: dict, record: dict, opts: dict) -> tuple[list[str], list[list[str]]]:
    ids = _report_field_ids(opts, record["_multi"])
    headers = [_REPORT_HEADERS[i] for i in ids]
    rows = []
    for row in record["_rows"]:
        m = _report_row_map(row)
        rows.append([f"{m[i]}%" if i.startswith("pct_") else (_fmt_number(m[i], 2) if i == "tonnes" else str(m[i])) for i in ids])
    if opts["rep_totals_row"]:
        total = ["" for _ in ids]
        total[0] = "TOTAL" if not record["_multi"] else "All sites"
        if record["_multi"]:
            total[1] = "TOTAL"
        total[ids.index("lifts")] = str(sum(r["lifts"] for r in record["_rows"]))
        total[ids.index("tonnes")] = _fmt_number(record["_total"], 2)
        rows.append(total)
    return headers, rows


def _hwcn_blocks(raw_config: dict, record: dict, distractor_plan, layout_plan) -> list[dict]:
    del raw_config, layout_plan
    contractor = record["_contractor"]
    consignee = record["_consignee"]
    date_text = _fmt_date(record["document_date"])
    contact = f"Tel {record['_phone']} · {record['_email']}"
    blocks = [
        {"type": "header", "title": "HAZARDOUS WASTE CONSIGNMENT NOTE", "lines": [f"Consignment note code: {record['document_number']}"], "doc_title": None},
        {"type": "kv", "title": "PART A – Notification details", "rows": [
            ("A1 Consignment note code", record["document_number"]),
            ("A2 The waste described below is to be removed from (name, address, postcode, telephone, email, fax)",
             f"{record['organisation']}, {record['site_address']} · {contact} · Fax N/A"),
            ("A3 The waste will be taken to (name, address and postcode)", f"{consignee['name']}, {consignee['address']}"),
            ("A4 The waste producer was (if different from A2)", "As A2")]},
        {"type": "kv", "title": "PART B – Description of the waste", "rows": [
            ("B1 Process giving rise to the waste", record["_process"]), ("B2 SIC (2007) for the process", record["sic_code"] or "")]},
        {"type": "table", "title": "B3 Waste details (where more than one waste type is collected all of the information given below must be completed for each EWC identified)",
         "headers": ["Description of waste", "List of wastes / EWC code", "Quantity (kg)", "Chemical/biological components and concentrations",
                     "Physical form", "Hazard code(s)", "Container type, number and size", "UN number"],
         "rows": record["_part_b"], "widths": [27, 18, 17, 33, 15, 18, 30, 22], "right_cols": [2]},
        {"type": "kv", "title": "PART C – Carrier's certificate", "rows": [
            ("I certify that I today collected the consignment and that the details in A2, A4 and B3 are correct", record["supplier_name"]),
            ("Carrier registration no.", record["carrier_reg"] or contractor["carrier_reg"]),
            ("Vehicle registration no. (or mode of transport if not road)", record["_vehicle"]),
            ("Date / time of collection", f"{date_text} {record['_collected_time']}")]},
        {"type": "kv", "title": "PART D – Consignor's certificate", "rows": [
            ("I certify that the information in A, B and C has been completed and is correct", f"{record['_signatory']}"),
            ("On behalf of", record["organisation"]), ("Date / time", f"{date_text} {record['_collected_time']}")]},
        {"type": "table", "title": "PART E – Consignee's certificate",
         "headers": ["Individual EWC code(s) received", "Quantity of each EWC code received (kg)", "EWC code accepted / rejected",
                     "Waste management operation (R or D code)"],
         "rows": record["_part_e"], "widths": [40, 50, 40, 50], "right_cols": [1]},
        {"type": "kv", "rows": [
            ("I received this waste at the address given in A3 on", f"{date_text} {record['_received_time']}"),
            ("Vehicle registration no.", record["_vehicle"]),
            ("Where waste is rejected please provide details", "–"),
            ("Name / on behalf of", f"{record['_consignee_signatory']}, {consignee['name']}, {consignee['address']}"),
            ("I certify that the waste permit or exempt waste operation number is", consignee["permit"])]},
    ]
    extras = _document_lines(distractor_plan)
    if extras:
        blocks.append({"type": "paragraph", "small": True, "text": " · ".join(extras)})
    return blocks


_BLOCK_BUILDERS = {
    ref.DOC_INVOICE: _invoice_blocks,
    ref.DOC_WTN: _wtn_blocks,
    ref.DOC_WEIGHBRIDGE: _weighbridge_blocks,
    ref.DOC_REPORT: _report_blocks,
    ref.DOC_HWCN: _hwcn_blocks,
}

_FOOTER = "Synthetic test document – fictional data – Scope 3 waste extraction test set"


# ── PDF ──────────────────────────────────────────────────────────────────────

_STYLES = getSampleStyleSheet()
_H1 = _STYLES["Title"]
_H3 = _STYLES["Heading3"]
_BODY = _STYLES["BodyText"]
_SMALL = _BODY.clone("waste_small", fontSize=7.5, leading=9.5)
_CELL = _BODY.clone("waste_cell", fontSize=8, leading=10)
_CELL_BOLD = _CELL.clone("waste_cell_bold", fontName="Helvetica-Bold")


def _p(text, style=_CELL) -> Paragraph:
    safe = str(text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
    return Paragraph(safe, style)


def _pdf_flowables(blocks: list[dict]) -> list:
    story: list = []
    accent = colors.HexColor("#2F5D3A")
    for block in blocks:
        kind = block["type"]
        if kind == "header":
            story.append(_p(block["title"], _H3.clone("waste_h3_bold", fontName="Helvetica-Bold")))
            for line in block.get("lines", []):
                story.append(_p(line, _SMALL))
            if block.get("doc_title"):
                story.append(_p(block["doc_title"], _H1))
            story.append(Spacer(1, 4))
        elif kind == "heading":
            story.append(_p(block["text"], _H3))
        elif kind == "paragraph":
            story.append(_p(block["text"], _SMALL if block.get("small") else _BODY))
            story.append(Spacer(1, 4))
        elif kind == "kv":
            if block.get("title"):
                story.append(_p(block["title"], _CELL_BOLD))
            style = _CELL.clone("kvbig", fontSize=10, leading=12) if block.get("big") else _CELL
            mono = style.clone("kvmono", fontName="Courier-Bold")
            mono_rows = set(block.get("mono_rows", []))
            rows = [[_p(k, style), _p(v, mono if i in mono_rows else style)] for i, (k, v) in enumerate(block["rows"])]
            t = Table(rows, colWidths=[60 * mm, 120 * mm])
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                   ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#F1F4F2"))]))
            story += [t, Spacer(1, 6)]
        elif kind == "kv2":
            rows = [[_p(block["left_title"], _CELL_BOLD), _p(block["right_title"], _CELL_BOLD)],
                    [_p("\n".join(block["left"])), _p("\n".join(block["right"]))]]
            t = Table(rows, colWidths=[90 * mm, 90 * mm])
            t.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.4, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
            story += [t, Spacer(1, 6)]
        elif kind == "table":
            if block.get("title"):
                story.append(_p(block["title"], _CELL_BOLD))
            right = set(block.get("right_cols", []))
            right_style = _CELL.clone("cell_r", alignment=2)
            data = [[_p(h, _CELL_BOLD) for h in block["headers"]]]
            for r_index, row in enumerate(block["rows"]):
                bold = block.get("bold_last") and r_index == len(block["rows"]) - 1
                data.append([_p(v, (_CELL_BOLD if bold else (right_style if c in right else _CELL))) for c, v in enumerate(row)])
            t = Table(data, colWidths=[w * mm for w in block["widths"]], repeatRows=1)
            t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                                   ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E3ECE6"))]))
            story += [t, Spacer(1, 6)]
        elif kind == "totals":
            data = [[_p(label, _CELL_BOLD if bold else _CELL), _p(value, _CELL_BOLD if bold else _CELL)] for label, value, bold in block["rows"]]
            t = Table(data, colWidths=[35 * mm, 35 * mm], hAlign="RIGHT")
            t.setStyle(TableStyle([("LINEABOVE", (0, -1), (-1, -1), 0.6, accent)]))
            story += [t, Spacer(1, 8)]
        elif kind == "form":
            title = Table([[_p(block["title"], _CELL_BOLD.clone("form_title", fontSize=10.5, leading=13))]], colWidths=[180 * mm])
            title.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, 0), 1.2, colors.black)]))
            n = max(len(block["left"]), len(block["right"]))
            rows = []
            for i in range(n):
                left = block["left"][i] if i < len(block["left"]) else ("", "")
                right = block["right"][i] if i < len(block["right"]) else ("", "")
                rows.append([_p(left[0], _SMALL), _p(left[1]), _p(right[0], _SMALL), _p(right[1])])
            t = Table(rows, colWidths=[38 * mm, 52 * mm, 52 * mm, 38 * mm])
            t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 1),
                                   ("BOTTOMPADDING", (0, 0), (-1, -1), 1), ("LINEBELOW", (1, 0), (1, -1), 0.3, colors.grey),
                                   ("LINEBELOW", (3, 0), (3, -1), 0.3, colors.grey)]))
            story += [KeepTogether([title, t]), Spacer(1, 3)]
    return story


def _render_pdf(raw_config: dict, records: list[dict]) -> bytes:
    builder = _BLOCK_BUILDERS[_document_type(raw_config)]
    distractor_plan = _distractor_plan(raw_config, "PDF")
    layout_plan = _layout_plan(raw_config, "PDF")
    story: list = []
    for index, record in enumerate(records):
        if index:
            story.append(PageBreak())
        story += _pdf_flowables(builder(raw_config, record, distractor_plan, layout_plan))
    buffer = BytesIO()
    doc_cfg = raw_config.get("document", {})

    def footer(canvas, _doc):
        canvas.setFont("Helvetica", 6)
        canvas.setFillColor(colors.grey)
        canvas.drawString(15 * mm, 8 * mm, _FOOTER)

    SimpleDocTemplate(buffer, pagesize=A4, leftMargin=15 * mm, rightMargin=15 * mm, topMargin=14 * mm, bottomMargin=14 * mm,
                      title=doc_cfg.get("title", ""), subject=doc_cfg.get("subject", "")).build(story, onFirstPage=footer, onLaterPages=footer)
    return buffer.getvalue()


# ── DOCX ─────────────────────────────────────────────────────────────────────


def _docx_table(document, rows: list[list[str]], *, header: bool = False, bold_last: bool = False):
    table = document.add_table(rows=len(rows), cols=max(len(r) for r in rows))
    table.style = "Table Grid"
    for r_index, row in enumerate(rows):
        for c_index, value in enumerate(row):
            cell = table.cell(r_index, c_index)
            bold = (header and r_index == 0) or (bold_last and r_index == len(rows) - 1)
            _set_docx_cell_text(cell, str(value or ""), bold=bold, size=8.5)
            if header and r_index == 0:
                _shade_docx_cell(cell, "E3ECE6")
    document.add_paragraph()
    return table


def _render_docx(raw_config: dict, records: list[dict]) -> bytes:
    builder = _BLOCK_BUILDERS[_document_type(raw_config)]
    distractor_plan = _distractor_plan(raw_config, "DOCX")
    layout_plan = _layout_plan(raw_config, "DOCX")
    document = Document()
    _style_docx_document(document)
    doc_cfg = raw_config.get("document", {})
    document.core_properties.title = doc_cfg.get("title", "")
    document.core_properties.subject = doc_cfg.get("subject", "")
    for index, record in enumerate(records):
        if index:
            document.add_page_break()
        for block in builder(raw_config, record, distractor_plan, layout_plan):
            kind = block["type"]
            if kind == "header":
                run = document.add_paragraph().add_run(block["title"])
                run.bold = True
                run.font.size = Pt(13)
                for line in block.get("lines", []):
                    p = document.add_paragraph(line)
                    p.runs[0].font.size = Pt(8)
                if block.get("doc_title"):
                    p = document.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(block["doc_title"])
                    r.bold = True
                    r.font.size = Pt(15)
            elif kind == "heading":
                r = document.add_paragraph().add_run(block["text"])
                r.bold = True
            elif kind == "paragraph":
                p = document.add_paragraph(block["text"])
                if block.get("small"):
                    p.runs[0].font.size = Pt(7.5)
            elif kind == "kv":
                if block.get("title"):
                    document.add_paragraph().add_run(block["title"]).bold = True
                _docx_table(document, [[k, v] for k, v in block["rows"]])
            elif kind == "kv2":
                _docx_table(document, [[block["left_title"], block["right_title"]], ["\n".join(block["left"]), "\n".join(block["right"])]], header=True)
            elif kind == "table":
                if block.get("title"):
                    document.add_paragraph().add_run(block["title"]).bold = True
                _docx_table(document, [block["headers"], *block["rows"]], header=True, bold_last=bool(block.get("bold_last")))
            elif kind == "totals":
                for label, value, bold in block["rows"]:
                    p = document.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    r = p.add_run(f"{label}: {value}")
                    r.bold = bold
            elif kind == "form":
                p = document.add_paragraph()
                r = p.add_run(block["title"])
                r.bold = True
                r.font.size = Pt(11)
                n = max(len(block["left"]), len(block["right"]))
                rows = []
                for i in range(n):
                    left = block["left"][i] if i < len(block["left"]) else ("", "")
                    right = block["right"][i] if i < len(block["right"]) else ("", "")
                    rows.append([left[0], left[1], right[0], right[1]])
                _docx_table(document, rows)
        footer = document.add_paragraph(_FOOTER)
        footer.runs[0].font.size = Pt(6)
        footer.runs[0].font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    buffer = BytesIO()
    document.save(buffer)
    return buffer.getvalue()


# ── tabular (report + log) ───────────────────────────────────────────────────

_LOG_FIELD_IDS = ["date", "ref", "size", "type", "ewc", "qty", "unit", "contractor", "notes"]
_LOG_HEADERS = {"date": "Date", "ref": "Skip / bin ref", "size": "Size", "type": "Waste type", "ewc": "EWC", "qty": "Qty",
                "unit": "Unit", "contractor": "Collected by", "notes": "Comments"}


def _tabular_spec(raw_config: dict, record: dict, output_format: str):
    """Return (title_rows, field_ids, headers, data_rows, total_row, notes) for report/log."""
    opts = _options(raw_config)
    plan = _layout_plan(raw_config, output_format)
    distractor_plan = _distractor_plan(raw_config, output_format)
    if record["document_type"] == ref.DOC_REPORT:
        base_ids = _report_field_ids(opts, record["_multi"])
        header_map = _REPORT_HEADERS
        row_maps = [_report_row_map(r) for r in record["_rows"]]
        title_rows = [[f"{record['supplier_name']} – Customer Portal export"],
                      [f"Account {record['_account']} · {record['organisation']} · Period "
                       f"{_fmt_date(record['period_start'])} – {_fmt_date(record['period_end'])}"], []]
    else:
        base_ids = list(_LOG_FIELD_IDS)
        header_map = _LOG_HEADERS
        row_maps = [{**r, "date": r["date"].isoformat()} for r in record["_rows"]]
        # Header block so the organisation and site can be parsed from the file itself.
        title_rows = [
            [record["_title"]],
            ["Company", record["organisation"]],
            ["Site", record["site_label"]],
            ["Site address", record["site_address"]],
            ["Period", f"{_fmt_date(record['period_start'])} – {_fmt_date(record['period_end'])}"],
            [],
        ]
    ordered = _augment_field_ids(_ordered_field_ids(plan, base_ids), distractor_plan)
    distractors = _distractor_field_map(distractor_plan)
    headers = [distractors[i].label if i in distractors else (plan.get("header_aliases", {}).get(i) or header_map[i]) for i in ordered]
    data = [_row_values(m, ordered, distractor_plan, row_key=f"r{n}") for n, m in enumerate(row_maps)]
    return plan, title_rows, ordered, headers, data


def _report_total_values(record: dict, ordered: list[str]) -> list[Any]:
    values: list[Any] = ["" for _ in ordered]
    values[0] = "TOTAL"
    if "lifts" in ordered:
        values[ordered.index("lifts")] = sum(r["lifts"] for r in record["_rows"])
    if "tonnes" in ordered:
        values[ordered.index("tonnes")] = record["_total"]
    return values


def _render_xlsx(raw_config: dict, records: list[dict]) -> bytes:
    record = records[0]
    opts = _options(raw_config)
    plan, title_rows, ordered, headers, data = _tabular_spec(raw_config, record, "XLSX")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Collections" if record["document_type"] == ref.DOC_REPORT else "Waste log"
    row = _write_xlsx_preamble(ws, 1, plan)
    for title in title_rows:
        for c, value in enumerate(title, start=1):
            ws.cell(row=row, column=c, value=value).font = Font(bold=(c == 1))
        row += 1
    header_row = row
    for c, value in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=value)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DDE7EE")
    row += 1
    first_data = row
    for n, values in enumerate(data):
        if record["document_type"] == ref.DOC_SPREADSHEET and opts["log_blank_rows"] and n == 3:
            row += 1
        for c, value in enumerate(values, start=1):
            ws.cell(row=row, column=c, value=value)
        row += 1
    if record["document_type"] == ref.DOC_REPORT and opts["rep_totals_row"]:
        for c, value in enumerate(_report_total_values(record, ordered), start=1):
            ws.cell(row=row, column=c, value=value).font = Font(bold=True)
    elif record["document_type"] == ref.DOC_SPREADSHEET and opts["log_total_row"]:
        row += 1
        ws.cell(row=row, column=1, value="TOTAL").font = Font(bold=True)
        if "qty" in ordered:
            col = openpyxl.utils.get_column_letter(ordered.index("qty") + 1)
            ws.cell(row=row, column=ordered.index("qty") + 1, value=f"=SUM({col}{first_data}:{col}{row - 2})")
            if "unit" in ordered:
                ws.cell(row=row, column=ordered.index("unit") + 1, value="mixed!" if opts["log_mixed_units"] else "t")
    for c in range(1, len(headers) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 16
    del header_row
    if record["document_type"] == ref.DOC_SPREADSHEET and opts["log_notes_sheet"]:
        notes = wb.create_sheet("Notes")
        notes.append(["Kept by facilities. Quantities as reported by collectors or estimated by site staff."])
        notes.append(["Skips marked 'not weighed' were charged per lift."])
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _render_csv(raw_config: dict, records: list[dict]) -> bytes:
    record = records[0]
    opts = _options(raw_config)
    plan, title_rows, ordered, headers, data = _tabular_spec(raw_config, record, "CSV")
    buffer = StringIO()
    writer = csv.writer(buffer)
    _write_csv_preamble(writer, plan)
    for title in title_rows:
        writer.writerow(title)
    writer.writerow(headers)
    for n, values in enumerate(data):
        if record["document_type"] == ref.DOC_SPREADSHEET and opts["log_blank_rows"] and n == 3:
            writer.writerow([])
        writer.writerow(["" if v is None else v for v in values])
    if record["document_type"] == ref.DOC_REPORT and opts["rep_totals_row"]:
        writer.writerow(_report_total_values(record, ordered))
    elif record["document_type"] == ref.DOC_SPREADSHEET and opts["log_total_row"]:
        values = ["" for _ in ordered]
        values[0] = "TOTAL"
        if "qty" in ordered:
            values[ordered.index("qty")] = _fmt_number(sum(r["qty"] for r in record["_rows"] if isinstance(r["qty"], (int, float))), 2, thousands=False)
        if "unit" in ordered:
            values[ordered.index("unit")] = "mixed!" if opts["log_mixed_units"] else "t"
        writer.writerow([])
        writer.writerow(values)
    return buffer.getvalue().encode("utf-8")


# ── public API ───────────────────────────────────────────────────────────────


def _filename(record: dict, index: int, ext: str) -> str:
    return f"{_slugify_filename(record.get('document_number'), f'waste_{index + 1:04d}')}.{ext}"


def _render_documents(raw_config: dict, render_one, ext: str) -> bytes:
    records = _records(raw_config)
    if len(records) > 1:
        return _zip_documents(raw_config, records, render_one, _filename, ext)
    return render_one(raw_config, records)


def generate_pdf(raw_config: dict) -> bytes:
    return _render_documents(raw_config, _render_pdf, "pdf")


def generate_docx(raw_config: dict) -> bytes:
    return _render_documents(raw_config, _render_docx, "docx")


def generate_xlsx(raw_config: dict) -> bytes:
    return _render_xlsx(raw_config, _records(raw_config))


def generate_csv(raw_config: dict) -> bytes:
    return _render_csv(raw_config, _records(raw_config))
