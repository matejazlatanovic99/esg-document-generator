"""Electricity business logic and section building."""
from __future__ import annotations

import random
from datetime import datetime, time, timedelta, timezone
from decimal import Decimal, ROUND_HALF_UP

from generators.shared_generator import (
    DEFAULT_COMPANY_STYLES,
    billing_period_dates,
    billing_period_factor,
    billing_period_label,
    derive_month_periods,
    invoice_suffix,
    normalize_billing_periods,
)
from utils.currency import currency_symbol

TWOPLACES = Decimal("0.01")
FOURPLACES = Decimal("0.0001")


def _q2(value) -> Decimal:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    return value.quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def _parse_decimal(value) -> Decimal:
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def _q4(value) -> Decimal:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    return value.quantize(FOURPLACES, rounding=ROUND_HALF_UP)


def _to_kwh(value, unit: str) -> Decimal:
    dec_value = _parse_decimal(value)
    return dec_value * Decimal("1000") if unit == "MWh" else dec_value


def _smart_meter_interval_minutes(config: dict) -> int:
    minutes = int(config["document"].get("smart_meter_interval_minutes", 30))
    return minutes if minutes in {15, 30, 60} else 30


def _smart_meter_interval_value_mode(config: dict) -> str:
    mode = str(config["document"].get("smart_meter_interval_value_mode", "consumption_diff")).lower()
    return mode if mode in {"consumption_diff", "cumulative_end_reading"} else "consumption_diff"


def _smart_meter_timestamp_format(config: dict) -> str:
    fmt = str(config["document"].get("smart_meter_timestamp_format", "iso_8601_utc")).lower()
    return fmt if fmt in {"iso_8601_utc", "datetime"} else "iso_8601_utc"


def _format_smart_meter_timestamp(ts: datetime, timestamp_format: str) -> str:
    if timestamp_format == "datetime":
        return ts.strftime("%Y-%m-%d %H:%M:%S")
    return ts.strftime("%Y-%m-%dT%H:%M:%SZ")


def _interval_timestamps(period_start, period_end, interval_minutes: int) -> list[datetime]:
    start_dt = datetime.combine(period_start, time.min, tzinfo=timezone.utc)
    end_dt = datetime.combine(period_end + timedelta(days=1), time.min, tzinfo=timezone.utc)
    timestamps: list[datetime] = []
    current = start_dt
    while current < end_dt:
        timestamps.append(current)
        current += timedelta(minutes=interval_minutes)
    return timestamps


def _interval_weight(ts: datetime, rng: random.Random) -> float:
    hour = ts.hour + (ts.minute / 60)
    if 0 <= hour < 5:
        base = 0.45
    elif 5 <= hour < 8:
        base = 0.8
    elif 8 <= hour < 18:
        base = 1.25
    elif 18 <= hour < 22:
        base = 1.05
    else:
        base = 0.7

    if ts.weekday() >= 5:
        base *= 0.85

    return max(base * rng.uniform(0.88, 1.12), 0.01)


def _distribute_interval_values(total_kwh: Decimal, timestamps: list[datetime], rng: random.Random) -> list[Decimal]:
    if not timestamps:
        return []

    weights = [_interval_weight(ts, rng) for ts in timestamps]
    weight_total = sum(weights)
    values = [_q4(total_kwh * Decimal(str(weight / weight_total))) for weight in weights]
    values[-1] = _q4(total_kwh - sum(values[:-1]))
    return values


def _monthly_smart_meter_rows(sections: list[dict]) -> list[dict]:
    rows: list[dict] = []
    for section in sections:
        company = section["company"]
        site = section["site"]
        unit = site.get("unit", "kWh")
        normalized_unit = "kWh"
        start_reading = int(_to_kwh(site["start_reading"], unit))
        end_reading = int(_to_kwh(site["end_reading"], unit))
        tariffs = site.get("tariffs", [])

        if tariffs:
            for tariff in tariffs:
                rows.append({
                    "meter_id": site["meter_id"],
                    "currency": company.get("currency", "GBP (£)").split()[0],
                    "site_label": site["label"],
                    "period_label": site["billing_period_label"],
                    "start_reading": start_reading,
                    "end_reading": end_reading,
                    "consumption": float(_q2(_to_kwh(tariff["quantity"], unit))),
                    "unit": normalized_unit,
                    "tariff_type": tariff["name"],
                    "cost": float(_q2(tariff["cost"])),
                })
        else:
            rows.append({
                "meter_id": site["meter_id"],
                "currency": company.get("currency", "GBP (£)").split()[0],
                "site_label": site["label"],
                "period_label": site["billing_period_label"],
                "start_reading": start_reading,
                "end_reading": end_reading,
                "consumption": float(_q2(_to_kwh(site["total_quantity"], unit))),
                "unit": normalized_unit,
                "tariff_type": "",
                "cost": float(_q2(site["total_cost"])),
            })
    return rows


def _interval_smart_meter_rows(config: dict, sections: list[dict]) -> list[dict]:
    rows: list[dict] = []
    interval_minutes = _smart_meter_interval_minutes(config)
    value_mode = _smart_meter_interval_value_mode(config)
    timestamp_format = _smart_meter_timestamp_format(config)
    seed = int(config.get("random_seed", 42))

    for section_index, section in enumerate(sections, start=1):
        site = section["site"]
        unit = site.get("unit", "kWh")
        timestamps = _interval_timestamps(site["period_start"], site["period_end"], interval_minutes)
        total_import_kwh = _to_kwh(site["total_quantity"], unit)
        rng = random.Random(f"{seed}:{site.get('_site_uid', site['meter_id'])}:{section_index}:{interval_minutes}")
        interval_values = _distribute_interval_values(total_import_kwh, timestamps, rng)
        cumulative_total = Decimal(str(site["start_reading"]))

        for ts, import_kwh in zip(timestamps, interval_values):
            row = {
                "meter_id": site["meter_id"],
                "timestamp": _format_smart_meter_timestamp(ts, timestamp_format),
                "unit": "kWh",
                "value_mode": value_mode,
            }
            if value_mode == "cumulative_end_reading":
                cumulative_total += import_kwh
                row["end_reading"] = float(_q4(cumulative_total))
            else:
                row["import_kwh"] = float(import_kwh)
                row["export_kwh"] = 0.0
            rows.append(row)

    return rows


def build_smart_meter_rows(config: dict, sections: list[dict]) -> list[dict]:
    granularity = str(config["document"].get("smart_meter_data_granularity", "monthly")).lower()
    if granularity == "interval":
        return _interval_smart_meter_rows(config, sections)
    return _monthly_smart_meter_rows(sections)


def _split_among_tariffs(
    total_qty: Decimal,
    total_cost: Decimal,
    tariff_names: list[str],
    rng: random.Random,
) -> list[dict]:
    n = len(tariff_names)
    if n == 0:
        return []
    if n == 1:
        unit_cost = _q2(total_cost / total_qty) if total_qty > 0 else Decimal("0")
        return [{"name": tariff_names[0], "quantity": total_qty, "unit_cost": unit_cost, "cost": total_cost}]

    qty_weights = [rng.random() for _ in range(n)]
    cost_weights = [w * rng.uniform(0.7, 1.3) for w in qty_weights]
    qty_total_w = sum(qty_weights)
    cost_total_w = sum(cost_weights)

    qtys = [_q2(total_qty * Decimal(str(w / qty_total_w))) for w in qty_weights]
    costs = [_q2(total_cost * Decimal(str(w / cost_total_w))) for w in cost_weights]
    qtys[-1] = _q2(total_qty - sum(qtys[:-1]))
    costs[-1] = _q2(total_cost - sum(costs[:-1]))

    result = []
    for name, qty, cost in zip(tariff_names, qtys, costs):
        unit_cost = _q2(cost / qty) if qty > 0 else Decimal("0")
        result.append({"name": name, "quantity": qty, "unit_cost": unit_cost, "cost": cost})
    return result


def _distribute_annual(
    annual: Decimal,
    factors: list[Decimal],
    rng: random.Random,
    jitter: float = 0.07,
) -> list[Decimal]:
    total_factor = sum(factors)
    if not factors or total_factor == 0:
        return []

    distributed: list[Decimal] = []
    for factor in factors:
        raw = annual * Decimal(str(factor)) / Decimal(str(total_factor))
        variation = Decimal(str(1 + rng.uniform(-jitter, jitter)))
        distributed.append(_q2(raw * variation))
    distributed[-1] = _q2(annual - sum(distributed[:-1]))
    return distributed


def _generate_period_records(annual_site: dict, billing_periods: list, rng: random.Random) -> list[dict]:
    factors = [Decimal(str(billing_period_factor(period))) for period in billing_periods]

    annual_qty = annual_site["total_quantity"]
    annual_cost = annual_site["total_cost"]
    supplier_ef = annual_site["supplier_ef"]
    has_supplier_ef = annual_site["_has_supplier_ef"]
    unit = annual_site["unit"]

    qtys = _distribute_annual(annual_qty, factors, rng)
    costs = _distribute_annual(annual_cost, factors, rng)

    annual_tariffs = annual_site["tariffs"]
    tariff_qty_by_period = [_distribute_annual(t["quantity"], factors, rng) for t in annual_tariffs]
    tariff_cost_by_period = [_distribute_annual(t["cost"], factors, rng) for t in annual_tariffs]

    prev = annual_site["start_reading"]
    records: list[dict] = []
    for idx, period in enumerate(billing_periods, start=1):
        first, last = billing_period_dates(period)
        qty = qtys[idx - 1]
        cost = costs[idx - 1]
        curr = prev + int(qty)

        if has_supplier_ef:
            qty_kwh = qty * Decimal("1000") if unit == "MWh" else qty
            emissions_kg = _q2(qty_kwh * supplier_ef)
            emissions_t = (emissions_kg / Decimal("1000")).quantize(Decimal("1.000"), rounding=ROUND_HALF_UP)
        else:
            emissions_kg = Decimal("0")
            emissions_t = Decimal("0")

        period_tariffs = [
            {
                "name": tariff["name"],
                "quantity": tariff_qty_by_period[k][idx - 1],
                "unit": unit,
                "unit_cost": tariff["unit_cost"],
                "cost": tariff_cost_by_period[k][idx - 1],
            }
            for k, tariff in enumerate(annual_tariffs)
        ]

        records.append({
            "label": annual_site["label"],
            "customer": annual_site["customer"],
            "customer_code": annual_site["customer_code"],
            "customer_address": annual_site["customer_address"],
            "city": annual_site["city"],
            "postcode": annual_site["postcode"],
            "meter_id": annual_site["meter_id"],
            "supplier_ef": supplier_ef,
            "unit": unit,
            "currency_symbol": annual_site["currency_symbol"],
            "_omit": annual_site["_omit"],
            "_site_uid": annual_site["_site_uid"],
            "start_reading": prev,
            "end_reading": curr,
            "total_quantity": qty,
            "total_cost": cost,
            "emissions_kg": emissions_kg,
            "emissions_t": emissions_t,
            "tariffs": period_tariffs,
            "period_start": first,
            "period_end": last,
            "billing_period_label": billing_period_label(period),
            "ref_no": f"{annual_site['_base_ref_no']}-{invoice_suffix(period, idx)}",
        })
        prev = curr
    return records


def normalize_company(company: dict, financial_period: dict, company_index: int) -> dict:
    style = DEFAULT_COMPANY_STYLES[company_index % len(DEFAULT_COMPANY_STYLES)]
    return {
        "label": company["label"],
        "supplier": company["supplier"],
        "supplier_code": company["supplier_code"],
        "supplier_address": company["supplier_address"],
        "customer": company["customer"],
        "customer_code": company["customer_code"],
        "accent": company.get("accent", style["accent"]),
        "accent_soft": company.get("accent_soft", style["accent_soft"]),
        "skew": float(company.get("skew", style["skew"])),
        "currency": company.get("currency", "GBP (£)"),
        "sites": [normalize_site(company, site, financial_period) for site in company["sites"]],
    }


def normalize_site(company: dict, site: dict, financial_period: dict) -> dict:
    site_omit = site.get("_omit", {})
    total_qty = _parse_decimal(site["total_quantity"])
    total_cost = _parse_decimal(site["total_cost"])
    unit = site.get("unit", "kWh")
    raw_label = site.get("label", site["meter_id"])

    supplier_ef_raw = site.get("supplier_ef")
    has_supplier_ef = (
        not site_omit.get("supplier_ef")
        and supplier_ef_raw not in (None, "", "0", 0, "0.0", 0.0)
    )
    supplier_ef = _parse_decimal(supplier_ef_raw) if has_supplier_ef else Decimal("0")

    start_reading = int(site["start_reading"])
    period_start = financial_period["start_date"]
    base_ref_no = f"{company['supplier_code']}-{company['customer_code']}-ELEC-{period_start.year}"

    symbol = currency_symbol(company.get("currency", "GBP (£)"))
    site_uid = f"{company['label']}|{raw_label or site.get('meter_id', '?')}"
    rng_site = random.Random(hash(site_uid) & 0xFFFFFFFF)

    raw_tariffs = site.get("tariffs", [])
    if raw_tariffs and "quantity" not in raw_tariffs[0]:
        tariff_names = [str(t.get("name", "")).strip() for t in raw_tariffs if str(t.get("name", "")).strip()]
        tariffs = _split_among_tariffs(total_qty, _q2(total_cost), tariff_names, rng_site)
        for tariff in tariffs:
            tariff["unit"] = unit
    else:
        tariffs = []
        for tariff in raw_tariffs:
            if not str(tariff.get("name", "")).strip():
                continue
            tariffs.append({
                "name": str(tariff["name"]),
                "quantity": _parse_decimal(tariff.get("quantity", 0)),
                "unit": unit,
                "unit_cost": _parse_decimal(tariff.get("unit_cost", 0)),
                "cost": _q2(_parse_decimal(tariff.get("cost", 0))),
            })

    annual_site = {
        "label": "" if site_omit.get("label", False) else raw_label,
        "customer": site.get("customer", company["customer"]),
        "customer_code": site.get("customer_code", company["customer_code"]),
        "customer_address": site["customer_address"],
        "city": "" if site_omit.get("city", False) else site["city"],
        "postcode": "" if site_omit.get("postcode", False) else site["postcode"],
        "meter_id": site["meter_id"],
        "supplier_ef": supplier_ef,
        "_has_supplier_ef": has_supplier_ef,
        "unit": unit,
        "start_reading": start_reading,
        "total_quantity": total_qty,
        "total_cost": _q2(total_cost),
        "tariffs": tariffs,
        "currency_symbol": symbol,
        "_omit": site_omit,
        "_base_ref_no": base_ref_no,
        "_site_uid": site_uid,
    }

    raw_periods = site.get("billing_periods")
    if raw_periods is not None:
        billing_periods = normalize_billing_periods(raw_periods, financial_period["start_date"].year)
    else:
        billing_periods = derive_month_periods(financial_period["start_date"], financial_period["end_date"])

    rng_site.seed(hash(site_uid) & 0xFFFFFFFF)
    period_records = _generate_period_records(annual_site, billing_periods, rng_site)

    return {**annual_site, "period_records": period_records}


def build_sections(config: dict) -> list[dict]:
    sections = []
    for company in config["companies"]:
        for site in company["sites"]:
            for record in site["period_records"]:
                sections.append({"company": company, "site": record})
    return sections


def render_pdf(config: dict, sections: list[dict], output_path: str, noise_level: float = 1.0):
    from generators.pdf_generator import render_pdf as render_format

    return render_format(config, sections, output_path, category="electricity", noise_level=noise_level)


def generate_xlsx(config: dict, sections: list[dict]) -> bytes:
    from generators.xlsx_generator import generate_xlsx as generate_format

    return generate_format(config, sections, category="electricity")


def generate_csv(config: dict, sections: list[dict]) -> bytes:
    from generators.csv_generator import generate_csv as generate_format

    return generate_format(config, sections, category="electricity")


def generate_docx(config: dict, sections: list[dict]) -> bytes:
    from generators.docx_generator import generate_docx as generate_format

    return generate_format(config, sections, category="electricity")
