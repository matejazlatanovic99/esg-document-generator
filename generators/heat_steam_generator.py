from __future__ import annotations

import random
from datetime import timedelta
from decimal import Decimal, ROUND_HALF_UP

from generators.shared_generator import (
    DEFAULT_COMPANY_STYLES,
    billing_period_dates,
    billing_period_factor,
    billing_period_label,
    derive_month_periods,
    invoice_suffix,
    normalize_billing_periods,
    parse_decimal,
    q2,
)


def normalize_company(company, financial_period, company_index):
    sites = company.get("sites", [])
    if not sites:
        raise ValueError(f"Company {company.get('label', '<unknown>')} must include at least one site")

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
        "sites": [normalize_site(company, site, financial_period) for site in sites],
    }


def normalize_site(company, site, financial_period):
    billing_periods = site.get("billing_periods", company.get("billing_periods"))
    if billing_periods is None:
        billing_periods = derive_month_periods(financial_period["start_date"], financial_period["end_date"])

    normalized_periods = normalize_billing_periods(billing_periods, financial_period["start_date"].year)
    billing_period_count = site.get("billing_period_count", company.get("billing_period_count"))
    if billing_period_count is not None:
        normalized_periods = normalized_periods[: int(billing_period_count)]

    if not normalized_periods:
        raise ValueError(f"Site {site.get('label', site.get('meter_id', '<unknown>'))} has no billing periods")

    return {
        "label": site.get("label", site["meter_id"]),
        "customer": site.get("customer", company["customer"]),
        "customer_code": site.get("customer_code", company["customer_code"]),
        "customer_address": site["customer_address"],
        "city": site["city"],
        "postcode": site["postcode"],
        "meter_id": site["meter_id"],
        "capacity_kw": int(site["capacity_kw"]),
        "capacity_rate": parse_decimal(site["capacity_rate"]),
        "supplier_ef": parse_decimal(site.get("supplier_ef", "0")),
        "base_consumption": int(site["base_consumption"]),
        "unit_price_base": parse_decimal(site["unit_price_base"]),
        "start_reading": int(site["start_reading"]),
        "billing_periods": normalized_periods,
    }


def generate_billing_records(company, site):
    prev = site["start_reading"]
    records = []
    for index, period in enumerate(site["billing_periods"], start=1):
        first, last = billing_period_dates(period)
        factor = billing_period_factor(period)
        variation = parse_decimal(1 + random.uniform(-0.05, 0.05))
        consumption = int(
            (parse_decimal(site["base_consumption"]) * factor * variation).quantize(
                Decimal("1"),
                rounding=ROUND_HALF_UP,
            )
        )
        consumption = max(5000, min(50000, consumption))
        curr = prev + consumption

        midpoint = first + timedelta(days=((last - first).days // 2))
        seasonal = parse_decimal("0.004") if midpoint.month in (1, 2, 11, 12) else parse_decimal("0.000")
        summer = parse_decimal("-0.002") if midpoint.month in (6, 7, 8) else parse_decimal("0.000")
        random_adjust = parse_decimal(round(random.uniform(-0.003, 0.003), 4))
        unit_price = site["unit_price_base"] + seasonal + summer + random_adjust
        unit_price = min(parse_decimal("0.120"), max(parse_decimal("0.040"), unit_price)).quantize(
            Decimal("1.000"),
            rounding=ROUND_HALF_UP,
        )

        heat_cost = q2(parse_decimal(consumption) * unit_price)
        capacity_charge = q2(parse_decimal(site["capacity_kw"]) * site["capacity_rate"])
        subtotal = q2(heat_cost + capacity_charge)
        vat = q2(subtotal * parse_decimal("0.05"))
        total = q2(subtotal + vat)

        issue_date = last + timedelta(days=4)
        due_date = issue_date + timedelta(days=14)
        invoice_no = f"{company['supplier_code']}-{site['customer_code']}-{invoice_suffix(period, index)}"

        records.append({
            "billing_period_label": billing_period_label(period),
            "supplier": company["supplier"],
            "customer": site["customer"],
            "site_label": site["label"],
            "city": site["city"],
            "postcode": site["postcode"],
            "period_start": first,
            "period_end": last,
            "meter_id": site["meter_id"],
            "prev_read": prev,
            "curr_read": curr,
            "consumption": consumption,
            "unit_price": unit_price,
            "heat_cost": heat_cost,
            "capacity_kw": site["capacity_kw"],
            "capacity_rate": site["capacity_rate"],
            "supplier_ef": site["supplier_ef"],
            "capacity_charge": capacity_charge,
            "subtotal": subtotal,
            "vat": vat,
            "total": total,
            "invoice_no": invoice_no,
            "issue_date": issue_date,
            "due_date": due_date,
        })
        prev = curr
    return records


def validate_records(records):
    for record in records:
        assert record["curr_read"] - record["prev_read"] == record["consumption"]
        assert 5000 <= record["consumption"] <= 50000
        assert parse_decimal("0.04") <= record["unit_price"] <= parse_decimal("0.12")
        assert 50 <= record["capacity_kw"] <= 500
        assert record["heat_cost"] == q2(parse_decimal(record["consumption"]) * record["unit_price"])
        assert record["capacity_charge"] == q2(parse_decimal(record["capacity_kw"]) * record["capacity_rate"])
        assert record["subtotal"] == q2(record["heat_cost"] + record["capacity_charge"])
        assert record["vat"] == q2(record["subtotal"] * parse_decimal("0.05"))
        assert record["total"] == q2(record["subtotal"] + record["vat"])


def build_sections(config):
    random.seed(config["random_seed"])
    sections = []
    for company in config["companies"]:
        for site in company["sites"]:
            records = generate_billing_records(company, site)
            validate_records(records)
            sections.append({
                "company": company,
                "site": site,
                "records": records,
            })
    return sections


def slugify(value):
    normalized = "".join(ch.lower() if ch.isalnum() else "-" for ch in value)
    collapsed = "-".join(part for part in normalized.split("-") if part)
    return collapsed or "company"


def filtered_config(config, company_labels):
    if not company_labels:
        return config

    available_labels = {company["label"] for company in config["companies"]}
    missing = [label for label in company_labels if label not in available_labels]
    if missing:
        raise ValueError(f"Unknown company labels: {', '.join(missing)}")

    filtered = dict(config)
    filtered["companies"] = [company for company in config["companies"] if company["label"] in company_labels]
    return filtered


def output_path_for_company(config, company):
    import os

    pdf_filename = config["document"]["pdf_filename"]
    base, ext = os.path.splitext(pdf_filename)
    filename = f"{base}-{slugify(company['label'])}{ext or '.pdf'}"
    return os.path.join(config["document"]["output_dir"], filename)
