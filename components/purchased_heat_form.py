from __future__ import annotations

import hashlib as _hashlib
import json as _json
import os as _os
import random as _random
from calendar import monthrange as _monthrange
from datetime import date

import streamlit as st

from components.purchased_heat.supplier_portal_data import render_supplier_portal_data_form
from components.purchased_heat.utility_bill import render_utility_bill_form
from components.sidebar import NEW_COMPANY_PLACEHOLDER, get_document_type_config
from utils.currency import (
    CURRENCY_DISPLAY as _CURRENCY_DISPLAY,
    currency_code as _currency_code,
    currency_index as _currency_index,
    currency_options as _currency_options,
)

_LANGUAGE_OPTIONS: dict[str, str] = {
    "English": "en",
    "French (Français)": "fr",
    "German (Deutsch)": "de",
    "Dutch (Nederlands)": "nl",
}

_CONFIG_DIR = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "..", "config")


def _load_json(filename: str):
    path = _os.path.join(_CONFIG_DIR, filename)
    try:
        with open(path, encoding="utf-8") as fh:
            return _json.load(fh)
    except (FileNotFoundError, _json.JSONDecodeError):
        return None


_COMPANIES_CONFIG: list = _load_json("companies.json") or []
_SUPPLIERS_CONFIG: dict = _load_json("suppliers.json") or {}
_SITES_CONFIG: list = _load_json("sites.json") or []


def _config_company(i: int) -> dict:
    if not _COMPANIES_CONFIG:
        return {}
    return _COMPANIES_CONFIG[i % len(_COMPANIES_CONFIG)]


def _config_site(i: int, j: int) -> dict:
    if not _SITES_CONFIG:
        return {}
    raw = f"site:{i}:{j}".encode()
    idx = int(_hashlib.sha1(raw).hexdigest()[:4], 16) % len(_SITES_CONFIG)
    return _SITES_CONFIG[idx]


def _config_supplier(scope_type: str, company_idx: int) -> dict:
    suppliers = _SUPPLIERS_CONFIG.get(scope_type, [])
    if not suppliers:
        return {}
    return suppliers[company_idx % len(suppliers)]


def _make_meter_id(scope_type: str, supplier_code: str, city: str, i: int, j: int) -> str:
    raw = f"{scope_type}:{i}:{j}:{city}".encode()
    num = int(_hashlib.sha1(raw).hexdigest()[:4], 16) % 90000 + 10000
    city_abbr = (city[:3] or "XXX").upper()
    type_code = "HT" if scope_type == "heat" else "EL"
    return f"{supplier_code}-{city_abbr}-{type_code}-{num}"


def _make_code(name: str) -> str:
    parts = []
    for word in name.split():
        if word.isupper() and len(word) <= 5:
            parts.append(word)
        else:
            parts.append(word[0].upper())
    return "".join(parts)


def _document_defaults(category: str, document_type: str | None) -> tuple[str, str]:
    cfg = get_document_type_config(category, document_type or "")
    return cfg.get("default_title", "Document"), cfg.get("default_subject", "")


def _sync_document_setting_defaults(category: str, document_type: str | None) -> None:
    selection_key = f"{category}:{document_type or ''}"
    if st.session_state.get("_document_settings_selection") == selection_key:
        return

    default_title, default_subject = _document_defaults(category, document_type)
    st.session_state["doc_title"] = default_title
    st.session_state["doc_subject"] = default_subject
    st.session_state["_document_settings_selection"] = selection_key


def _render_document_settings(category: str, document_type: str | None) -> None:
    _sync_document_setting_defaults(category, document_type)
    default_title, default_subject = _document_defaults(category, document_type)

    with st.expander("Document Settings", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.text_input("Document Title", value=default_title, key="doc_title")
        with col2:
            st.text_input("Document Subject", value=default_subject, key="doc_subject")
            st.number_input(
                "Random Seed",
                value=20260325,
                min_value=0,
                max_value=2**31 - 1,
                step=1,
                key="doc_seed",
                help="Controls randomised meter readings and pricing variations.",
            )

        output_format = st.session_state.get("sidebar_format", "PDF").removesuffix(" (Coming Soon)")

        st.selectbox(
            "Document Language",
            options=list(_LANGUAGE_OPTIONS.keys()),
            key="doc_language_label",
            help="Language used for field labels and headings in the generated document.",
        )

        st.checkbox(
            "Inject special characters",
            key="doc_inject_special_chars",
            help=(
                "Append QA characters (& < \" £ € \\u00a0 — \\u200f) to every non-empty text field "
                "in the generated document. Tests parser robustness."
            ),
        )

        if output_format == "PDF":
            st.slider(
                "Scan noise level",
                min_value=0.0,
                max_value=1.0,
                value=0.0,
                step=0.05,
                key="doc_noise",
                help="Controls background texture and page skew intensity.",
            )

        if document_type == "utility_bill" and output_format in {"PDF", "DOCX"}:
            st.checkbox(
                "Generate one file per month and export as ZIP",
                key="doc_monthly_zip",
                help="Creates one bill document per month for the selected financial period and packages them into a ZIP archive.",
            )

        if output_format == "XLSX" and document_type == "supplier_portal_data":
            st.checkbox(
                "Include summary sheet",
                key="xlsx_include_summary",
                value=False,
                help="Add a summary worksheet alongside the detailed supplier portal export.",
            )

        if output_format == "XLSX":
            st.checkbox(
                "Split into one sheet per company",
                key="xlsx_split_by_company",
                help="Generate a separate billing detail sheet for each company instead of one combined sheet.",
            )


def _rand_financial_period() -> tuple[date, date, str]:
    current_year = 2026
    year = _random.randint(current_year - 4, current_year)
    start_month = _random.randint(1, 12)
    n_months = _random.randint(1, 12)
    start = date(year, start_month, 1)
    end_year = year
    end_month = start_month + n_months - 1
    while end_month > 12:
        end_month -= 12
        end_year += 1
    last_day = _monthrange(end_year, end_month)[1]
    end = date(end_year, end_month, last_day)
    label = f"Financial Year {year}" if end_year == year else f"Financial Period {year}–{end_year}"
    return start, end, label


def _render_financial_period() -> list[tuple[int, int]]:
    if "fp_start" not in st.session_state:
        fp_start_rand, fp_end_rand, fp_label_rand = _rand_financial_period()
        st.session_state["fp_start"] = fp_start_rand
        st.session_state["fp_end"] = fp_end_rand
        st.session_state["fp_label"] = fp_label_rand

    st.markdown("#### Financial Period")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.text_input("Period Label", key="fp_label")
    with col2:
        st.date_input("Start Date", key="fp_start")
    with col3:
        st.date_input("End Date", key="fp_end")

    fp_start: date = st.session_state.get("fp_start", date(2026, 1, 1))
    fp_end: date = st.session_state.get("fp_end", date(2026, 12, 31))
    if fp_end < fp_start:
        st.error("End date must be after start date.")
        return []

    fp_months = _months_in_range(fp_start, fp_end)
    st.caption(f"Period spans {len(fp_months)} month(s).")
    return fp_months


def _init_heat_global_random() -> None:
    if "heat_global_capacity_kw" in st.session_state:
        return
    st.session_state["heat_global_capacity_kw"] = _random.choice(range(80, 505, 5))
    st.session_state["heat_global_capacity_rate"] = round(_random.uniform(4.00, 8.50), 2)
    st.session_state["heat_global_base_consumption"] = _random.choice(range(8000, 46000, 100))
    st.session_state["heat_global_unit_price_base"] = round(_random.uniform(0.050, 0.115), 3)
    st.session_state["heat_global_start_reading"] = _random.choice(range(50000, 1000000, 1000))
    st.session_state["heat_global_supplier_ef"] = round(_random.uniform(0.035, 0.120), 4)
    st.session_state["heat_global_supplier_ef_omit"] = _random.random() < 0.25


def _render_heat_global_config(document_type: str | None) -> None:
    _init_heat_global_random()
    st.markdown("#### Global Heat Configuration")
    st.caption("These defaults apply to every site. Sites can optionally override individual values.")
    gd = _HEAT_GLOBAL_DEFAULTS
    col1, col2 = st.columns(2)
    with col1:
        st.number_input(
            "Contracted Capacity (kW)",
            key="heat_global_capacity_kw",
            value=int(gd["capacity_kw"]),
            min_value=10,
            max_value=2000,
            step=5,
            help="Default contracted capacity applied to all sites.",
        )
        st.number_input(
            f"Capacity Rate ({_heat_currency_code()}/kW/month)",
            key="heat_global_capacity_rate",
            value=float(gd["capacity_rate"]),
            min_value=0.01,
            max_value=50.0,
            step=0.05,
            format="%.2f",
            help="Standing charge per kW of contracted capacity per month.",
        )
        st.number_input(
            "Base Monthly Consumption (kWh)",
            key="heat_global_base_consumption",
            value=int(gd["base_consumption"]),
            min_value=100,
            max_value=500_000,
            step=100,
            help="Monthly heat consumption used unless overridden per site.",
        )
    with col2:
        st.number_input(
            f"Base Unit Price ({_heat_currency_code()}/kWh)",
            key="heat_global_unit_price_base",
            value=float(gd["unit_price_base"]),
            min_value=0.010,
            max_value=0.500,
            step=0.001,
            format="%.3f",
            help="Starting unit price before seasonal adjustments.",
        )
        if _heat_field_can_be_omitted(document_type, "start_reading"):
            _field(
                st.number_input, "Start Meter Reading (kWh)", "heat_global_start_reading",
                value=int(gd["start_reading"]),
                min_value=0, max_value=9_999_999, step=1000,
                help="Opening meter reading used unless overridden per site.",
            )
        else:
            st.number_input(
                "Start Meter Reading (kWh)",
                key="heat_global_start_reading",
                value=int(gd["start_reading"]),
                min_value=0,
                max_value=9_999_999,
                step=1000,
                help="Opening meter reading used unless overridden per site.",
            )
        if _heat_field_can_be_omitted(document_type, "supplier_ef"):
            _field(
                st.number_input, "Supplier Emission Factor (kg CO₂e/kWh)", "heat_global_supplier_ef",
                value=float(gd["supplier_ef"]),
                min_value=0.0, max_value=2.0, step=0.001, format="%.4f",
                help="Supplier-reported emission factor for purchased heat.",
            )
        else:
            st.number_input(
                "Supplier Emission Factor (kg CO₂e/kWh)",
                key="heat_global_supplier_ef",
                value=float(gd["supplier_ef"]),
                min_value=0.0,
                max_value=2.0,
                step=0.001,
                format="%.4f",
                help="Supplier-reported emission factor for purchased heat.",
            )
    st.divider()


_HEAT_SITE_CONSUMPTION: list[list[dict]] = [
    [
        {"meter_id": "ADES-RHL-HT-90533", "capacity_kw": 265, "capacity_rate": 6.15, "base_consumption": 28100, "unit_price_base": 0.079, "start_reading": 915860},
        {"meter_id": "ADES-MAN-HT-90534", "capacity_kw": 180, "capacity_rate": 5.95, "base_consumption": 19400, "unit_price_base": 0.074, "start_reading": 624180},
    ],
    [
        {"meter_id": "BTG-BRU-HT-31021", "capacity_kw": 132, "capacity_rate": 5.35, "base_consumption": 15600, "unit_price_base": 0.067, "start_reading": 351220},
        {"meter_id": "BTG-ANR-HT-31022", "capacity_kw": 118, "capacity_rate": 5.10, "base_consumption": 14100, "unit_price_base": 0.065, "start_reading": 287640},
    ],
    [
        {"meter_id": "BHD-BTS-HT-44201", "capacity_kw": 98, "capacity_rate": 4.80, "base_consumption": 11200, "unit_price_base": 0.061, "start_reading": 198450},
    ],
    [
        {"meter_id": "DTN-DUB-HT-57301", "capacity_kw": 110, "capacity_rate": 5.20, "base_consumption": 12800, "unit_price_base": 0.068, "start_reading": 224780},
    ],
    [
        {"meter_id": "DFF-BAL-HT-62401", "capacity_kw": 88, "capacity_rate": 4.60, "base_consumption": 10100, "unit_price_base": 0.058, "start_reading": 176320},
    ],
]

_SITE_IDENTITY_FIELDS = ["label", "city", "postcode"]


def _co_default(i: int, field: str, fallback: str = "") -> str:
    co = _config_company(i)
    if field in ("label", "customer"):
        return co.get("name", fallback)
    if field == "customer_code":
        return co.get("code", _make_code(co.get("name", "")) or fallback)
    if field == "currency":
        return _CURRENCY_DISPLAY.get(co.get("currency", "EUR"), "EUR (€)")
    supplier = _config_supplier("heat", i)
    if field == "supplier":
        return supplier.get("name", fallback)
    if field == "supplier_code":
        return supplier.get("code", fallback)
    if field == "supplier_address":
        return supplier.get("address", fallback)
    return fallback


def _site_default(i: int, j: int, field: str, fallback=None):
    site = _config_site(i, j)
    if field in ("label", "city", "postcode", "address"):
        return site.get(field, fallback)
    if field == "meter_id":
        if i < len(_HEAT_SITE_CONSUMPTION) and j < len(_HEAT_SITE_CONSUMPTION[i]):
            meter_id = _HEAT_SITE_CONSUMPTION[i][j].get("meter_id")
            if meter_id:
                return meter_id
        supplier = _config_supplier("heat", i)
        return _make_meter_id("heat", supplier.get("code", "SUP"), site.get("city", ""), i, j)
    if i < len(_HEAT_SITE_CONSUMPTION) and j < len(_HEAT_SITE_CONSUMPTION[i]):
        return _HEAT_SITE_CONSUMPTION[i][j].get(field, fallback)
    return fallback


_HEAT_GLOBAL_DEFAULTS = {
    "capacity_kw": 150,
    "capacity_rate": 5.50,
    "base_consumption": 15000,
    "unit_price_base": 0.070,
    "start_reading": 400000,
    "supplier_ef": 0.065,
}

_HEAT_CONSUMPTION_FIELDS = [
    "capacity_kw",
    "capacity_rate",
    "base_consumption",
    "unit_price_base",
    "start_reading",
    "supplier_ef",
]


def _heat_field_can_be_omitted(document_type: str | None, field: str) -> bool:
    if field == "supplier_ef":
        return True
    return document_type == "utility_bill" and field in {"start_reading"}


def _heat_global_default(field: str, fallback=None):
    return _HEAT_GLOBAL_DEFAULTS.get(field, fallback)


def _heat_site_override_val(i: int, j: int, field: str, fallback=None):
    val = _site_default(i, j, field)
    if val is None:
        return _heat_global_default(field, fallback)
    return val


def _heat_currency_code(currency: str | None = None) -> str:
    selected = currency or st.session_state.get("co_0_currency") or _co_default(0, "currency", "GBP (£)")
    return _currency_code(selected)


def _heat_company_currency_code(i: int) -> str:
    return _heat_currency_code(st.session_state.get(f"co_{i}_currency") or _co_default(i, "currency", "GBP (£)"))


def _field(widget_fn, label: str, key: str, omit_default: bool = False, **kwargs) -> None:
    is_omitted: bool = st.session_state.get(f"{key}_omit", omit_default)
    field_col, omit_col = st.columns([8, 1])
    with field_col:
        widget_fn(label, key=key, disabled=is_omitted, **kwargs)
    with omit_col:
        st.checkbox(
            "Omit",
            value=omit_default,
            key=f"{key}_omit",
            help="Leave this field blank in the generated document (QA testing).",
        )


def _render_companies_section(fp_months: list[tuple[int, int]], document_type: str | None) -> None:
    st.markdown("#### Companies")
    n_companies = st.number_input(
        "Number of companies",
        min_value=1,
        max_value=10,
        value=1,
        step=1,
        key="n_companies",
    )
    for i in range(int(n_companies)):
        co_label = st.session_state.get(f"co_{i}_label") or _co_default(i, "label") or f"Company {i + 1}"
        with st.expander(f"Company {i + 1}: {co_label}", expanded=(i == 0)):
            _render_company_form(i, fp_months, document_type)


def _render_company_form(i: int, fp_months: list[tuple[int, int]], document_type: str | None) -> None:
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Company Label", key=f"co_{i}_label", value=_co_default(i, "label"))
        st.text_input("Supplier Name", key=f"co_{i}_supplier", value=_co_default(i, "supplier"))
        st.text_input(
            "Supplier Code",
            key=f"co_{i}_supplier_code",
            value=_co_default(i, "supplier_code"),
            help="Short alphanumeric code used in invoice numbers.",
        )
        st.text_area("Supplier Address", key=f"co_{i}_supplier_address", value=_co_default(i, "supplier_address"), height=104)
    with col2:
        st.text_input("Customer Name", key=f"co_{i}_customer", value=_co_default(i, "customer"))
        st.text_input("Customer Code", key=f"co_{i}_customer_code", value=_co_default(i, "customer_code"))
        st.selectbox(
            "Currency",
            options=_currency_options(),
            index=_currency_index(
                _co_default(i, "currency", _CURRENCY_DISPLAY.get(NEW_COMPANY_PLACEHOLDER["currency"], "EUR (€)"))
            ),
            key=f"co_{i}_currency",
        )
        with st.expander("Advanced Options"):
            st.color_picker("Accent Colour", value="#1E5B88", key=f"co_{i}_accent")

    st.markdown(f"**Sites for Company {i + 1}**")
    n_sites = st.number_input(
        "Number of sites",
        min_value=1,
        max_value=20,
        value=1,
        step=1,
        key=f"n_sites_{i}",
    )
    for j in range(int(n_sites)):
        site_label = st.session_state.get(f"site_{i}_{j}_label") or _site_default(i, j, "label") or f"Site {j + 1}"
        with st.expander(f"Site {j + 1}: {site_label}", expanded=(j == 0 and i == 0)):
            _render_site_form(i, j, fp_months, document_type)


def _render_site_form(i: int, j: int, fp_months: list[tuple[int, int]], document_type: str | None) -> None:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Identity**")
        _field(st.text_input, "Site Label", f"site_{i}_{j}_label", value=_site_default(i, j, "label", ""))
        st.text_area("Customer Address", key=f"site_{i}_{j}_address", value=_site_default(i, j, "address", ""), height=104)
        _field(st.text_input, "City", f"site_{i}_{j}_city", value=_site_default(i, j, "city", ""))
        _field(st.text_input, "Postcode", f"site_{i}_{j}_postcode", value=_site_default(i, j, "postcode", ""))
        st.text_input("Heat Meter ID", key=f"site_{i}_{j}_meter_id", value=_site_default(i, j, "meter_id", ""))
    with col2:
        st.markdown("**Consumption**")
        override = st.checkbox(
            "Override global consumption defaults",
            key=f"site_{i}_{j}_override",
            value=False,
            help="When enabled, use site-specific capacity and consumption values instead of the global defaults.",
        )
        if override:
            st.number_input("Contracted Capacity (kW)", key=f"site_{i}_{j}_capacity_kw", min_value=10, max_value=2000, step=5, value=int(_heat_site_override_val(i, j, "capacity_kw", 150)))
            st.number_input(f"Capacity Rate ({_heat_company_currency_code(i)}/kW/month)", key=f"site_{i}_{j}_capacity_rate", min_value=0.01, max_value=50.0, step=0.05, format="%.2f", value=float(_heat_site_override_val(i, j, "capacity_rate", 5.50)))
            st.number_input("Base Monthly Consumption (kWh)", key=f"site_{i}_{j}_base_consumption", min_value=100, max_value=500_000, step=100, value=int(_heat_site_override_val(i, j, "base_consumption", 15000)))
            st.number_input(f"Base Unit Price ({_heat_company_currency_code(i)}/kWh)", key=f"site_{i}_{j}_unit_price_base", min_value=0.010, max_value=0.500, step=0.001, format="%.3f", value=float(_heat_site_override_val(i, j, "unit_price_base", 0.070)))
            if _heat_field_can_be_omitted(document_type, "start_reading"):
                _field(st.number_input, "Start Meter Reading (kWh)", f"site_{i}_{j}_start_reading", min_value=0, max_value=9_999_999, step=1000, value=int(_heat_site_override_val(i, j, "start_reading", 400000)))
            else:
                st.number_input("Start Meter Reading (kWh)", key=f"site_{i}_{j}_start_reading", min_value=0, max_value=9_999_999, step=1000, value=int(_heat_site_override_val(i, j, "start_reading", 400000)))
            if _heat_field_can_be_omitted(document_type, "supplier_ef"):
                _field(st.number_input, "Supplier Emission Factor (kg CO₂e/kWh)", f"site_{i}_{j}_supplier_ef", min_value=0.0, max_value=2.0, step=0.001, format="%.4f", value=float(_heat_site_override_val(i, j, "supplier_ef", 0.065)))
            else:
                st.number_input("Supplier Emission Factor (kg CO₂e/kWh)", key=f"site_{i}_{j}_supplier_ef", min_value=0.0, max_value=2.0, step=0.001, format="%.4f", value=float(_heat_site_override_val(i, j, "supplier_ef", 0.065)))
        else:
            st.caption("Using global heat configuration values.")

    st.markdown("**Billing Periods**")
    period_mode: str = st.radio(
        "Billing period mode",
        options=["All months in financial period", "Custom months"],
        horizontal=True,
        key=f"site_{i}_{j}_period_mode",
        label_visibility="collapsed",
    )

    if period_mode == "Custom months":
        if not fp_months:
            st.warning("Define a valid financial period first.")
        else:
            month_options = [date(year, month, 1).strftime("%B %Y") for year, month in fp_months]
            st.multiselect(
                "Select billing months",
                options=month_options,
                default=month_options,
                key=f"site_{i}_{j}_months",
            )
    else:
        st.caption(
            f"Will generate {len(fp_months)} monthly billing statement(s) covering the full financial period."
        )


def _months_in_range(start: date, end: date) -> list[tuple[int, int]]:
    months: list[tuple[int, int]] = []
    current = date(start.year, start.month, 1)
    while current <= end:
        months.append((current.year, current.month))
        if current.month == 12:
            current = date(current.year + 1, 1, 1)
        else:
            current = date(current.year, current.month + 1, 1)
    return months


def _collect_form_data(document_type: str | None) -> dict:
    s = st.session_state
    default_title, default_subject = _document_defaults("purchased_heat_steam_cooling", document_type)

    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end: date = s.get("fp_end", date(2026, 12, 31))
    fp_months = _months_in_range(fp_start, fp_end)
    month_label_map = {date(year, month, 1).strftime("%B %Y"): (year, month) for year, month in fp_months}

    g_cap_kw = int(s.get("heat_global_capacity_kw", _heat_global_default("capacity_kw", 150)))
    g_cap_rate = float(s.get("heat_global_capacity_rate", _heat_global_default("capacity_rate", 5.50)))
    g_base_cons = int(s.get("heat_global_base_consumption", _heat_global_default("base_consumption", 15000)))
    g_unit_price = float(s.get("heat_global_unit_price_base", _heat_global_default("unit_price_base", 0.070)))
    g_start_reading = int(s.get("heat_global_start_reading", _heat_global_default("start_reading", 400000)))
    g_supplier_ef = float(s.get("heat_global_supplier_ef", _heat_global_default("supplier_ef", 0.065)))
    g_omit = {
        field: (_heat_field_can_be_omitted(document_type, field) and bool(s.get(f"heat_global_{field}_omit", False)))
        for field in _HEAT_CONSUMPTION_FIELDS
    }

    n_companies = int(s.get("n_companies", 1))
    companies: list[dict] = []

    for i in range(n_companies):
        n_sites = int(s.get(f"n_sites_{i}", 1))
        sites: list[dict] = []

        for j in range(n_sites):
            period_mode = s.get(f"site_{i}_{j}_period_mode", "All months in financial period")
            billing_periods: list[dict] | None = None

            if period_mode == "Custom months":
                selected_labels: list[str] = s.get(f"site_{i}_{j}_months", [])
                billing_periods = [
                    {"year": month_label_map[label][0], "month": month_label_map[label][1]}
                    for label in selected_labels
                    if label in month_label_map
                ]

            override = bool(s.get(f"site_{i}_{j}_override", False))
            if override:
                cap_kw = int(s.get(f"site_{i}_{j}_capacity_kw", g_cap_kw))
                cap_rate = float(s.get(f"site_{i}_{j}_capacity_rate", g_cap_rate))
                base_cons = int(s.get(f"site_{i}_{j}_base_consumption", g_base_cons))
                unit_price = float(s.get(f"site_{i}_{j}_unit_price_base", g_unit_price))
                start_reading = int(s.get(f"site_{i}_{j}_start_reading", g_start_reading))
                supplier_ef = float(s.get(f"site_{i}_{j}_supplier_ef", g_supplier_ef))
                cons_omit = {
                    field: (_heat_field_can_be_omitted(document_type, field) and bool(s.get(f"site_{i}_{j}_{field}_omit", False)))
                    for field in _HEAT_CONSUMPTION_FIELDS
                }
            else:
                cap_kw, cap_rate, base_cons, unit_price, start_reading = (
                    g_cap_kw,
                    g_cap_rate,
                    g_base_cons,
                    g_unit_price,
                    g_start_reading,
                )
                supplier_ef = g_supplier_ef
                cons_omit = dict(g_omit)

            site: dict = {
                "label": s.get(f"site_{i}_{j}_label", "") or f"Site {j + 1}",
                "customer_address": [line for line in s.get(f"site_{i}_{j}_address", "").split("\n") if line.strip()],
                "city": s.get(f"site_{i}_{j}_city", ""),
                "postcode": s.get(f"site_{i}_{j}_postcode", ""),
                "meter_id": s.get(f"site_{i}_{j}_meter_id", ""),
                "capacity_kw": cap_kw,
                "capacity_rate": str(cap_rate),
                "base_consumption": base_cons,
                "unit_price_base": str(unit_price),
                "start_reading": start_reading,
                "supplier_ef": str(supplier_ef),
                "_omit": {
                    **{field: bool(s.get(f"site_{i}_{j}_{field}_omit", False)) for field in _SITE_IDENTITY_FIELDS},
                    **cons_omit,
                },
            }
            if billing_periods is not None:
                site["billing_periods"] = billing_periods

            sites.append(site)

        companies.append({
            "label": s.get(f"co_{i}_label", "") or f"Company {i + 1}",
            "supplier": s.get(f"co_{i}_supplier", ""),
            "supplier_code": s.get(f"co_{i}_supplier_code", ""),
            "supplier_address": [line for line in s.get(f"co_{i}_supplier_address", "").split("\n") if line.strip()],
            "customer": s.get(f"co_{i}_customer", ""),
            "customer_code": s.get(f"co_{i}_customer_code", ""),
            "currency": s.get(f"co_{i}_currency", "GBP (£)"),
            "accent": s.get(f"co_{i}_accent", "#1E5B88"),
            "sites": sites,
            "_omit": {},
        })

    return {
        "_category": "purchased_heat_steam_cooling",
        "document_type": document_type or "utility_bill",
        "doc_title": s.get("doc_title", default_title),
        "doc_subject": s.get("doc_subject", default_subject),
        "doc_seed": int(s.get("doc_seed", 20260325)),
        "fp_label": s.get("fp_label", "Financial Year 2026"),
        "fp_start": fp_start.isoformat(),
        "fp_end": fp_end.isoformat(),
        "doc_language": _LANGUAGE_OPTIONS.get(s.get("doc_language_label", "English"), "en"),
        "doc_noise": float(s.get("doc_noise", 0.0)),
        "doc_monthly_zip": bool(s.get("doc_monthly_zip", False)),
        "doc_inject_special_chars": bool(s.get("doc_inject_special_chars", False)),
        "xlsx_include_summary": bool(s.get("xlsx_include_summary", False)),
        "xlsx_split_by_company": bool(s.get("xlsx_split_by_company", False)),
        "companies": companies,
    }


def render_purchased_heat_form(document_type: str | None) -> dict:
    if document_type == "supplier_portal_data":
        return render_supplier_portal_data_form(
            _render_document_settings,
            _render_financial_period,
            lambda: _render_heat_global_config(document_type),
            lambda fp_months: _render_companies_section(fp_months, document_type),
            _collect_form_data,
        )
    return render_utility_bill_form(
        _render_document_settings,
        _render_financial_period,
        lambda: _render_heat_global_config(document_type),
        lambda fp_months: _render_companies_section(fp_months, document_type),
        _collect_form_data,
    )
