from __future__ import annotations

import hashlib as _hashlib
import json as _json
import os as _os
import random as _random
from calendar import monthrange as _monthrange
from datetime import date

import streamlit as st

from components.electricity.electricity_bill import render_electricity_bill_form
from components.electricity.smart_meter_data import render_smart_meter_data_form
from components.electricity.supplier_portal_data import render_electricity_supplier_portal_data_form
from components.sidebar import NEW_COMPANY_PLACEHOLDER, get_document_type_config

_CURRENCY_DISPLAY: dict[str, str] = {
    "GBP": "GBP (£)",
    "EUR": "EUR (€)",
    "USD": "USD ($)",
    "JPY": "JPY (¥)",
    "DKK": "DKK (kr)",
    "HUF": "HUF (Ft)",
}

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

        if document_type == "electricity_bill" and output_format in {"PDF", "DOCX"}:
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


_ELEC_GRID_REGIONS = [
    ("UK National Grid", 0.2070),
    ("Germany (Bundesnetz)", 0.3800),
    ("France (RTE)", 0.0520),
    ("Belgium (Elia)", 0.1670),
    ("Netherlands (TenneT)", 0.2840),
    ("Ireland (EirGrid)", 0.2950),
    ("Denmark (Energinet)", 0.1430),
    ("Slovakia (SEPS)", 0.1120),
    ("Hungary (MAVIR)", 0.2600),
    ("Japan (TEPCO)", 0.4510),
    ("USA (EPA eGrid avg)", 0.3860),
    ("Norway (Statnett)", 0.0280),
    ("Sweden (SVK)", 0.0450),
    ("Poland (PSE)", 0.7120),
    ("Spain (REE)", 0.1910),
]

_ELEC_TARIFF_POOLS = [
    ["Day Rate", "Night Rate (Economy 7)"],
    ["Peak", "Off-Peak"],
    ["Standard Unit Rate"],
    ["Day", "Evening", "Night"],
    [],
]


def _init_elec_global_random() -> None:
    if "elec_global_start_reading" in st.session_state:
        return
    _, region_ef = _random.choice(_ELEC_GRID_REGIONS)
    ef = round(region_ef * _random.uniform(0.92, 1.08), 4)
    quantity = round(_random.choice(range(25000, 151000, 100)), 2)
    unit_rate = round(_random.uniform(0.16, 0.34), 4)
    cost = round(quantity * unit_rate, 2)
    tariffs = _random.choice(_ELEC_TARIFF_POOLS)
    st.session_state["elec_global_supplier_ef"] = ef
    st.session_state["elec_global_supplier_ef_omit"] = _random.random() < 0.25
    st.session_state["elec_global_start_reading"] = _random.choice(range(5000, 90001, 100))
    st.session_state["elec_global_total_quantity"] = quantity
    st.session_state["elec_global_total_cost"] = cost
    st.session_state["elec_global_n_tariffs"] = len(tariffs)
    for idx, name in enumerate(tariffs):
        st.session_state[f"elec_global_t{idx}_name"] = name


_ELEC_SITE_SUPPLEMENTS: list[list[dict]] = [
    [
        {"meter_id": "OVO-RHL-EL-90533"},
        {
            "meter_id": "OVO-MAN-EL-90534",
            "_override": {
                "start_reading": 28_190,
                "total_quantity": 54_200.0,
                "total_cost": 14_159.60,
            },
            "_tariff_names": [],
        },
    ],
    [
        {
            "meter_id": "LMN-BRU-EL-31021",
            "_override": {
                "supplier_ef": None,
                "start_reading": 31_050,
                "total_quantity": 52_400.0,
                "total_cost": 13_052.80,
                "_omit_fields": ["supplier_ef"],
            },
        },
    ],
    [
        {
            "meter_id": "SEA-BTS-EL-44201",
            "_override": {
                "supplier_ef": 0.1120,
                "start_reading": 19_840,
                "total_quantity": 38_600.0,
                "total_cost": 8_878.00,
            },
        },
    ],
    [
        {
            "meter_id": "EI-DUB-EL-57301",
            "_override": {
                "supplier_ef": 0.2950,
                "start_reading": 22_670,
                "total_quantity": 47_800.0,
                "total_cost": 12_446.30,
            },
        },
    ],
    [
        {
            "meter_id": "ORD-BAL-EL-62401",
            "_override": {
                "supplier_ef": None,
                "start_reading": 14_920,
                "total_quantity": 33_400.0,
                "total_cost": 8_519.40,
                "_omit_fields": ["supplier_ef"],
            },
            "_tariff_names": [],
        },
    ],
]

_ELECTRICITY_GLOBAL_DEFAULTS = {
    "supplier_ef": 0.2070,
    "unit": "kWh",
    "start_reading": 45_280,
    "total_quantity": 87_600.0,
    "total_cost": 22_922.00,
    "tariffs": ["Day Rate", "Night Rate (Economy 7)"],
}

_ELEC_SITE_OPTIONAL_FIELDS = {"label", "city", "postcode"}


def _elec_site_field_can_be_omitted(document_type: str | None, field: str) -> bool:
    if field == "supplier_ef":
        return True
    if field in _ELEC_SITE_OPTIONAL_FIELDS:
        return True
    return document_type == "electricity_bill" and field in {"start_reading", "total_cost"}


def _elec_global_default(field: str, fallback=None):
    return _ELECTRICITY_GLOBAL_DEFAULTS.get(field, fallback)


def _elec_global_tariff_default(idx: int, fallback: str = "") -> str:
    tariffs = _ELECTRICITY_GLOBAL_DEFAULTS.get("tariffs", [])
    return tariffs[idx] if idx < len(tariffs) else fallback


def _elec_co_default(i: int, field: str, fallback: str = "") -> str:
    co = _config_company(i)
    if field in ("label", "customer"):
        return co.get("name", fallback)
    if field == "customer_code":
        return co.get("code", _make_code(co.get("name", "")) or fallback)
    if field == "currency":
        return _CURRENCY_DISPLAY.get(co.get("currency", "EUR"), "EUR (€)")
    supplier = _config_supplier("electricity", i)
    if field == "supplier":
        return supplier.get("name", fallback)
    if field == "supplier_code":
        return supplier.get("code", fallback)
    if field == "supplier_address":
        return supplier.get("address", fallback)
    return fallback


def _elec_site_default(i: int, j: int, field: str, fallback=None):
    site = _config_site(i, j)
    if field in ("label", "city", "postcode", "address"):
        return site.get(field, fallback)
    if field == "meter_id":
        if i < len(_ELEC_SITE_SUPPLEMENTS) and j < len(_ELEC_SITE_SUPPLEMENTS[i]):
            meter_id = _ELEC_SITE_SUPPLEMENTS[i][j].get("meter_id")
            if meter_id:
                return meter_id
        supplier = _config_supplier("electricity", i)
        return _make_meter_id("electricity", supplier.get("code", "SUP"), site.get("city", ""), i, j)
    if i < len(_ELEC_SITE_SUPPLEMENTS) and j < len(_ELEC_SITE_SUPPLEMENTS[i]):
        return _ELEC_SITE_SUPPLEMENTS[i][j].get(field, fallback)
    return fallback


def _elec_site_has_override(i: int, j: int) -> bool:
    return bool(_elec_site_default(i, j, "_override"))


def _elec_site_override_val(i: int, j: int, field: str, fallback=None):
    override = _elec_site_default(i, j, "_override") or {}
    return override.get(field, _elec_global_default(field, fallback))


def _elec_site_override_omit(i: int, j: int, field: str) -> bool:
    override = _elec_site_default(i, j, "_override") or {}
    return field in override.get("_omit_fields", [])


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


def _render_electricity_global_config(document_type: str | None) -> None:
    _init_elec_global_random()
    st.session_state["elec_global_total_quantity_omit"] = False
    st.markdown("#### Global Electricity Configuration")
    st.caption("These defaults apply to every site. Sites can optionally override individual fields.")

    gd = _ELECTRICITY_GLOBAL_DEFAULTS
    col1, col2 = st.columns(2)
    with col1:
        if _elec_site_field_can_be_omitted(document_type, "supplier_ef"):
            _field(
                st.number_input, "Supplier Emission Factor (kg CO₂e/kWh)", "elec_global_supplier_ef",
                value=float(gd["supplier_ef"]),
                min_value=0.0, max_value=2.0, step=0.001, format="%.4f",
                help="Supplier-reported emission factor for purchased electricity.",
            )
        else:
            st.number_input(
                "Supplier Emission Factor (kg CO₂e/kWh)",
                key="elec_global_supplier_ef",
                value=float(gd["supplier_ef"]),
                min_value=0.0,
                max_value=2.0,
                step=0.001,
                format="%.4f",
                help="Supplier-reported emission factor for purchased electricity.",
            )
        st.selectbox(
            "Measurement Unit",
            options=["kWh", "MWh"],
            index=0 if gd["unit"] == "kWh" else 1,
            key="elec_global_unit",
        )
    with col2:
        if _elec_site_field_can_be_omitted(document_type, "start_reading"):
            _field(
                st.number_input, "Default Start Meter Reading", "elec_global_start_reading",
                value=int(gd["start_reading"]),
                min_value=0, max_value=999_999_999, step=100,
                help="Opening meter reading used unless overridden per site.",
            )
        else:
            st.number_input(
                "Default Start Meter Reading",
                key="elec_global_start_reading",
                value=int(gd["start_reading"]),
                min_value=0,
                max_value=999_999_999,
                step=100,
                help="Opening meter reading used unless overridden per site.",
            )
        st.number_input(
            "Default Annual Quantity",
            value=float(gd["total_quantity"]),
            key="elec_global_total_quantity",
            min_value=0.0, step=100.0, format="%.2f",
            help="Total annual electricity consumption used unless overridden per site.",
        )
        if _elec_site_field_can_be_omitted(document_type, "total_cost"):
            _field(
                st.number_input, "Default Annual Cost", "elec_global_total_cost",
                value=float(gd["total_cost"]),
                min_value=0.0, step=10.0, format="%.2f",
                help="Total annual electricity cost used unless overridden per site.",
            )
        else:
            st.number_input(
                "Default Annual Cost",
                key="elec_global_total_cost",
                value=float(gd["total_cost"]),
                min_value=0.0,
                step=10.0,
                format="%.2f",
                help="Total annual electricity cost used unless overridden per site.",
            )

    st.markdown("**Tariff Rates**")
    st.caption("Define tariff names only. Values are randomly split from the annual totals at generation time.")
    n_tariffs_global = st.number_input(
        "Number of tariff rates",
        min_value=0,
        max_value=10,
        value=len(gd["tariffs"]),
        step=1,
        key="elec_global_n_tariffs",
        help="Define shared tariff names here. Sites can select which ones to include.",
    )
    if int(n_tariffs_global) == 0:
        st.caption("No tariffs defined — all sites will show totals only.")
    else:
        for idx in range(int(n_tariffs_global)):
            st.text_input(
                f"Tariff {idx + 1} name",
                key=f"elec_global_t{idx}_name",
                value=_elec_global_tariff_default(idx, f"Tariff {idx + 1}"),
            )

    st.divider()


def _render_electricity_companies_section(fp_months: list[tuple[int, int]], document_type: str | None) -> None:
    st.markdown("#### Companies")
    n_companies = st.number_input(
        "Number of companies",
        min_value=1,
        max_value=10,
        value=1,
        step=1,
        key="elec_n_companies",
    )
    for i in range(int(n_companies)):
        co_label = st.session_state.get(f"elec_co_{i}_label") or _elec_co_default(i, "label") or f"Company {i + 1}"
        with st.expander(f"Company {i + 1}: {co_label}", expanded=(i == 0)):
            _render_electricity_company_form(i, fp_months, document_type)


def _render_electricity_company_form(i: int, fp_months: list[tuple[int, int]], document_type: str | None) -> None:
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Company Label", key=f"elec_co_{i}_label", value=_elec_co_default(i, "label"))
        st.text_input("Supplier Name", key=f"elec_co_{i}_supplier", value=_elec_co_default(i, "supplier"))
        st.text_input(
            "Supplier Code",
            key=f"elec_co_{i}_supplier_code",
            value=_elec_co_default(i, "supplier_code"),
            help="Short alphanumeric code used in reference numbers.",
        )
        st.text_area("Supplier Address", key=f"elec_co_{i}_supplier_address", value=_elec_co_default(i, "supplier_address"), height=104)
    with col2:
        st.text_input("Customer Name", key=f"elec_co_{i}_customer", value=_elec_co_default(i, "customer"))
        st.text_input("Customer Code", key=f"elec_co_{i}_customer_code", value=_elec_co_default(i, "customer_code"))
        with st.expander("Advanced Options"):
            st.text_input(
                "Currency",
                value=_elec_co_default(i, "currency", _CURRENCY_DISPLAY.get(NEW_COMPANY_PLACEHOLDER["currency"], "EUR (€)")),
                key=f"elec_co_{i}_currency",
            )
            st.color_picker("Accent Colour", value="#1E5B88", key=f"elec_co_{i}_accent")

    st.markdown(f"**Sites for Company {i + 1}**")
    n_sites = st.number_input(
        "Number of sites",
        min_value=1,
        max_value=20,
        value=1,
        step=1,
        key=f"elec_n_sites_{i}",
    )
    for j in range(int(n_sites)):
        site_label = st.session_state.get(f"elec_site_{i}_{j}_label") or _elec_site_default(i, j, "label") or f"Site {j + 1}"
        with st.expander(f"Site {j + 1}: {site_label}", expanded=(j == 0 and i == 0)):
            _render_electricity_site_form(i, j, fp_months, document_type)


def _render_electricity_site_form(i: int, j: int, fp_months: list[tuple[int, int]], document_type: str | None) -> None:
    st.session_state[f"elec_site_{i}_{j}_total_quantity_omit"] = False
    col1, col2 = st.columns(2)
    with col1:
        _field(st.text_input, "Site Label", f"elec_site_{i}_{j}_label", value=_elec_site_default(i, j, "label", ""))
        st.text_area("Customer Address", key=f"elec_site_{i}_{j}_address", value=_elec_site_default(i, j, "address", ""), height=104)
        _field(st.text_input, "City", f"elec_site_{i}_{j}_city", value=_elec_site_default(i, j, "city", ""))
        _field(st.text_input, "Postcode", f"elec_site_{i}_{j}_postcode", value=_elec_site_default(i, j, "postcode", ""))
        st.text_input("Electricity Meter ID", key=f"elec_site_{i}_{j}_meter_id", value=_elec_site_default(i, j, "meter_id", ""))

    with col2:
        override_key = f"elec_site_{i}_{j}_override"
        override = st.checkbox(
            "Override global grid & consumption settings",
            value=st.session_state.get(override_key, False),
            key=override_key,
        )
        if override:
            if _elec_site_field_can_be_omitted(document_type, "supplier_ef"):
                _field(
                    st.number_input, "Supplier Emission Factor (kg CO₂e/kWh)",
                    f"elec_site_{i}_{j}_supplier_ef",
                    omit_default=_elec_site_override_omit(i, j, "supplier_ef"),
                    value=float(_elec_site_override_val(i, j, "supplier_ef") or 0.0),
                    min_value=0.0, max_value=2.0, step=0.001, format="%.4f",
                )
            else:
                st.number_input(
                    "Supplier Emission Factor (kg CO₂e/kWh)",
                    key=f"elec_site_{i}_{j}_supplier_ef",
                    value=float(_elec_site_override_val(i, j, "supplier_ef") or 0.0),
                    min_value=0.0,
                    max_value=2.0,
                    step=0.001,
                    format="%.4f",
                )
            st.selectbox(
                "Measurement Unit",
                options=["kWh", "MWh"],
                index=0 if _elec_site_override_val(i, j, "unit", "kWh") == "kWh" else 1,
                key=f"elec_site_{i}_{j}_unit",
            )
            if _elec_site_field_can_be_omitted(document_type, "start_reading"):
                _field(
                    st.number_input, "Start Meter Reading", f"elec_site_{i}_{j}_start_reading",
                    value=int(_elec_site_override_val(i, j, "start_reading", 0)),
                    min_value=0, max_value=999_999_999, step=100,
                )
            else:
                st.number_input(
                    "Start Meter Reading",
                    key=f"elec_site_{i}_{j}_start_reading",
                    value=int(_elec_site_override_val(i, j, "start_reading", 0)),
                    min_value=0,
                    max_value=999_999_999,
                    step=100,
                )
            st.number_input(
                "Annual Quantity",
                value=float(_elec_site_override_val(i, j, "total_quantity", 0.0)),
                key=f"elec_site_{i}_{j}_total_quantity",
                min_value=0.0, step=100.0, format="%.2f",
            )
            if _elec_site_field_can_be_omitted(document_type, "total_cost"):
                _field(
                    st.number_input, "Annual Cost", f"elec_site_{i}_{j}_total_cost",
                    value=float(_elec_site_override_val(i, j, "total_cost", 0.0)),
                    min_value=0.0, step=10.0, format="%.2f",
                )
            else:
                st.number_input(
                    "Annual Cost",
                    key=f"elec_site_{i}_{j}_total_cost",
                    value=float(_elec_site_override_val(i, j, "total_cost", 0.0)),
                    min_value=0.0,
                    step=10.0,
                    format="%.2f",
                )
        else:
            g_qty = st.session_state.get("elec_global_total_quantity", _ELECTRICITY_GLOBAL_DEFAULTS["total_quantity"])
            g_cost = st.session_state.get("elec_global_total_cost", _ELECTRICITY_GLOBAL_DEFAULTS["total_cost"])
            st.caption(f"Using global defaults — {float(g_qty):,.0f} kWh, £{float(g_cost):,.2f}.")

    session = st.session_state
    global_n_tariffs = int(session.get("elec_global_n_tariffs", len(_ELECTRICITY_GLOBAL_DEFAULTS["tariffs"])))
    global_tariff_names = [
        session.get(f"elec_global_t{idx}_name", _elec_global_tariff_default(idx, f"Tariff {idx + 1}"))
        for idx in range(global_n_tariffs)
        if session.get(f"elec_global_t{idx}_name", _elec_global_tariff_default(idx, "")).strip()
    ]

    if global_tariff_names:
        st.markdown("**Tariffs to include**")
        sample = _elec_site_default(i, j, "_tariff_names")
        if sample is None:
            default_selection = global_tariff_names
        else:
            default_selection = [name for name in global_tariff_names if name in sample] if sample else []

        tariff_mode: str = st.radio(
            "Tariff mode",
            options=["All global tariffs", "Custom selection"],
            horizontal=True,
            key=f"elec_site_{i}_{j}_tariff_mode",
            label_visibility="collapsed",
        )
        if tariff_mode == "Custom selection":
            st.multiselect(
                "Select tariffs",
                options=global_tariff_names,
                default=[name for name in default_selection if name in global_tariff_names],
                key=f"elec_site_{i}_{j}_tariffs",
            )
        else:
            st.caption(f"All {len(global_tariff_names)} global tariff(s) will be included.")
    else:
        st.caption("No global tariffs defined — output will show totals only.")

    st.markdown("**Billing Periods**")
    period_mode: str = st.radio(
        "Billing period mode",
        options=["All months in financial period", "Custom months"],
        horizontal=True,
        key=f"elec_site_{i}_{j}_period_mode",
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
                key=f"elec_site_{i}_{j}_months",
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


def _collect_electricity_form_data(document_type: str | None) -> dict:
    s = st.session_state
    default_title, default_subject = _document_defaults("electricity", document_type)

    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end: date = s.get("fp_end", date(2026, 12, 31))
    fp_months = _months_in_range(fp_start, fp_end)
    month_label_map = {date(year, month, 1).strftime("%B %Y"): (year, month) for year, month in fp_months}

    g_supplier_ef = s.get("elec_global_supplier_ef", _ELECTRICITY_GLOBAL_DEFAULTS["supplier_ef"])
    g_supplier_ef_omit = document_type == "electricity_bill" and bool(s.get("elec_global_supplier_ef_omit", False))
    g_unit = s.get("elec_global_unit", _ELECTRICITY_GLOBAL_DEFAULTS["unit"])
    g_start_reading = s.get("elec_global_start_reading", _ELECTRICITY_GLOBAL_DEFAULTS["start_reading"])
    g_start_reading_omit = document_type == "electricity_bill" and bool(s.get("elec_global_start_reading_omit", False))
    g_total_quantity = s.get("elec_global_total_quantity", _ELECTRICITY_GLOBAL_DEFAULTS["total_quantity"])
    g_total_qty_omit = False
    g_total_cost = s.get("elec_global_total_cost", _ELECTRICITY_GLOBAL_DEFAULTS["total_cost"])
    g_total_cost_omit = document_type == "electricity_bill" and bool(s.get("elec_global_total_cost_omit", False))

    global_n_tariffs = int(s.get("elec_global_n_tariffs", len(_ELECTRICITY_GLOBAL_DEFAULTS["tariffs"])))
    global_tariffs: list[dict] = []
    for idx in range(global_n_tariffs):
        name = s.get(f"elec_global_t{idx}_name", _elec_global_tariff_default(idx, "")).strip()
        if name:
            global_tariffs.append({"name": name})
    global_tariff_names = [tariff["name"] for tariff in global_tariffs]

    n_companies = int(s.get("elec_n_companies", 1))
    companies: list[dict] = []

    for i in range(n_companies):
        n_sites = int(s.get(f"elec_n_sites_{i}", 1))
        sites: list[dict] = []

        for j in range(n_sites):
            has_override = bool(s.get(f"elec_site_{i}_{j}_override", _elec_site_has_override(i, j)))
            if has_override:
                supplier_ef_raw = s.get(f"elec_site_{i}_{j}_supplier_ef", _elec_site_override_val(i, j, "supplier_ef", g_supplier_ef))
                supplier_ef_omit = (
                    document_type == "electricity_bill"
                    and bool(s.get(f"elec_site_{i}_{j}_supplier_ef_omit", _elec_site_override_omit(i, j, "supplier_ef")))
                )
                unit = s.get(f"elec_site_{i}_{j}_unit", _elec_site_override_val(i, j, "unit", g_unit))
                start_reading = s.get(f"elec_site_{i}_{j}_start_reading", _elec_site_override_val(i, j, "start_reading", g_start_reading))
                start_reading_omit = document_type == "electricity_bill" and bool(s.get(f"elec_site_{i}_{j}_start_reading_omit", False))
                total_quantity = s.get(f"elec_site_{i}_{j}_total_quantity", _elec_site_override_val(i, j, "total_quantity", g_total_quantity))
                total_qty_omit = False
                total_cost = s.get(f"elec_site_{i}_{j}_total_cost", _elec_site_override_val(i, j, "total_cost", g_total_cost))
                total_cost_omit = document_type == "electricity_bill" and bool(s.get(f"elec_site_{i}_{j}_total_cost_omit", False))
            else:
                supplier_ef_raw = g_supplier_ef
                supplier_ef_omit = g_supplier_ef_omit
                unit = g_unit
                start_reading = g_start_reading
                start_reading_omit = g_start_reading_omit
                total_quantity = g_total_quantity
                total_qty_omit = g_total_qty_omit
                total_cost = g_total_cost
                total_cost_omit = g_total_cost_omit

            tariff_mode = s.get(f"elec_site_{i}_{j}_tariff_mode", "All global tariffs")
            if tariff_mode == "Custom selection" and global_tariffs:
                selected = set(s.get(f"elec_site_{i}_{j}_tariffs", global_tariff_names))
                site_tariffs = [tariff for tariff in global_tariffs if tariff["name"] in selected]
            elif not global_tariffs:
                site_tariffs = []
            else:
                sample_names = _elec_site_default(i, j, "_tariff_names")
                if sample_names is not None:
                    site_tariffs = [tariff for tariff in global_tariffs if tariff["name"] in sample_names]
                else:
                    site_tariffs = global_tariffs

            period_mode = s.get(f"elec_site_{i}_{j}_period_mode", "All months in financial period")
            billing_periods: list[dict] | None = None
            if period_mode == "Custom months":
                selected_labels: list[str] = s.get(f"elec_site_{i}_{j}_months", [])
                billing_periods = [
                    {"year": month_label_map[label][0], "month": month_label_map[label][1]}
                    for label in selected_labels
                    if label in month_label_map
                ]

            sites.append({
                "label": s.get(f"elec_site_{i}_{j}_label", "") or f"Site {j + 1}",
                "customer_address": [line for line in s.get(f"elec_site_{i}_{j}_address", "").split("\n") if line.strip()],
                "city": s.get(f"elec_site_{i}_{j}_city", ""),
                "postcode": s.get(f"elec_site_{i}_{j}_postcode", ""),
                "meter_id": s.get(f"elec_site_{i}_{j}_meter_id", ""),
                "supplier_ef": str(supplier_ef_raw) if supplier_ef_raw is not None else "0",
                "unit": unit,
                "start_reading": int(start_reading),
                "total_quantity": str(total_quantity),
                "total_cost": str(total_cost),
                "tariffs": site_tariffs,
                "_omit": {
                    **{
                        field: bool(s.get(f"elec_site_{i}_{j}_{field}_omit", False))
                        for field in _ELEC_SITE_OPTIONAL_FIELDS
                    },
                    "supplier_ef": supplier_ef_omit,
                    "start_reading": start_reading_omit,
                    "total_quantity": total_qty_omit,
                    "total_cost": total_cost_omit,
                },
                **({"billing_periods": billing_periods} if billing_periods is not None else {}),
            })

        companies.append({
            "label": s.get(f"elec_co_{i}_label", "") or f"Company {i + 1}",
            "supplier": s.get(f"elec_co_{i}_supplier", ""),
            "supplier_code": s.get(f"elec_co_{i}_supplier_code", ""),
            "supplier_address": [line for line in s.get(f"elec_co_{i}_supplier_address", "").split("\n") if line.strip()],
            "customer": s.get(f"elec_co_{i}_customer", ""),
            "customer_code": s.get(f"elec_co_{i}_customer_code", ""),
            "currency": s.get(f"elec_co_{i}_currency", "GBP (£)"),
            "accent": s.get(f"elec_co_{i}_accent", "#1E5B88"),
            "sites": sites,
            "_omit": {},
        })

    return {
        "_category": "electricity",
        "document_type": document_type or "electricity_bill",
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
        "smart_meter_data_granularity": s.get("smart_meter_data_granularity_label", "Monthly").lower(),
        "smart_meter_interval_minutes": int(str(s.get("smart_meter_interval_label", "30 minutes")).split()[0]),
        "xlsx_include_summary": bool(s.get("xlsx_include_summary", False)),
        "xlsx_split_by_company": bool(s.get("xlsx_split_by_company", False)),
        "companies": companies,
    }


def render_electricity_form(document_type: str | None) -> dict:
    if document_type == "smart_meter_data":
        return render_smart_meter_data_form(document_type)
    if document_type == "supplier_portal_data":
        return render_electricity_supplier_portal_data_form(
            _render_document_settings,
            _render_financial_period,
            lambda: _render_electricity_global_config(document_type),
            lambda fp_months: _render_electricity_companies_section(fp_months, document_type),
            _collect_electricity_form_data,
        )
    return render_electricity_bill_form(
        _render_document_settings,
        _render_financial_period,
        lambda: _render_electricity_global_config(document_type),
        lambda fp_months: _render_electricity_companies_section(fp_months, document_type),
        _collect_electricity_form_data,
    )
