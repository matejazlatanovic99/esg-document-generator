from __future__ import annotations

import hashlib as _hashlib
import json as _json
import os as _os
import random as _random
from calendar import monthrange as _monthrange
from datetime import date

import streamlit as st

from components.sidebar import NEW_COMPANY_PLACEHOLDER

_CURRENCY_DISPLAY: dict[str, str] = {
    "GBP": "GBP (£)",
    "EUR": "EUR (€)",
    "USD": "USD ($)",
    "JPY": "JPY (¥)",
    "DKK": "DKK (kr)",
    "HUF": "HUF (Ft)",
}

# ---------------------------------------------------------------------------
# External config: companies.json + suppliers.json
# ---------------------------------------------------------------------------

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
    """Return companies.json entry for company index i (wraps around)."""
    if not _COMPANIES_CONFIG:
        return {}
    return _COMPANIES_CONFIG[i % len(_COMPANIES_CONFIG)]


def _config_site(i: int, j: int) -> dict:
    """Return a deterministic-random site from sites.json, keyed by (company_idx, site_idx)."""
    if not _SITES_CONFIG:
        return {}
    raw = f"site:{i}:{j}".encode()
    idx = int(_hashlib.sha1(raw).hexdigest()[:4], 16) % len(_SITES_CONFIG)
    return _SITES_CONFIG[idx]


def _config_supplier(scope_type: str, company_idx: int) -> dict:
    """Pick a supplier deterministically by company index from suppliers.json."""
    suppliers = _SUPPLIERS_CONFIG.get(scope_type, [])
    if not suppliers:
        return {}
    return suppliers[company_idx % len(suppliers)]


def _make_meter_id(scope_type: str, supplier_code: str, city: str, i: int, j: int) -> str:
    """Generate a stable deterministic meter ID from scope + supplier + location + indices."""
    raw = f"{scope_type}:{i}:{j}:{city}".encode()
    num = int(_hashlib.sha1(raw).hexdigest()[:4], 16) % 90000 + 10000
    city_abbr = (city[:3] or "XXX").upper()
    type_code = "HT" if scope_type == "heat" else "EL"
    return f"{supplier_code}-{city_abbr}-{type_code}-{num}"


def _make_code(name: str) -> str:
    parts = []
    for w in name.split():
        if w.isupper() and len(w) <= 5:
            parts.append(w)
        else:
            parts.append(w[0].upper())
    return "".join(parts)

_LANGUAGE_OPTIONS: dict[str, str] = {
    "English":            "en",
    "French (Français)":  "fr",
    "German (Deutsch)":   "de",
    "Dutch (Nederlands)": "nl",
}


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def render_scope_form(scope: str, category: str) -> dict | None:
    """Render the appropriate form. Returns form_data dict or None if not implemented."""
    if category == "purchased_heat_steam_cooling":
        return _render_purchased_heat_form()
    if category == "electricity":
        return _render_electricity_form()
    _render_coming_soon(scope, category)
    return None


# ---------------------------------------------------------------------------
# Coming-soon placeholder
# ---------------------------------------------------------------------------

def _render_coming_soon(scope: str, category: str) -> None:
    label = category.replace("_", " ").title()
    st.info(
        f"**{label}** ({scope}) is not yet available.\n\n"
        "This generator currently supports **Scope 2 – Purchased Heat / Steam / Cooling**. "
        "Additional scopes and categories are planned for future releases.",
        icon="🚧",
    )


# ---------------------------------------------------------------------------
# Scope 2 / Purchased Heat form
# ---------------------------------------------------------------------------

def _render_purchased_heat_form() -> dict:
    st.subheader("Purchased Heat / Steam / Cooling")
    st.caption("District heating and cooling billing document configuration.")

    _render_document_settings()
    fp_months = _render_financial_period()
    _render_heat_global_config()
    _render_companies_section(fp_months)

    return _collect_form_data()


def _render_document_settings() -> None:
    with st.expander("Document Settings", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.text_input(
                "Document Title",
                value="District Heating Billing Statement",
                key="doc_title",
            )
        with col2:
            st.text_input(
                "Document Subject",
                value="Purchased Heat billing statements",
                key="doc_subject",
            )
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
            help="Language used for field labels and headings in the generated PDF.",
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
                value=0.3,
                step=0.05,
                key="doc_noise",
                help="Controls background texture and page skew intensity. 0.0 = clean/flat, 1.0 = maximum scan noise.",
            )

        if output_format == "XLSX":
            st.checkbox(
                "Split into one sheet per company",
                key="xlsx_split_by_company",
                help="Generate a separate billing detail sheet for each company instead of one combined sheet.",
            )


def _rand_financial_period() -> tuple[date, date, str]:
    """Return a randomised (start, end, label) within the last 5 calendar years."""
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
    """Render financial period widgets and return list of (year, month) tuples."""
    # Randomize on first page load only (before widgets initialise session state)
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


# ---------------------------------------------------------------------------
# Random global config initialisers — called once per session before widgets
# ---------------------------------------------------------------------------

def _init_heat_global_random() -> None:
    """Pre-populate heat global session state keys with randomised values on first load."""
    if "heat_global_capacity_kw" in st.session_state:
        return
    st.session_state["heat_global_capacity_kw"] = _random.choice(range(80, 505, 5))
    st.session_state["heat_global_capacity_rate"] = round(_random.uniform(4.00, 8.50), 2)
    st.session_state["heat_global_base_consumption"] = _random.choice(range(8000, 46000, 100))
    st.session_state["heat_global_unit_price_base"] = round(_random.uniform(0.050, 0.115), 3)
    st.session_state["heat_global_start_reading"] = _random.choice(range(50000, 1000000, 1000))
    st.session_state["heat_global_supplier_ef"] = round(_random.uniform(0.035, 0.120), 4)
    st.session_state["heat_global_supplier_ef_omit"] = _random.random() < 0.25


_ELEC_GRID_REGIONS = [
    ("UK National Grid",      0.2070),
    ("Germany (Bundesnetz)",  0.3800),
    ("France (RTE)",          0.0520),
    ("Belgium (Elia)",        0.1670),
    ("Netherlands (TenneT)",  0.2840),
    ("Ireland (EirGrid)",     0.2950),
    ("Denmark (Energinet)",   0.1430),
    ("Slovakia (SEPS)",       0.1120),
    ("Hungary (MAVIR)",       0.2600),
    ("Japan (TEPCO)",         0.4510),
    ("USA (EPA eGrid avg)",   0.3860),
    ("Norway (Statnett)",     0.0280),
    ("Sweden (SVK)",          0.0450),
    ("Poland (PSE)",          0.7120),
    ("Spain (REE)",           0.1910),
]

_ELEC_TARIFF_POOLS = [
    ["Day Rate", "Night Rate (Economy 7)"],
    ["Peak", "Off-Peak"],
    ["Standard Unit Rate"],
    ["Day", "Evening", "Night"],
    [],
]


def _init_elec_global_random() -> None:
    """Pre-populate electricity global session state keys with randomised values on first load."""
    if "elec_global_start_reading" in st.session_state:
        return
    region_name, region_ef = _random.choice(_ELEC_GRID_REGIONS)
    # Add ±8% noise to emission factor
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
    for k, name in enumerate(tariffs):
        st.session_state[f"elec_global_t{k}_name"] = name


def _render_heat_global_config() -> None:
    """Global capacity / consumption defaults inherited by all heat sites."""
    _init_heat_global_random()
    st.markdown("#### Global Heat Configuration")
    st.caption(
        "These defaults apply to every site. Sites can optionally override individual values."
    )
    gd = _HEAT_GLOBAL_DEFAULTS
    col1, col2 = st.columns(2)
    with col1:
        _field(
            st.number_input, "Contracted Capacity (kW)", "heat_global_capacity_kw",
            value=int(gd["capacity_kw"]),
            min_value=10, max_value=2000, step=5,
            help="Default contracted capacity applied to all sites.",
        )
        _field(
            st.number_input, "Capacity Rate (£/kW/month)", "heat_global_capacity_rate",
            value=float(gd["capacity_rate"]),
            min_value=0.01, max_value=50.0, step=0.05, format="%.2f",
            help="Standing charge per kW of contracted capacity per month.",
        )
        _field(
            st.number_input, "Base Monthly Consumption (kWh)", "heat_global_base_consumption",
            value=int(gd["base_consumption"]),
            min_value=100, max_value=500_000, step=100,
            help="Monthly heat consumption used unless overridden per site.",
        )
    with col2:
        _field(
            st.number_input, "Base Unit Price (£/kWh)", "heat_global_unit_price_base",
            value=float(gd["unit_price_base"]),
            min_value=0.010, max_value=0.500, step=0.001, format="%.3f",
            help="Starting unit price before seasonal adjustments.",
        )
        _field(
            st.number_input, "Start Meter Reading (kWh)", "heat_global_start_reading",
            value=int(gd["start_reading"]),
            min_value=0, max_value=9_999_999, step=1000,
            help="Opening meter reading used unless overridden per site.",
        )
        _field(
            st.number_input, "Supplier Emission Factor (kg CO₂e/kWh)", "heat_global_supplier_ef",
            value=float(gd["supplier_ef"]),
            min_value=0.0, max_value=2.0, step=0.001, format="%.4f",
            help="Supplier-reported emission factor for purchased heat.",
        )
    st.divider()


def _render_companies_section(fp_months: list[tuple[int, int]]) -> None:
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
            _render_company_form(i, fp_months)


# ---------------------------------------------------------------------------
# Heat: per-demo-site consumption overrides (parallel to companies.json order)
# Identity fields (label, address, city, postcode) come from companies.json.
# ---------------------------------------------------------------------------

_HEAT_SITE_CONSUMPTION: list[list[dict]] = [
    [  # TFS UK (companies.json index 0)
        {"meter_id": "ADES-RHL-HT-90533", "capacity_kw": 265, "capacity_rate": 6.15, "base_consumption": 28100, "unit_price_base": 0.079, "start_reading": 915860},
        {"meter_id": "ADES-MAN-HT-90534", "capacity_kw": 180, "capacity_rate": 5.95, "base_consumption": 19400, "unit_price_base": 0.074, "start_reading": 624180},
    ],
    [  # TFS Belgium
        {"meter_id": "BTG-BRU-HT-31021", "capacity_kw": 132, "capacity_rate": 5.35, "base_consumption": 15600, "unit_price_base": 0.067, "start_reading": 351220},
        {"meter_id": "BTG-ANR-HT-31022", "capacity_kw": 118, "capacity_rate": 5.10, "base_consumption": 14100, "unit_price_base": 0.065, "start_reading": 287640},
    ],
    [  # TFS Slovakia
        {"meter_id": "BHD-BTS-HT-44201", "capacity_kw": 98, "capacity_rate": 4.80, "base_consumption": 11200, "unit_price_base": 0.061, "start_reading": 198450},
    ],
    [  # TFS Ireland
        {"meter_id": "DTN-DUB-HT-57301", "capacity_kw": 110, "capacity_rate": 5.20, "base_consumption": 12800, "unit_price_base": 0.068, "start_reading": 224780},
    ],
    [  # TFS Denmark
        {"meter_id": "DFF-BAL-HT-62401", "capacity_kw": 88, "capacity_rate": 4.60, "base_consumption": 10100, "unit_price_base": 0.058, "start_reading": 176320},
    ],
]

_CO_OMIT_FIELDS = ["label", "supplier", "supplier_code", "supplier_address", "customer", "customer_code"]
# Only identity fields can be omitted at site level; consumption fields tracked separately
_SITE_IDENTITY_FIELDS = ["label", "address", "city", "postcode", "meter_id"]


def _co_default(i: int, field: str, fallback: str = "") -> str:
    """Return pre-filled value from companies.json / suppliers.json ('heat' scope)."""
    co = _config_company(i)
    if field in ("label", "customer"):
        return co.get("name", fallback)
    if field == "customer_code":
        return co.get("code", _make_code(co.get("name", "")) or fallback)
    if field == "currency":
        return _CURRENCY_DISPLAY.get(co.get("currency", "EUR"), "EUR (€)")
    sup = _config_supplier("heat", i)
    if field == "supplier":
        return sup.get("name", fallback)
    if field == "supplier_code":
        return sup.get("code", fallback)
    if field == "supplier_address":
        return sup.get("address", fallback)
    return fallback


def _site_default(i: int, j: int, field: str, fallback=None):
    """Return identity fields from companies.json; consumption fields from _HEAT_SITE_CONSUMPTION."""
    site = _config_site(i, j)
    if field in ("label", "city", "postcode", "address"):
        return site.get(field, fallback)
    if field == "meter_id":
        # Use hardcoded demo ID if available; otherwise generate
        if i < len(_HEAT_SITE_CONSUMPTION) and j < len(_HEAT_SITE_CONSUMPTION[i]):
            mid = _HEAT_SITE_CONSUMPTION[i][j].get("meter_id")
            if mid:
                return mid
        sup = _config_supplier("heat", i)
        return _make_meter_id("heat", sup.get("code", "SUP"), site.get("city", ""), i, j)
    # Consumption fields from demo overrides for known sites
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
    "capacity_kw", "capacity_rate", "base_consumption", "unit_price_base", "start_reading",
    "supplier_ef",
]


def _heat_global_default(field: str, fallback=None):
    return _HEAT_GLOBAL_DEFAULTS.get(field, fallback)


def _heat_site_override_val(i: int, j: int, field: str, fallback=None):
    """Return per-site demo value if available, else global default."""
    val = _site_default(i, j, field)
    if val is None:
        return _heat_global_default(field, fallback)
    return val


def _field(widget_fn, label: str, key: str, omit_default: bool = False, **kwargs) -> None:
    """Render a form widget with an inline Omit checkbox for QA testing."""
    is_omitted: bool = st.session_state.get(f"{key}_omit", omit_default)
    f_col, x_col = st.columns([8, 1])
    with f_col:
        widget_fn(label, key=key, disabled=is_omitted, **kwargs)
    with x_col:
        st.checkbox(
            "Omit",
            value=omit_default,
            key=f"{key}_omit",
            help="Leave this field blank in the generated document (QA testing).",
        )


def _render_company_form(i: int, fp_months: list[tuple[int, int]]) -> None:
    col1, col2 = st.columns(2)
    with col1:
        _field(st.text_input, "Company Label", f"co_{i}_label",
               value=_co_default(i, "label"))
        _field(st.text_input, "Supplier Name", f"co_{i}_supplier",
               value=_co_default(i, "supplier"))
        _field(st.text_input, "Supplier Code", f"co_{i}_supplier_code",
               value=_co_default(i, "supplier_code"),
               help="Short alphanumeric code used in invoice numbers.")
        _field(st.text_area, "Supplier Address", f"co_{i}_supplier_address",
               value=_co_default(i, "supplier_address"), height=104)
    with col2:
        _field(st.text_input, "Customer Name", f"co_{i}_customer",
               value=_co_default(i, "customer"))
        _field(st.text_input, "Customer Code", f"co_{i}_customer_code",
               value=_co_default(i, "customer_code"))
        with st.expander("Advanced Options"):
            st.text_input("Currency", value=_co_default(i, "currency", _CURRENCY_DISPLAY.get(NEW_COMPANY_PLACEHOLDER["currency"], "EUR (€)")), key=f"co_{i}_currency")
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
            _render_site_form(i, j, fp_months)


def _render_site_form(i: int, j: int, fp_months: list[tuple[int, int]]) -> None:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Identity**")
        _field(st.text_input, "Site Label", f"site_{i}_{j}_label",
               value=_site_default(i, j, "label", ""))
        _field(st.text_area, "Customer Address", f"site_{i}_{j}_address",
               value=_site_default(i, j, "address", ""), height=104)
        _field(st.text_input, "City", f"site_{i}_{j}_city",
               value=_site_default(i, j, "city", ""))
        _field(st.text_input, "Postcode", f"site_{i}_{j}_postcode",
               value=_site_default(i, j, "postcode", ""))
        _field(st.text_input, "Heat Meter ID", f"site_{i}_{j}_meter_id",
               value=_site_default(i, j, "meter_id", ""))
    with col2:
        st.markdown("**Consumption**")
        override = st.checkbox(
            "Override global consumption defaults",
            key=f"site_{i}_{j}_override",
            value=False,
            help="When enabled, use site-specific capacity and consumption values instead of the global defaults.",
        )
        if override:
            _field(st.number_input, "Contracted Capacity (kW)", f"site_{i}_{j}_capacity_kw",
                   min_value=10, max_value=2000, step=5,
                   value=int(_heat_site_override_val(i, j, "capacity_kw", 150)))
            _field(st.number_input, "Capacity Rate (\u00a3/kW/month)", f"site_{i}_{j}_capacity_rate",
                   min_value=0.01, max_value=50.0, step=0.05, format="%.2f",
                   value=float(_heat_site_override_val(i, j, "capacity_rate", 5.50)))
            _field(st.number_input, "Base Monthly Consumption (kWh)", f"site_{i}_{j}_base_consumption",
                   min_value=100, max_value=500_000, step=100,
                   value=int(_heat_site_override_val(i, j, "base_consumption", 15000)))
            _field(st.number_input, "Base Unit Price (\u00a3/kWh)", f"site_{i}_{j}_unit_price_base",
                   min_value=0.010, max_value=0.500, step=0.001, format="%.3f",
                   value=float(_heat_site_override_val(i, j, "unit_price_base", 0.070)))
            _field(st.number_input, "Start Meter Reading (kWh)", f"site_{i}_{j}_start_reading",
                   min_value=0, max_value=9_999_999, step=1000,
                   value=int(_heat_site_override_val(i, j, "start_reading", 400000)))
            _field(st.number_input, "Supplier Emission Factor (kg CO₂e/kWh)",
                   f"site_{i}_{j}_supplier_ef",
                   min_value=0.0, max_value=2.0, step=0.001, format="%.4f",
                   value=float(_heat_site_override_val(i, j, "supplier_ef", 0.065)))
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
            month_options = [date(y, m, 1).strftime("%B %Y") for y, m in fp_months]
            st.multiselect(
                "Select billing months",
                options=month_options,
                default=month_options,
                key=f"site_{i}_{j}_months",
            )
    else:
        count = len(fp_months)
        st.caption(
            f"Will generate {count} monthly billing statement(s) "
            "covering the full financial period."
        )


# ---------------------------------------------------------------------------
# Data collection helpers
# ---------------------------------------------------------------------------

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


def _collect_form_data() -> dict:
    """Read all widget values from session state and return a structured dict."""
    s = st.session_state

    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end: date = s.get("fp_end", date(2026, 12, 31))
    fp_months = _months_in_range(fp_start, fp_end)
    month_label_map = {date(y, m, 1).strftime("%B %Y"): (y, m) for y, m in fp_months}

    # Global consumption defaults (read from rendered global-config widgets)
    g_cap_kw = int(s.get("heat_global_capacity_kw", _heat_global_default("capacity_kw", 150)))
    g_cap_rate = float(s.get("heat_global_capacity_rate", _heat_global_default("capacity_rate", 5.50)))
    g_base_cons = int(s.get("heat_global_base_consumption", _heat_global_default("base_consumption", 15000)))
    g_unit_price = float(s.get("heat_global_unit_price_base", _heat_global_default("unit_price_base", 0.070)))
    g_start_reading = int(s.get("heat_global_start_reading", _heat_global_default("start_reading", 400000)))
    g_supplier_ef = float(s.get("heat_global_supplier_ef", _heat_global_default("supplier_ef", 0.065)))
    g_omit = {
        f: bool(s.get(f"heat_global_{f}_omit", False))
        for f in _HEAT_CONSUMPTION_FIELDS
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
                    {"year": month_label_map[lbl][0], "month": month_label_map[lbl][1]}
                    for lbl in selected_labels
                    if lbl in month_label_map
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
                    f: bool(s.get(f"site_{i}_{j}_{f}_omit", False))
                    for f in _HEAT_CONSUMPTION_FIELDS
                }
            else:
                cap_kw, cap_rate, base_cons, unit_price, start_reading = (
                    g_cap_kw, g_cap_rate, g_base_cons, g_unit_price, g_start_reading
                )
                supplier_ef = g_supplier_ef
                cons_omit = dict(g_omit)

            site: dict = {
                "label": s.get(f"site_{i}_{j}_label", "") or f"Site {j + 1}",
                "customer_address": [
                    ln for ln in s.get(f"site_{i}_{j}_address", "").split("\n") if ln.strip()
                ],
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
                    **{f: bool(s.get(f"site_{i}_{j}_{f}_omit", False)) for f in _SITE_IDENTITY_FIELDS},
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
            "supplier_address": [
                ln for ln in s.get(f"co_{i}_supplier_address", "").split("\n") if ln.strip()
            ],
            "customer": s.get(f"co_{i}_customer", ""),
            "customer_code": s.get(f"co_{i}_customer_code", ""),
            "currency": s.get(f"co_{i}_currency", "GBP (£)"),
            "accent": s.get(f"co_{i}_accent", "#1E5B88"),
            "sites": sites,
            "_omit": {
                f: bool(s.get(f"co_{i}_{f}_omit", False))
                for f in _CO_OMIT_FIELDS
            },
        })

    return {
        "doc_title": s.get("doc_title", "District Heating Billing Statement"),
        "doc_subject": s.get("doc_subject", "Purchased Heat billing statements"),
        "doc_seed": int(s.get("doc_seed", 20260325)),
        "fp_label": s.get("fp_label", "Financial Year 2026"),
        "fp_start": fp_start.isoformat(),
        "fp_end": fp_end.isoformat(),
        "doc_language": _LANGUAGE_OPTIONS.get(s.get("doc_language_label", "English"), "en"),
        "doc_noise": float(s.get("doc_noise", 0.3)),
        "doc_inject_special_chars": bool(s.get("doc_inject_special_chars", False)),
        "xlsx_split_by_company": bool(s.get("xlsx_split_by_company", False)),
        "companies": companies,
    }


# ===========================================================================
# Scope 2 / Electricity form
# ===========================================================================

# ---------------------------------------------------------------------------
# Electricity: per-demo-site scope-specific supplements.
# Identity fields (label, address, city, postcode) come from companies.json.
# _override keys take precedence over global config; _tariff_names overrides
# the global tariff list; meter_id is the site meter identifier.
# ---------------------------------------------------------------------------

_ELEC_SITE_SUPPLEMENTS: list[list[dict]] = [
    [  # TFS UK (companies.json index 0)
        {"meter_id": "OVO-RHL-EL-90533"},
        {
            "meter_id": "OVO-MAN-EL-90534",
            "_override": {
                "start_reading": 28_190,
                "total_quantity": 54_200.0,
                "total_cost": 14_159.60,
            },
            "_tariff_names": [],  # no tariff breakdown for this site
        },
    ],
    [  # TFS Belgium
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
    [  # TFS Slovakia
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
    [  # TFS Ireland
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
    [  # TFS Denmark
        {
            "meter_id": "ORD-BAL-EL-62401",
            "_override": {
                "supplier_ef": None,
                "start_reading": 14_920,
                "total_quantity": 33_400.0,
                "total_cost": 8_519.40,
                "_omit_fields": ["supplier_ef"],
            },
            "_tariff_names": [],  # no tariff breakdown
        },
    ],
]

# Global defaults for consumption fields — inherited by all sites.
_ELECTRICITY_GLOBAL_DEFAULTS = {
    "supplier_ef": 0.2070,
    "unit": "kWh",
    "start_reading": 45_280,
    "total_quantity": 87_600.0,
    "total_cost": 22_922.00,
    "tariffs": ["Day Rate", "Night Rate (Economy 7)"],
}

_ELEC_CO_OMIT_FIELDS = ["label", "supplier", "supplier_code", "supplier_address", "customer", "customer_code"]
# Site-level omit only applies to identity fields; grid/consumption come from global or override
_ELEC_SITE_OMIT_FIELDS = ["label", "address", "city", "postcode", "meter_id"]


def _elec_global_default(field: str, fallback=None):
    return _ELECTRICITY_GLOBAL_DEFAULTS.get(field, fallback)


def _elec_global_tariff_default(k: int, fallback: str = "") -> str:
    tariffs = _ELECTRICITY_GLOBAL_DEFAULTS.get("tariffs", [])
    return tariffs[k] if k < len(tariffs) else fallback


def _elec_co_default(i: int, field: str, fallback: str = "") -> str:
    """Return pre-filled value from companies.json / suppliers.json ('electricity' scope)."""
    co = _config_company(i)
    if field in ("label", "customer"):
        return co.get("name", fallback)
    if field == "customer_code":
        return co.get("code", _make_code(co.get("name", "")) or fallback)
    if field == "currency":
        return _CURRENCY_DISPLAY.get(co.get("currency", "EUR"), "EUR (€)")
    sup = _config_supplier("electricity", i)
    if field == "supplier":
        return sup.get("name", fallback)
    if field == "supplier_code":
        return sup.get("code", fallback)
    if field == "supplier_address":
        return sup.get("address", fallback)
    return fallback


def _elec_site_default(i: int, j: int, field: str, fallback=None):
    """Return identity from companies.json; scope-specific data from _ELEC_SITE_SUPPLEMENTS."""
    site = _config_site(i, j)
    if field in ("label", "city", "postcode", "address"):
        return site.get(field, fallback)
    if field == "meter_id":
        if i < len(_ELEC_SITE_SUPPLEMENTS) and j < len(_ELEC_SITE_SUPPLEMENTS[i]):
            mid = _ELEC_SITE_SUPPLEMENTS[i][j].get("meter_id")
            if mid:
                return mid
        sup = _config_supplier("electricity", i)
        return _make_meter_id("electricity", sup.get("code", "SUP"), site.get("city", ""), i, j)
    if i < len(_ELEC_SITE_SUPPLEMENTS) and j < len(_ELEC_SITE_SUPPLEMENTS[i]):
        return _ELEC_SITE_SUPPLEMENTS[i][j].get(field, fallback)
    return fallback


def _elec_site_has_override(i: int, j: int) -> bool:
    """True when the sample site has pre-configured overrides."""
    return bool(_elec_site_default(i, j, "_override"))


def _elec_site_override_val(i: int, j: int, field: str, fallback=None):
    """Return a site-level override default from the sample data, or fallback."""
    ov = _elec_site_default(i, j, "_override") or {}
    return ov.get(field, _elec_global_default(field, fallback))


def _elec_site_override_omit(i: int, j: int, field: str) -> bool:
    ov = _elec_site_default(i, j, "_override") or {}
    return field in ov.get("_omit_fields", [])


# ---------------------------------------------------------------------------
# Electricity: entry point
# ---------------------------------------------------------------------------

def _render_electricity_form() -> dict:
    st.subheader("Electricity")
    st.caption("Scope 2 purchased electricity — consumption statements with per-billing-period breakdown.")

    _render_document_settings()
    _render_financial_period()

    s = st.session_state
    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end:   date = s.get("fp_end",   date(2026, 12, 31))
    fp_months = _months_in_range(fp_start, fp_end)

    _render_electricity_global_config()
    _render_electricity_companies_section(fp_months)

    return _collect_electricity_form_data()


def _render_electricity_global_config() -> None:
    """Global grid / consumption / tariff defaults inherited by all sites."""
    _init_elec_global_random()
    st.markdown("#### Global Electricity Configuration")
    st.caption(
        "These defaults apply to every site. Sites can optionally override individual fields."
    )

    gd = _ELECTRICITY_GLOBAL_DEFAULTS
    col1, col2 = st.columns(2)
    with col1:
        _field(
            st.number_input, "Supplier Emission Factor (kg CO\u2082e/kWh)", "elec_global_supplier_ef",
            value=float(gd["supplier_ef"]),
            min_value=0.0, max_value=2.0, step=0.001, format="%.4f",
            help="Supplier-reported emission factor for purchased electricity.",
        )
        st.selectbox(
            "Measurement Unit",
            options=["kWh", "MWh"],
            index=0 if gd["unit"] == "kWh" else 1,
            key="elec_global_unit",
        )
    with col2:
        _field(
            st.number_input, "Default Start Meter Reading", "elec_global_start_reading",
            value=int(gd["start_reading"]),
            min_value=0, max_value=999_999_999, step=100,
            help="Opening meter reading used unless overridden per site.",
        )
        _field(
            st.number_input, "Default Annual Quantity", "elec_global_total_quantity",
            value=float(gd["total_quantity"]),
            min_value=0.0, step=100.0, format="%.2f",
            help="Total annual electricity consumption used unless overridden per site.",
        )
        _field(
            st.number_input, "Default Annual Cost", "elec_global_total_cost",
            value=float(gd["total_cost"]),
            min_value=0.0, step=10.0, format="%.2f",
            help="Total annual electricity cost used unless overridden per site.",
        )

    st.markdown("**Tariff Rates**")
    st.caption(
        "Define tariff names only. Values are randomly split from the annual totals at generation time."
    )
    n_tariffs_global = st.number_input(
        "Number of tariff rates",
        min_value=0, max_value=10,
        value=len(gd["tariffs"]),
        step=1,
        key="elec_global_n_tariffs",
        help="Define shared tariff names here. Sites can select which ones to include.",
    )
    if int(n_tariffs_global) == 0:
        st.caption("No tariffs defined — all sites will show totals only.")
    else:
        for k in range(int(n_tariffs_global)):
            st.text_input(
                f"Tariff {k + 1} name",
                key=f"elec_global_t{k}_name",
                value=_elec_global_tariff_default(k, f"Tariff {k + 1}"),
            )

    st.divider()


def _render_electricity_companies_section(fp_months: list[tuple[int, int]]) -> None:
    st.markdown("#### Companies")
    n_companies = st.number_input(
        "Number of companies",
        min_value=1, max_value=10, value=1, step=1,
        key="elec_n_companies",
    )
    for i in range(int(n_companies)):
        co_label = st.session_state.get(f"elec_co_{i}_label") or _elec_co_default(i, "label") or f"Company {i + 1}"
        with st.expander(f"Company {i + 1}: {co_label}", expanded=(i == 0)):
            _render_electricity_company_form(i, fp_months)


def _render_electricity_company_form(i: int, fp_months: list[tuple[int, int]]) -> None:
    col1, col2 = st.columns(2)
    with col1:
        _field(st.text_input, "Company Label", f"elec_co_{i}_label",
               value=_elec_co_default(i, "label"))
        _field(st.text_input, "Supplier Name", f"elec_co_{i}_supplier",
               value=_elec_co_default(i, "supplier"))
        _field(st.text_input, "Supplier Code", f"elec_co_{i}_supplier_code",
               value=_elec_co_default(i, "supplier_code"),
               help="Short alphanumeric code used in reference numbers.")
        _field(st.text_area, "Supplier Address", f"elec_co_{i}_supplier_address",
               value=_elec_co_default(i, "supplier_address"), height=104)
    with col2:
        _field(st.text_input, "Customer Name", f"elec_co_{i}_customer",
               value=_elec_co_default(i, "customer"))
        _field(st.text_input, "Customer Code", f"elec_co_{i}_customer_code",
               value=_elec_co_default(i, "customer_code"))
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
        min_value=1, max_value=20, value=1, step=1,
        key=f"elec_n_sites_{i}",
    )
    for j in range(int(n_sites)):
        site_label = st.session_state.get(f"elec_site_{i}_{j}_label") or _elec_site_default(i, j, "label") or f"Site {j + 1}"
        with st.expander(f"Site {j + 1}: {site_label}", expanded=(j == 0 and i == 0)):
            _render_electricity_site_form(i, j, fp_months)


def _render_electricity_site_form(i: int, j: int, fp_months: list[tuple[int, int]]) -> None:
    col1, col2 = st.columns(2)
    with col1:
        _field(st.text_input, "Site Label", f"elec_site_{i}_{j}_label",
               value=_elec_site_default(i, j, "label", ""))
        _field(st.text_area, "Customer Address", f"elec_site_{i}_{j}_address",
               value=_elec_site_default(i, j, "address", ""), height=104)
        _field(st.text_input, "City", f"elec_site_{i}_{j}_city",
               value=_elec_site_default(i, j, "city", ""))
        _field(st.text_input, "Postcode", f"elec_site_{i}_{j}_postcode",
               value=_elec_site_default(i, j, "postcode", ""))
        _field(st.text_input, "Electricity Meter ID", f"elec_site_{i}_{j}_meter_id",
               value=_elec_site_default(i, j, "meter_id", ""))

    with col2:
        # ── Grid / consumption override ───────────────────────────────────────
        override_key = f"elec_site_{i}_{j}_override"
        override = st.checkbox(
            "Override global grid & consumption settings",
            value=st.session_state.get(override_key, False),
            key=override_key,
        )
        if override:
            _field(
                st.number_input, "Supplier Emission Factor (kg CO\u2082e/kWh)",
                f"elec_site_{i}_{j}_supplier_ef",
                omit_default=_elec_site_override_omit(i, j, "supplier_ef"),
                value=float(_elec_site_override_val(i, j, "supplier_ef") or 0.0),
                min_value=0.0, max_value=2.0, step=0.001, format="%.4f",
            )
            st.selectbox(
                "Measurement Unit", options=["kWh", "MWh"],
                index=0 if _elec_site_override_val(i, j, "unit", "kWh") == "kWh" else 1,
                key=f"elec_site_{i}_{j}_unit",
            )
            _field(
                st.number_input, "Start Meter Reading", f"elec_site_{i}_{j}_start_reading",
                value=int(_elec_site_override_val(i, j, "start_reading", 0)),
                min_value=0, max_value=999_999_999, step=100,
            )
            _field(
                st.number_input, "Annual Quantity", f"elec_site_{i}_{j}_total_quantity",
                value=float(_elec_site_override_val(i, j, "total_quantity", 0.0)),
                min_value=0.0, step=100.0, format="%.2f",
            )
            _field(
                st.number_input, "Annual Cost", f"elec_site_{i}_{j}_total_cost",
                value=float(_elec_site_override_val(i, j, "total_cost", 0.0)),
                min_value=0.0, step=10.0, format="%.2f",
            )
        else:
            g_qty  = st.session_state.get("elec_global_total_quantity",
                                          _ELECTRICITY_GLOBAL_DEFAULTS["total_quantity"])
            g_cost = st.session_state.get("elec_global_total_cost",
                                          _ELECTRICITY_GLOBAL_DEFAULTS["total_cost"])
            st.caption(
                f"Using global defaults — {float(g_qty):,.0f} kWh, £{float(g_cost):,.2f}."
            )

    # ── Tariff selection ──────────────────────────────────────────────────────
    s = st.session_state
    global_n_tariffs = int(s.get("elec_global_n_tariffs",
                                  len(_ELECTRICITY_GLOBAL_DEFAULTS["tariffs"])))
    global_tariff_names = [
        s.get(f"elec_global_t{k}_name", _elec_global_tariff_default(k, f"Tariff {k + 1}"))
        for k in range(global_n_tariffs)
        if s.get(f"elec_global_t{k}_name", _elec_global_tariff_default(k, "")).strip()
    ]

    if global_tariff_names:
        st.markdown("**Tariffs to include**")
        # Pre-selected default from sample data
        sample_tariff_names_key = "_tariff_names"
        sample = _elec_site_default(i, j, sample_tariff_names_key)
        if sample is None:
            default_selection = global_tariff_names
        else:
            # sample is a list; empty list = no tariffs, list = specific selection
            default_selection = [n for n in global_tariff_names if n in sample] if sample else []

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
                default=[n for n in default_selection if n in global_tariff_names],
                key=f"elec_site_{i}_{j}_tariffs",
            )
        else:
            st.caption(f"All {len(global_tariff_names)} global tariff(s) will be included.")
    else:
        st.caption("No global tariffs defined — output will show totals only.")

    # ── Billing Periods ───────────────────────────────────────────────────────
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
            month_options = [date(y, m, 1).strftime("%B %Y") for y, m in fp_months]
            st.multiselect(
                "Select billing months",
                options=month_options,
                default=month_options,
                key=f"elec_site_{i}_{j}_months",
            )
    else:
        count = len(fp_months)
        st.caption(
            f"Will generate {count} monthly billing statement(s) "
            "covering the full financial period."
        )


# ---------------------------------------------------------------------------
# Electricity: data collection
# ---------------------------------------------------------------------------

def _collect_electricity_form_data() -> dict:
    s = st.session_state

    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end: date   = s.get("fp_end",   date(2026, 12, 31))
    fp_months      = _months_in_range(fp_start, fp_end)
    month_label_map = {date(y, m, 1).strftime("%B %Y"): (y, m) for y, m in fp_months}

    # ── Global grid / consumption defaults ───────────────────────────────────
    g_supplier_ef        = s.get("elec_global_supplier_ef",
                                  _ELECTRICITY_GLOBAL_DEFAULTS["supplier_ef"])
    g_supplier_ef_omit   = bool(s.get("elec_global_supplier_ef_omit", False))
    g_unit               = s.get("elec_global_unit",
                                  _ELECTRICITY_GLOBAL_DEFAULTS["unit"])
    g_start_reading      = s.get("elec_global_start_reading",
                                  _ELECTRICITY_GLOBAL_DEFAULTS["start_reading"])
    g_start_reading_omit = bool(s.get("elec_global_start_reading_omit", False))
    g_total_quantity     = s.get("elec_global_total_quantity",
                                  _ELECTRICITY_GLOBAL_DEFAULTS["total_quantity"])
    g_total_qty_omit     = bool(s.get("elec_global_total_quantity_omit", False))
    g_total_cost         = s.get("elec_global_total_cost",
                                  _ELECTRICITY_GLOBAL_DEFAULTS["total_cost"])
    g_total_cost_omit    = bool(s.get("elec_global_total_cost_omit", False))

    # ── Global tariffs (name-only; values split randomly at generation time) ──
    global_n_tariffs = int(s.get("elec_global_n_tariffs",
                                  len(_ELECTRICITY_GLOBAL_DEFAULTS["tariffs"])))
    global_tariffs: list[dict] = []
    for k in range(global_n_tariffs):
        name = s.get(f"elec_global_t{k}_name",
                     _elec_global_tariff_default(k, "")).strip()
        if name:
            global_tariffs.append({"name": name})
    global_tariff_names = [t["name"] for t in global_tariffs]

    n_companies = int(s.get("elec_n_companies", 1))
    companies: list[dict] = []

    for i in range(n_companies):
        n_sites = int(s.get(f"elec_n_sites_{i}", 1))
        sites: list[dict] = []

        for j in range(n_sites):
            # ── Grid / consumption: global or site override ───────────────────
            has_override = bool(s.get(f"elec_site_{i}_{j}_override",
                                      _elec_site_has_override(i, j)))
            if has_override:
                supplier_ef_raw  = s.get(f"elec_site_{i}_{j}_supplier_ef",
                                         _elec_site_override_val(i, j, "supplier_ef", g_supplier_ef))
                supplier_ef_omit = bool(s.get(f"elec_site_{i}_{j}_supplier_ef_omit",
                                              _elec_site_override_omit(i, j, "supplier_ef")))
                unit             = s.get(f"elec_site_{i}_{j}_unit",
                                         _elec_site_override_val(i, j, "unit", g_unit))
                start_reading    = s.get(f"elec_site_{i}_{j}_start_reading",
                                         _elec_site_override_val(i, j, "start_reading", g_start_reading))
                start_reading_omit = bool(s.get(f"elec_site_{i}_{j}_start_reading_omit", False))
                total_quantity   = s.get(f"elec_site_{i}_{j}_total_quantity",
                                         _elec_site_override_val(i, j, "total_quantity", g_total_quantity))
                total_qty_omit   = bool(s.get(f"elec_site_{i}_{j}_total_quantity_omit", False))
                total_cost       = s.get(f"elec_site_{i}_{j}_total_cost",
                                         _elec_site_override_val(i, j, "total_cost", g_total_cost))
                total_cost_omit  = bool(s.get(f"elec_site_{i}_{j}_total_cost_omit", False))
            else:
                supplier_ef_raw    = g_supplier_ef
                supplier_ef_omit   = g_supplier_ef_omit
                unit               = g_unit
                start_reading      = g_start_reading
                start_reading_omit = g_start_reading_omit
                total_quantity     = g_total_quantity
                total_qty_omit     = g_total_qty_omit
                total_cost         = g_total_cost
                total_cost_omit    = g_total_cost_omit

            # ── Tariff selection ──────────────────────────────────────────────
            tariff_mode = s.get(f"elec_site_{i}_{j}_tariff_mode", "All global tariffs")
            if tariff_mode == "Custom selection" and global_tariffs:
                selected = set(s.get(f"elec_site_{i}_{j}_tariffs", global_tariff_names))
                site_tariffs = [t for t in global_tariffs if t["name"] in selected]
            elif not global_tariffs:
                site_tariffs = []
            else:
                # check sample default (e.g. _tariff_names: [] means no tariffs for this site)
                sample_names = _elec_site_default(i, j, "_tariff_names")
                if sample_names is not None:
                    site_tariffs = [t for t in global_tariffs if t["name"] in sample_names]
                else:
                    site_tariffs = global_tariffs

            # ── Billing periods ───────────────────────────────────────────────
            period_mode = s.get(f"elec_site_{i}_{j}_period_mode",
                                "All months in financial period")
            billing_periods: list[dict] | None = None
            if period_mode == "Custom months":
                selected_labels: list[str] = s.get(f"elec_site_{i}_{j}_months", [])
                billing_periods = [
                    {"year": month_label_map[lbl][0], "month": month_label_map[lbl][1]}
                    for lbl in selected_labels
                    if lbl in month_label_map
                ]

            sites.append({
                "label": s.get(f"elec_site_{i}_{j}_label", "") or f"Site {j + 1}",
                "customer_address": [
                    ln for ln in s.get(f"elec_site_{i}_{j}_address", "").split("\n") if ln.strip()
                ],
                "city":         s.get(f"elec_site_{i}_{j}_city", ""),
                "postcode":     s.get(f"elec_site_{i}_{j}_postcode", ""),
                "meter_id":     s.get(f"elec_site_{i}_{j}_meter_id", ""),
                "supplier_ef":  str(supplier_ef_raw) if supplier_ef_raw is not None else "0",
                "unit":         unit,
                "start_reading": int(start_reading),
                "total_quantity": str(total_quantity),
                "total_cost":    str(total_cost),
                "tariffs":       site_tariffs,
                "_omit": {
                    **{f: bool(s.get(f"elec_site_{i}_{j}_{f}_omit", False))
                       for f in _ELEC_SITE_OMIT_FIELDS},
                    "supplier_ef":    supplier_ef_omit,
                    "start_reading":  start_reading_omit,
                    "total_quantity": total_qty_omit,
                    "total_cost":     total_cost_omit,
                },
                **({"billing_periods": billing_periods} if billing_periods is not None else {}),
            })

        companies.append({
            "label": s.get(f"elec_co_{i}_label", "") or f"Company {i + 1}",
            "supplier": s.get(f"elec_co_{i}_supplier", ""),
            "supplier_code": s.get(f"elec_co_{i}_supplier_code", ""),
            "supplier_address": [
                ln for ln in s.get(f"elec_co_{i}_supplier_address", "").split("\n") if ln.strip()
            ],
            "customer": s.get(f"elec_co_{i}_customer", ""),
            "customer_code": s.get(f"elec_co_{i}_customer_code", ""),
            "currency": s.get(f"elec_co_{i}_currency", "GBP (£)"),
            "accent": s.get(f"elec_co_{i}_accent", "#1E5B88"),
            "sites": sites,
            "_omit": {
                f: bool(s.get(f"elec_co_{i}_{f}_omit", False))
                for f in _ELEC_CO_OMIT_FIELDS
            },
        })

    return {
        "_category": "electricity",
        "doc_title": s.get("doc_title", "Electricity Consumption Statement"),
        "doc_subject": s.get("doc_subject", "Scope 2 purchased electricity"),
        "doc_seed": int(s.get("doc_seed", 20260325)),
        "fp_label": s.get("fp_label", "Financial Year 2026"),
        "fp_start": fp_start.isoformat(),
        "fp_end": fp_end.isoformat(),
        "doc_language": _LANGUAGE_OPTIONS.get(s.get("doc_language_label", "English"), "en"),
        "doc_noise": float(s.get("doc_noise", 0.3)),
        "doc_inject_special_chars": bool(s.get("doc_inject_special_chars", False)),
        "xlsx_split_by_company": bool(s.get("xlsx_split_by_company", False)),
        "companies": companies,
    }
