from __future__ import annotations

from datetime import date

import streamlit as st

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
            st.text_input(
                "Output Filename",
                value="billing_statement.pdf",
                key="doc_filename",
                help="Filename for the downloaded PDF.",
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


def _render_financial_period() -> list[tuple[int, int]]:
    """Render financial period widgets and return list of (year, month) tuples."""
    st.markdown("#### Financial Period")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.text_input("Period Label", value="Financial Year 2026", key="fp_label")
    with col2:
        st.date_input("Start Date", value=date(2026, 1, 1), key="fp_start")
    with col3:
        st.date_input("End Date", value=date(2026, 12, 31), key="fp_end")

    fp_start: date = st.session_state.get("fp_start", date(2026, 1, 1))
    fp_end: date = st.session_state.get("fp_end", date(2026, 12, 31))

    if fp_end < fp_start:
        st.error("End date must be after start date.")
        return []

    fp_months = _months_in_range(fp_start, fp_end)
    st.caption(f"Period spans {len(fp_months)} month(s).")
    return fp_months


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
        co_label = st.session_state.get(f"co_{i}_label") or f"Company {i + 1}"
        with st.expander(f"Company {i + 1}: {co_label}", expanded=(i == 0)):
            _render_company_form(i, fp_months)


_COMPANY_DEFAULTS = [
    {
        "label": "Toyota Financial Services UK",
        "supplier": "Albion District Energy Services",
        "supplier_code": "ADES",
        "supplier_address": "Meridian Utilities House\n41 Station Approach\nReading RG1 1LX\nUnited Kingdom",
        "customer": "Toyota Financial Services UK",
        "customer_code": "TFSUK",
    },
    {
        "label": "Toyota Financial Services Belgium",
        "supplier": "Benelux Thermal Grid NV",
        "supplier_code": "BTG",
        "supplier_address": "Canal Logistics Campus\n27 Havenlaan\n1000 Brussels\nBelgium",
        "customer": "Toyota Financial Services Belgium",
        "customer_code": "TFSBE",
    },
]

_SITE_DEFAULTS = [
    [
        {
            "label": "Redhill HQ",
            "address": "Great Burgh, Burgh Heath\nEpsom Road\nRedhill RH1 5UZ\nUnited Kingdom",
            "city": "Redhill",
            "postcode": "RH1 5UZ",
            "meter_id": "ADES-RHL-HT-90533",
            "capacity_kw": 265,
            "capacity_rate": 6.15,
            "base_consumption": 28100,
            "unit_price_base": 0.079,
            "start_reading": 915860,
        },
        {
            "label": "Manchester Office",
            "address": "5 New Bailey Square\nStanley Street\nSalford M3 5JL\nUnited Kingdom",
            "city": "Manchester",
            "postcode": "M3 5JL",
            "meter_id": "ADES-MAN-HT-90534",
            "capacity_kw": 180,
            "capacity_rate": 5.95,
            "base_consumption": 19400,
            "unit_price_base": 0.074,
            "start_reading": 624180,
        },
    ],
    [
        {
            "label": "Brussels Office",
            "address": "Avenue du Bourget 42\n1130 Brussels\nBelgium",
            "city": "Brussels",
            "postcode": "1130",
            "meter_id": "BTG-BRU-HT-31021",
            "capacity_kw": 132,
            "capacity_rate": 5.35,
            "base_consumption": 15600,
            "unit_price_base": 0.067,
            "start_reading": 351220,
        },
        {
            "label": "Antwerp Office",
            "address": "Noorderlaan 147\n2030 Antwerp\nBelgium",
            "city": "Antwerp",
            "postcode": "2030",
            "meter_id": "BTG-ANR-HT-31022",
            "capacity_kw": 118,
            "capacity_rate": 5.10,
            "base_consumption": 14100,
            "unit_price_base": 0.065,
            "start_reading": 287640,
        },
    ],
]


_CO_OMIT_FIELDS = ["label", "supplier", "supplier_code", "supplier_address", "customer", "customer_code"]
_SITE_OMIT_FIELDS = [
    "label", "address", "city", "postcode", "meter_id",
    "capacity_kw", "capacity_rate", "base_consumption", "unit_price_base", "start_reading",
]


def _co_default(i: int, field: str, fallback: str = "") -> str:
    """Return pre-filled demo value only for the first company; new companies start blank."""
    if i == 0 and _COMPANY_DEFAULTS:
        return _COMPANY_DEFAULTS[0].get(field, fallback)
    return fallback


def _site_default(i: int, j: int, field: str, fallback=None):
    """Return pre-filled demo value only for the first company's sites."""
    if i == 0 and _SITE_DEFAULTS and j < len(_SITE_DEFAULTS[0]):
        return _SITE_DEFAULTS[0][j].get(field, fallback)
    return fallback


def _field(widget_fn, label: str, key: str, **kwargs) -> None:
    """Render a form widget with an inline Omit checkbox for QA testing."""
    is_omitted: bool = st.session_state.get(f"{key}_omit", False)
    f_col, x_col = st.columns([8, 1])
    with f_col:
        widget_fn(label, key=key, disabled=is_omitted, **kwargs)
    with x_col:
        st.checkbox(
            "Omit",
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
            st.text_input("Currency", value="GBP (£)", key=f"co_{i}_currency")
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
        site_label = st.session_state.get(f"site_{i}_{j}_label") or f"Site {j + 1}"
        with st.expander(f"Site {j + 1}: {site_label}", expanded=(j == 0 and i == 0)):
            _render_site_form(i, j, fp_months)


def _render_site_form(i: int, j: int, fp_months: list[tuple[int, int]]) -> None:
    col1, col2 = st.columns(2)
    with col1:
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
        _field(st.number_input, "Contracted Capacity (kW)", f"site_{i}_{j}_capacity_kw",
               min_value=50, max_value=500, step=5,
               value=_site_default(i, j, "capacity_kw", 265))
        _field(st.number_input, "Capacity Rate (£/kW/month)", f"site_{i}_{j}_capacity_rate",
               min_value=0.01, max_value=50.0, step=0.05, format="%.2f",
               value=_site_default(i, j, "capacity_rate", 6.15))
        _field(st.number_input, "Base Monthly Consumption (kWh)", f"site_{i}_{j}_base_consumption",
               min_value=5000, max_value=50000, step=100,
               value=_site_default(i, j, "base_consumption", 28100))
        _field(st.number_input, "Base Unit Price (£/kWh)", f"site_{i}_{j}_unit_price_base",
               min_value=0.040, max_value=0.120, step=0.001, format="%.3f",
               value=_site_default(i, j, "unit_price_base", 0.079),
               help="Starting unit price before seasonal adjustments.")
        _field(st.number_input, "Start Meter Reading (kWh)", f"site_{i}_{j}_start_reading",
               min_value=0, max_value=9_999_999, step=1000,
               value=_site_default(i, j, "start_reading", 915860))

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

            site: dict = {
                "label": s.get(f"site_{i}_{j}_label", "") or f"Site {j + 1}",
                "customer_address": [
                    ln for ln in s.get(f"site_{i}_{j}_address", "").split("\n") if ln.strip()
                ],
                "city": s.get(f"site_{i}_{j}_city", ""),
                "postcode": s.get(f"site_{i}_{j}_postcode", ""),
                "meter_id": s.get(f"site_{i}_{j}_meter_id", ""),
                "capacity_kw": int(s.get(f"site_{i}_{j}_capacity_kw", 200)),
                "capacity_rate": str(s.get(f"site_{i}_{j}_capacity_rate", 6.00)),
                "base_consumption": int(s.get(f"site_{i}_{j}_base_consumption", 20000)),
                "unit_price_base": str(s.get(f"site_{i}_{j}_unit_price_base", 0.075)),
                "start_reading": int(s.get(f"site_{i}_{j}_start_reading", 500000)),
                "_omit": {
                    f: bool(s.get(f"site_{i}_{j}_{f}_omit", False))
                    for f in _SITE_OMIT_FIELDS
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
        "doc_filename": s.get("doc_filename", "billing_statement.pdf"),
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
