from __future__ import annotations

import json
import os
import random
from calendar import monthrange
from datetime import date

import streamlit as st

from components.distractor_fields_controls import (
    collect_distractor_field_form_data as _collect_distractor_fields,
    render_distractor_field_controls as _render_distractor_field_controls,
)
from components.invalid_data_controls import (
    collect_invalid_data_form_data as _collect_invalid_data,
    render_invalid_data_controls as _render_invalid_data_controls,
)
from components.layout_controls import (
    collect_layout_form_data as _collect_layout,
    render_layout_controls as _render_layout_controls,
)
from components.sidebar import get_document_type_config
from generators.waste_reference import (
    COUNCILS,
    DEFAULT_SITE_STREAMS,
    SIC_CODES,
    STREAM_PRESETS,
    WASTE_CONTRACTORS,
)
from utils.currency import currency_index as _currency_index, currency_options as _currency_options
from utils.document_catalog import document_type_requires_company_currency

_CATEGORY = "waste_generated_in_operations"
_CONFIG_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config")
_MULTI_DOCUMENT_TYPES = {
    "waste_transfer_note",
    "waste_contractor_invoice",
    "weighbridge_ticket",
    "hazardous_waste_consignment_note",
}
_CURRENCY_DISPLAY = {"GBP": "GBP (£)", "EUR": "EUR (€)", "USD": "USD ($)"}
_STREAM_LABELS = {key: preset["label"] for key, preset in STREAM_PRESETS.items()}


def _load_json(filename: str) -> list[dict]:
    try:
        with open(os.path.join(_CONFIG_DIR, filename), encoding="utf-8") as fh:
            return json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


_COMPANIES_CONFIG = _load_json("companies.json")
_SITES_CONFIG = _load_json("sites.json")


def _company_default(i: int) -> dict:
    return _COMPANIES_CONFIG[i % len(_COMPANIES_CONFIG)] if _COMPANIES_CONFIG else {}


def _site_default(i: int, j: int) -> dict:
    return _SITES_CONFIG[(i * 3 + j) % len(_SITES_CONFIG)] if _SITES_CONFIG else {}


def _document_defaults(document_type: str | None) -> tuple[str, str]:
    cfg = get_document_type_config(_CATEGORY, document_type or "")
    return cfg.get("default_title", "Document"), cfg.get("default_subject", "")


def _sync_document_setting_defaults(document_type: str | None) -> None:
    selection_key = f"{_CATEGORY}:{document_type or ''}"
    if st.session_state.get("_document_settings_selection") == selection_key:
        return
    title, subject = _document_defaults(document_type)
    st.session_state["doc_title"] = title
    st.session_state["doc_subject"] = subject
    st.session_state["_document_settings_selection"] = selection_key


def _render_document_settings(document_type: str | None) -> None:
    _sync_document_setting_defaults(document_type)
    title, subject = _document_defaults(document_type)
    with st.expander("Document Settings", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.text_input("Document Title", value=title, key="doc_title")
        with col2:
            st.text_input("Document Subject", value=subject, key="doc_subject")
            st.number_input("Random Seed", value=20261116, min_value=0, max_value=2**31 - 1, step=1, key="doc_seed")
        st.checkbox("Inject special characters", key="doc_inject_special_chars",
                    help="Append QA characters to generated text values.")
        _render_distractor_field_controls()
        _render_layout_controls()
        st.divider()
        _render_invalid_data_controls()


def _render_financial_period() -> None:
    if "fp_start" not in st.session_state:
        year, month = 2026, random.randint(1, 10)
        st.session_state["fp_start"] = date(year, month, 1)
        st.session_state["fp_end"] = date(year, month, monthrange(year, month)[1])
        st.session_state["fp_label"] = st.session_state["fp_start"].strftime("%B %Y")
    st.markdown("#### Reporting Period")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.text_input("Period Label", key="fp_label")
    with col2:
        st.date_input("Start Date", key="fp_start")
    with col3:
        st.date_input("End Date", key="fp_end")
    if st.session_state["fp_end"] < st.session_state["fp_start"]:
        st.error("End date must be after start date.")


# ── special cases per document type ──────────────────────────────────────────


def _render_special_cases(document_type: str | None) -> None:
    s = st.session_state
    st.markdown("#### Generation Defaults & Special Cases")
    col1, col2 = st.columns(2)
    with col1:
        if document_type in _MULTI_DOCUMENT_TYPES:
            st.number_input("Number of documents to generate", min_value=1, max_value=200, step=1, key="waste_doc_count",
                            help="Generate 1 for a single file, or more to download them as a ZIP archive.")
    with col2:
        st.number_input("Container fill factor", min_value=0.1, max_value=1.0, step=0.05, key="waste_fill_factor",
                        help="Used for volume → tonnes conversion (EA density factors assume a full container).")

    if document_type == "waste_transfer_note":
        st.radio("Transfer type", ["single", "season_ticket"], horizontal=True, key="waste_wtn_mode",
                 format_func=lambda v: {"single": "Single transfer", "season_ticket": "Season ticket (12-month estimate)"}[v])
        st.radio("Quantity written as", ["approx", "container"], horizontal=True, key="waste_wtn_quantity_style",
                 format_func=lambda v: {"approx": "Approx. tonnes", "container": "Container count (1 x 8yd skip / n x 1100L)"}[v])
        st.checkbox("R/D code in destination text", key="waste_wtn_rd_in_destination")
    elif document_type == "waste_contractor_invoice":
        col1, col2 = st.columns(2)
        with col1:
            st.radio("Document role", ["invoice", "credit_note"], horizontal=True, key="waste_inv_role",
                     format_func=lambda v: {"invoice": "Invoice", "credit_note": "Credit note"}[v])
            st.selectbox("Quantity basis", ["lift_priced", "weighed", "unweighed", "mixed"], key="waste_inv_basis",
                         format_func=lambda v: {"lift_priced": "Bins per lift, skips weighed (typical)",
                                                "weighed": "Weighed – on-board bin weighing (kg)",
                                                "unweighed": "Unweighed (skip volume / bin lifts)", "mixed": "Mixed"}[v])
            st.selectbox("Billing period", ["calendar_month", "four_five_week"], key="waste_inv_billing_period",
                         format_func=lambda v: {"calendar_month": "Calendar month", "four_five_week": "4/5-week billing period"}[v])
            st.number_input("VAT Rate (%)", min_value=0, max_value=100, step=1, key="waste_vat_rate")
        with col2:
            st.checkbox("Show treatment route on lines", key="waste_inv_show_routes")
            st.checkbox("Cardboard rebate line", key="waste_inv_rebate")
            st.checkbox("Landfill tax line", key="waste_inv_landfill_tax")
            st.checkbox("Non-waste charges (rental, permit, surcharge…)", key="waste_inv_non_waste_charges")
            st.checkbox("Excess weight line (kg over allowance, £/t)", key="waste_inv_excess_weight")
            st.checkbox("Weight column without unit", key="waste_inv_unitless_weight")
            st.checkbox("European number / date format", key="waste_inv_eu_format")
            st.checkbox("Weighbridge references (possible duplicates)", key="waste_inv_weighbridge_refs")
            st.checkbox("Service period spans two months", key="waste_inv_two_month_period")
            st.checkbox("No waste lines (service charges only)", key="waste_inv_no_waste_lines")
    elif document_type == "weighbridge_ticket":
        st.radio("Weight unit", ["kg", "t"], horizontal=True, key="waste_wb_unit")
        st.radio("Weights printed", ["gross_tare_net", "stored_tare", "gross_only", "net_mismatch"], horizontal=True, key="waste_wb_mode",
                 format_func=lambda v: {"gross_tare_net": "Gross / tare / net", "stored_tare": "Stored tare",
                                        "gross_only": "Gross weight only", "net_mismatch": "Net ≠ gross − tare"}[v])
    elif document_type == "waste_report":
        st.checkbox("Route % split (Recycled / Energy recovery / AD / Landfill)", key="waste_rep_route_split")
        st.checkbox("Multi-site report (all sites of the first organisation)", key="waste_rep_multi_site")
        st.checkbox("TOTAL row", key="waste_rep_totals_row")
    elif document_type == "hazardous_waste_consignment_note":
        st.checkbox("Part E received quantity differs from Part B estimate", key="waste_hwcn_part_e_differs")
        st.checkbox("Interim operation (R13 / D15 – final route unknown)", key="waste_hwcn_interim")
        st.checkbox("Include a non-hazardous line", key="waste_hwcn_non_hazardous_line")
    elif document_type == "waste_spreadsheet":
        st.checkbox("Mixed units (kg / t / KG / bins)", key="waste_log_mixed_units")
        st.checkbox("Volume-only skip rows (not weighed)", key="waste_log_volume_rows")
        st.checkbox("Blank rows", key="waste_log_blank_rows")
        st.checkbox("TOTAL row", key="waste_log_total_row")
        st.checkbox("Hazardous row (asbestos) on a non-HWCN document", key="waste_log_hazardous_row")
        if s.get("sidebar_format") == "XLSX":
            st.checkbox("Extra Notes sheet", key="waste_log_notes_sheet")
    st.divider()


_OPTION_KEYS = [
    "wtn_mode", "wtn_quantity_style", "wtn_rd_in_destination",
    "inv_role", "inv_basis", "inv_excess_weight", "inv_billing_period", "inv_show_routes", "inv_rebate", "inv_landfill_tax", "inv_non_waste_charges",
    "inv_unitless_weight", "inv_eu_format", "inv_no_waste_lines", "inv_two_month_period", "inv_weighbridge_refs",
    "wb_unit", "wb_mode",
    "rep_route_split", "rep_multi_site", "rep_totals_row",
    "hwcn_part_e_differs", "hwcn_interim", "hwcn_non_hazardous_line",
    "log_mixed_units", "log_volume_rows", "log_blank_rows", "log_total_row", "log_hazardous_row", "log_notes_sheet",
]


def _init_defaults() -> None:
    from generators.waste_generator import DEFAULT_OPTIONS

    s = st.session_state
    s.setdefault("waste_n_companies", 1)
    s.setdefault("waste_doc_count", 1)
    s.setdefault("waste_vat_rate", 20)
    s.setdefault("waste_fill_factor", 1.0)
    for key in _OPTION_KEYS:
        s.setdefault(f"waste_{key}", DEFAULT_OPTIONS[key])


# ── organisations, contractors, sites, streams ───────────────────────────────


def _render_companies(document_type: str | None) -> None:
    st.markdown("#### Organisations, Contractors & Sites")
    st.number_input("Number of organisations", min_value=1, max_value=10, step=1, key="waste_n_companies")
    for i in range(int(st.session_state.get("waste_n_companies", 1))):
        company = _company_default(i)
        contractor = WASTE_CONTRACTORS[i % len(WASTE_CONTRACTORS)]
        key = f"waste_co_{i}"
        with st.expander(f"Organisation {i + 1}", expanded=i == 0):
            col1, col2 = st.columns(2)
            with col1:
                st.text_input("Organisation (waste producer)", value=company.get("name", f"Organisation {i + 1}"), key=f"{key}_name")
                st.text_input("Organisation Code", value=company.get("code", f"ORG{i + 1}"), key=f"{key}_code")
                st.text_input("SIC code (2007)", value=SIC_CODES[i % len(SIC_CODES)], key=f"{key}_sic_code")
                st.text_input("Council / unitary authority", value=COUNCILS[i % len(COUNCILS)], key=f"{key}_council")
                if document_type_requires_company_currency(_CATEGORY, document_type or ""):
                    default_currency = _CURRENCY_DISPLAY.get(company.get("currency", "GBP"), "GBP (£)")
                    st.selectbox("Currency", options=_currency_options(),
                                 index=_currency_index(st.session_state.get(f"{key}_currency", default_currency)),
                                 key=f"{key}_currency")
            with col2:
                st.text_input("Waste contractor / carrier", value=contractor["name"], key=f"{key}_contractor_name")
                st.text_input("Contractor Code", value=contractor["code"], key=f"{key}_contractor_code")
                st.text_input("Waste carrier registration no.", value=contractor["carrier_reg"], key=f"{key}_contractor_reg")
                st.text_area("Contractor Address", value="\n".join(contractor["address"]), key=f"{key}_contractor_address", height=100)

            st.session_state.setdefault(f"waste_n_sites_{i}", 2)
            st.number_input("Number of sites", min_value=1, max_value=10, step=1, key=f"waste_n_sites_{i}")
            for j in range(int(st.session_state.get(f"waste_n_sites_{i}", 2))):
                _render_site(i, j)


def _render_site(i: int, j: int) -> None:
    site = _site_default(i, j)
    key = f"waste_s_{i}_{j}"
    st.markdown(f"**Site {j + 1}**")
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Site Label", value=site.get("label", f"Site {j + 1}"), key=f"{key}_label")
        st.text_area("Site Address", value=site.get("address", ""), key=f"{key}_address", height=100)
    with col2:
        defaults = DEFAULT_SITE_STREAMS[(i + j) % len(DEFAULT_SITE_STREAMS)]
        streams = st.multiselect("Waste streams", options=list(STREAM_PRESETS), default=defaults,
                                 format_func=lambda k: _STREAM_LABELS[k], key=f"{key}_streams",
                                 help="Hazardous streams (*) are used for consignment notes; the others for all other documents.")
        for stream in streams:
            rng = random.Random(f"waste_lifts:{i}:{j}:{stream}")
            default_lifts = rng.choice([1, 2, 4, 8, 13]) if STREAM_PRESETS[stream]["container"][0] == "bin" else rng.choice([1, 2, 3])
            st.number_input(f"Lifts / collections per month – {_STREAM_LABELS[stream]}", min_value=1, max_value=60, step=1,
                            value=default_lifts, key=f"{key}_lifts_{stream}")
    st.divider()


def _lines(text: str) -> list[str]:
    return [line.strip() for line in str(text or "").split("\n") if line.strip()]


def _collect_companies(document_type: str | None) -> list[dict]:
    s = st.session_state
    companies = []
    for i in range(int(s.get("waste_n_companies", 1))):
        key = f"waste_co_{i}"
        sites = []
        for j in range(int(s.get(f"waste_n_sites_{i}", 2))):
            site_key = f"waste_s_{i}_{j}"
            sites.append({
                "label": s.get(f"{site_key}_label", ""),
                "address": _lines(s.get(f"{site_key}_address", "")),
                "streams": [
                    {"preset": stream, "lifts_per_month": int(s.get(f"{site_key}_lifts_{stream}", 4) or 4)}
                    for stream in s.get(f"{site_key}_streams", [])
                ],
            })
        companies.append({
            "label": s.get(f"{key}_name", "") or f"Organisation {i + 1}",
            "name": s.get(f"{key}_name", ""),
            "code": s.get(f"{key}_code", ""),
            "sic_code": s.get(f"{key}_sic_code", ""),
            "council": s.get(f"{key}_council", ""),
            "currency": (s.get(f"{key}_currency", "GBP (£)")
                         if document_type_requires_company_currency(_CATEGORY, document_type or "") else "GBP (£)"),
            "contractor": {
                "name": s.get(f"{key}_contractor_name", ""),
                "code": s.get(f"{key}_contractor_code", ""),
                "carrier_reg": s.get(f"{key}_contractor_reg", ""),
                "address": _lines(s.get(f"{key}_contractor_address", "")),
            },
            "sites": sites,
        })
    return companies


def render_waste_form(document_type: str | None) -> dict:
    st.subheader("Waste Generated in Operations")
    captions = {
        "waste_transfer_note": "Duty of care waste transfer note (official A–D layout) – single transfer or season ticket.",
        "waste_contractor_invoice": "Waste contractor invoice or credit note with waste, rebate, landfill tax and non-waste lines.",
        "weighbridge_ticket": "Weighbridge ticket with gross / tare / net weights.",
        "waste_report": "Contractor waste & diversion report or portal export, per stream and route.",
        "hazardous_waste_consignment_note": "Hazardous waste consignment note (Parts A–E).",
        "waste_spreadsheet": "Customer-maintained skip / waste collection log.",
    }
    st.caption(captions.get(document_type, "Scope 3 Category 5 waste document configuration."))

    _init_defaults()
    _render_document_settings(document_type)
    _render_financial_period()
    _render_special_cases(document_type)
    _render_companies(document_type)

    s = st.session_state
    fp_start: date = s.get("fp_start") or date(2026, 1, 1)
    fp_end: date = s.get("fp_end") or date(2026, 1, 31)
    title, subject = _document_defaults(document_type)
    options = {key: s.get(f"waste_{key}") for key in _OPTION_KEYS}
    options["fill_factor"] = float(s.get("waste_fill_factor", 1.0) or 1.0)

    return {
        "_category": _CATEGORY,
        "document_type": document_type or "waste_contractor_invoice",
        "doc_count": int(s.get("waste_doc_count", 1) or 1) if document_type in _MULTI_DOCUMENT_TYPES else 1,
        "waste_vat_rate": int(s.get("waste_vat_rate", 20) or 0),
        "waste_options": options,
        "doc_title": s.get("doc_title", title),
        "doc_subject": s.get("doc_subject", subject),
        "doc_seed": int(s.get("doc_seed", 20261116)),
        "fp_label": s.get("fp_label", ""),
        "fp_start": fp_start.isoformat(),
        "fp_end": fp_end.isoformat(),
        "doc_language": "en",
        "doc_noise": 0.0,
        "doc_inject_special_chars": bool(s.get("doc_inject_special_chars", False)),
        "companies": _collect_companies(document_type),
        **_collect_distractor_fields(s),
        **_collect_layout(s),
        **_collect_invalid_data(s),
    }
