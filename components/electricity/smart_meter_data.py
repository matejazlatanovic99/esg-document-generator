from __future__ import annotations

import hashlib
import json
import os
import random
from calendar import monthrange
from datetime import date

import streamlit as st

from components.sidebar import get_document_type_config
from utils.currency import currency_code, currency_index, currency_options

_LANGUAGE_OPTIONS: dict[str, str] = {
    "English": "en",
    "French (Français)": "fr",
    "German (Deutsch)": "de",
    "Dutch (Nederlands)": "nl",
}

_DEFAULT_COMPANY = {
    "label": "Smart Meter Portfolio",
    "supplier": "Smart Meter Export",
    "supplier_code": "SMX",
    "supplier_address": ["Digital Meter Export"],
    "customer": "Metered Sites",
    "customer_code": "MTR",
    "currency": "GBP (£)",
    "accent": "#1E5B88",
    "_omit": {},
}

_SMART_METER_TARIFF_POOLS = [
    ["Day/Night"],
    ["Peak", "Off-Peak"],
    ["Standard", "Night"],
    ["Day", "Evening", "Night"],
    ["Weekday", "Weekend"],
]

_CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "config")


def _load_json(filename: str):
    path = os.path.join(_CONFIG_DIR, filename)
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


_SITES_CONFIG: list = _load_json("sites.json") or []


def _document_defaults(category: str, document_type: str | None) -> tuple[str, str]:
    cfg = get_document_type_config(category, document_type or "")
    return cfg.get("default_title", "Document"), cfg.get("default_subject", "")


def _sync_document_setting_defaults(category: str, document_type: str | None) -> None:
    selection_key = f"smart_meter:{category}:{document_type or ''}"
    if st.session_state.get("_document_settings_selection") == selection_key:
        return

    default_title, default_subject = _document_defaults(category, document_type)
    st.session_state["doc_title"] = default_title
    st.session_state["doc_subject"] = default_subject
    st.session_state["_document_settings_selection"] = selection_key


def _rand_financial_period() -> tuple[date, date, str]:
    current_year = 2026
    year = random.randint(current_year - 4, current_year)
    start_month = random.randint(1, 12)
    n_months = random.randint(1, 12)
    start = date(year, start_month, 1)
    end_year = year
    end_month = start_month + n_months - 1
    while end_month > 12:
        end_month -= 12
        end_year += 1
    last_day = monthrange(end_year, end_month)[1]
    end = date(end_year, end_month, last_day)
    label = f"Financial Year {year}" if end_year == year else f"Financial Period {year}–{end_year}"
    return start, end, label


def _render_financial_period() -> None:
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
        return

    total_days = (fp_end - fp_start).days + 1
    st.caption(f"Period spans {total_days} day(s).")


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
            )

        st.selectbox(
            "Document Language",
            options=list(_LANGUAGE_OPTIONS.keys()),
            key="doc_language_label",
        )
        st.checkbox(
            "Inject special characters",
            key="doc_inject_special_chars",
            help="Append QA characters to generated text values.",
        )


def _init_smart_meter_defaults() -> None:
    if "smart_meter_tariff_pool" not in st.session_state:
        st.session_state["smart_meter_tariff_pool"] = random.choice(_SMART_METER_TARIFF_POOLS)


def _smart_meter_tariff_default(idx: int) -> str:
    pool = st.session_state.get("smart_meter_tariff_pool", _SMART_METER_TARIFF_POOLS[0])
    if idx < len(pool):
        return pool[idx]
    return f"Tariff {idx + 1}"


def _smart_meter_consumption_default(idx: int) -> float:
    seed = int(st.session_state.get("doc_seed", 20260325))
    rng = random.Random(f"{seed}:smart_meter_total_consumption:{idx}")
    return float(round(rng.uniform(12_000.0, 95_000.0), 2))


def _estimated_total_cost(total_quantity: float, meter_id: str, seed: int) -> float:
    rng = random.Random(f"{seed}:{meter_id}:cost")
    unit_rate = rng.uniform(0.22, 0.38)
    return round(total_quantity * unit_rate, 2)


def _smart_meter_total_cost_default(idx: int) -> float:
    seed = int(st.session_state.get("doc_seed", 20260325))
    meter_id = st.session_state.get(f"smart_meter_meter_id_{idx}", _default_meter_id(idx))
    total_quantity = float(
        st.session_state.get(
            f"smart_meter_total_consumption_{idx}",
            _smart_meter_consumption_default(idx),
        )
    )
    return _estimated_total_cost(total_quantity, meter_id, seed)


def _config_site(idx: int) -> dict:
    if not _SITES_CONFIG:
        return {}
    raw = f"site:0:{idx}".encode()
    site_idx = int(hashlib.sha1(raw).hexdigest()[:4], 16) % len(_SITES_CONFIG)
    return _SITES_CONFIG[site_idx]


def _smart_meter_site_default(idx: int, field: str, fallback: str = "") -> str:
    site = _config_site(idx)
    return site.get(field, fallback)


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


def _render_data_settings() -> None:
    st.markdown("#### Data Settings")
    granularity = st.radio(
        "Smart Meter Data Type",
        options=["Monthly", "Interval"],
        horizontal=True,
        key="smart_meter_data_granularity_label",
    )

    if granularity == "Interval":
        col1, col2, col3 = st.columns(3)
        with col1:
            st.selectbox(
                "Interval Split",
                options=["15 minutes", "30 minutes", "60 minutes"],
                key="smart_meter_interval_label",
            )
        with col2:
            st.radio(
                "Interval Value Mode",
                options=["Consumption Diff", "End Reading (Cumulated)"],
                horizontal=True,
                key="smart_meter_interval_value_mode_label",
            )
        with col3:
            st.selectbox(
                "Timestamp Format",
                options=["ISO 8601 UTC", "Datetime"],
                key="smart_meter_timestamp_format_label",
            )
    else:
        st.selectbox(
            "Currency",
            options=currency_options(),
            index=currency_index(_DEFAULT_COMPANY["currency"]),
            key="smart_meter_currency",
        )
        st.caption("Monthly output includes one row per period and tariff type.")


def _render_tariff_types() -> None:
    _init_smart_meter_defaults()
    st.markdown("#### Tariff Types")
    n_tariffs = st.number_input(
        "Number of tariff types",
        min_value=0,
        max_value=10,
        value=len(st.session_state.get("smart_meter_tariff_pool", _SMART_METER_TARIFF_POOLS[0])),
        step=1,
        key="smart_meter_n_tariffs",
    )
    if int(n_tariffs) == 0:
        st.caption("No tariff breakdown will be generated.")
        return

    for idx in range(int(n_tariffs)):
        st.text_input(
            f"Tariff {idx + 1}",
            value=_smart_meter_tariff_default(idx),
            key=f"smart_meter_tariff_{idx}",
        )


def _default_meter_id(idx: int) -> str:
    return f"SM-{1001 + idx}"


def _render_meter_inputs() -> None:
    st.markdown("#### Meter IDs")
    meter_count = st.number_input(
        "Number of meters",
        min_value=1,
        max_value=50,
        value=1,
        step=1,
        key="smart_meter_n_meters",
    )

    granularity = st.session_state.get("smart_meter_data_granularity_label", "Monthly")
    show_total_cost = granularity == "Monthly"
    for idx in range(int(meter_count)):
        meter_id = st.session_state.get(f"smart_meter_meter_id_{idx}", _default_meter_id(idx))
        with st.expander(f"Meter {idx + 1}: {meter_id}", expanded=(idx == 0)):
            meter_cols = st.columns(3 if show_total_cost else 2)
            col1, col2 = meter_cols[:2]
            with col1:
                st.text_input(
                    "Meter ID",
                    value=_default_meter_id(idx),
                    key=f"smart_meter_meter_id_{idx}",
                )
            with col2:
                st.number_input(
                    "Total Consumption (kWh)",
                    min_value=0.0,
                    step=100.0,
                    format="%.2f",
                    value=_smart_meter_consumption_default(idx),
                    key=f"smart_meter_total_consumption_{idx}",
                )
            if show_total_cost:
                with meter_cols[2]:
                    currency = currency_code(
                        st.session_state.get("smart_meter_currency", _DEFAULT_COMPANY["currency"])
                    )
                    cost_omit_key = f"smart_meter_total_cost_{idx}_omit"
                    st.number_input(
                        f"Total Cost ({currency})",
                        min_value=0.0,
                        step=100.0,
                        format="%.2f",
                        value=_smart_meter_total_cost_default(idx),
                        key=f"smart_meter_total_cost_{idx}",
                        disabled=bool(st.session_state.get(cost_omit_key, False)),
                    )
                    st.checkbox(
                        "Omit total cost",
                        key=cost_omit_key,
                        help="Leave currency and cost blank in monthly smart meter exports.",
                    )

            if granularity == "Monthly":
                site_key = f"smart_meter_site_label_{idx}"
                default_site_label = _smart_meter_site_default(idx, "label", "")
                if site_key not in st.session_state or (
                    not st.session_state.get(site_key)
                    and f"{site_key}_omit" not in st.session_state
                ):
                    st.session_state[site_key] = default_site_label
                _field(
                    st.text_input,
                    "Site",
                    site_key,
                    help="Site label shown in monthly smart meter output.",
                )


def _estimated_start_reading(meter_id: str, seed: int) -> int:
    raw = f"{seed}:{meter_id}".encode()
    return int(hashlib.sha1(raw).hexdigest()[:6], 16) % 900_000 + 50_000


def _collect_form_data(document_type: str | None) -> dict:
    s = st.session_state
    default_title, default_subject = _document_defaults("electricity", document_type)
    seed = int(s.get("doc_seed", 20260325))
    granularity = s.get("smart_meter_data_granularity_label", "Monthly").lower()
    interval_minutes = int(str(s.get("smart_meter_interval_label", "30 minutes")).split()[0])
    interval_value_mode = s.get("smart_meter_interval_value_mode_label", "Consumption Diff")
    timestamp_format = s.get("smart_meter_timestamp_format_label", "ISO 8601 UTC")
    meter_count = int(s.get("smart_meter_n_meters", 1))
    tariff_count = int(s.get("smart_meter_n_tariffs", 0))

    tariffs = []
    for idx in range(tariff_count):
        name = s.get(f"smart_meter_tariff_{idx}", "").strip()
        if name:
            tariffs.append({"name": name})

    sites = []
    for idx in range(meter_count):
        meter_id = s.get(f"smart_meter_meter_id_{idx}", "").strip() or _default_meter_id(idx)
        total_quantity = float(s.get(f"smart_meter_total_consumption_{idx}", 0.0))
        if granularity == "monthly":
            site_label = (
                s.get(f"smart_meter_site_label_{idx}", "").strip()
                or _smart_meter_site_default(idx, "label", "")
            )
            omit_site_label = bool(s.get(f"smart_meter_site_label_{idx}_omit", False))
        else:
            site_label = ""
            omit_site_label = True
        sites.append({
            "label": site_label,
            "customer_address": [],
            "city": "",
            "postcode": "",
            "meter_id": meter_id,
            "supplier_ef": "0",
            "unit": "kWh",
            "start_reading": _estimated_start_reading(meter_id, seed),
            "total_quantity": str(total_quantity),
            "total_cost": str(
                s.get(
                    f"smart_meter_total_cost_{idx}",
                    _estimated_total_cost(total_quantity, meter_id, seed),
                )
            ),
            "tariffs": tariffs,
            "_omit": {
                "address": True,
                "city": True,
                "postcode": True,
                "label": omit_site_label,
                "total_cost": bool(s.get(f"smart_meter_total_cost_{idx}_omit", False)),
            },
        })

    companies = [{**_DEFAULT_COMPANY, "currency": s.get("smart_meter_currency", _DEFAULT_COMPANY["currency"]), "sites": sites}]
    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end: date = s.get("fp_end", date(2026, 12, 31))

    return {
        "_category": "electricity",
        "document_type": document_type or "smart_meter_data",
        "doc_title": s.get("doc_title", default_title),
        "doc_subject": s.get("doc_subject", default_subject),
        "doc_seed": seed,
        "fp_label": s.get("fp_label", "Financial Year 2026"),
        "fp_start": fp_start.isoformat(),
        "fp_end": fp_end.isoformat(),
        "doc_language": _LANGUAGE_OPTIONS.get(s.get("doc_language_label", "English"), "en"),
        "doc_noise": 0.0,
        "doc_inject_special_chars": bool(s.get("doc_inject_special_chars", False)),
        "smart_meter_data_granularity": granularity,
        "smart_meter_interval_minutes": interval_minutes,
        "smart_meter_interval_value_mode": "cumulative_end_reading" if interval_value_mode == "End Reading (Cumulated)" else "consumption_diff",
        "smart_meter_timestamp_format": "datetime" if timestamp_format == "Datetime" else "iso_8601_utc",
        "xlsx_split_by_company": False,
        "companies": companies,
    }


def render_smart_meter_data_form(document_type: str | None) -> dict:
    st.subheader("Electricity")
    st.caption("Smart meter export configuration for Scope 2 purchased electricity.")

    _render_document_settings("electricity", document_type)
    _render_financial_period()
    _render_data_settings()
    _render_tariff_types()
    _render_meter_inputs()

    return _collect_form_data(document_type)
