from __future__ import annotations

import hashlib
import json
import os
import random
from calendar import monthrange
from datetime import date

import streamlit as st

from components.sidebar import NEW_COMPANY_PLACEHOLDER, get_document_type_config
from components.stationary_combustion.bems import (
    bems_report_type_value as _bems_report_type_value_module,
    render_bems_defaults as _render_bems_defaults_module,
    render_bems_report_type_selector as _render_bems_report_type_selector_module,
    render_bems_site_fields as _render_bems_site_fields_module,
)
from components.stationary_combustion.delivery_note import (
    render_delivery_note_defaults as _render_delivery_note_defaults_module,
)
from components.stationary_combustion.fuel_card import (
    render_fuel_card_defaults as _render_fuel_card_defaults_module,
)
from components.stationary_combustion.fuel_invoice import (
    render_invoice_defaults as _render_invoice_defaults_module,
)
from components.stationary_combustion.generator_log import (
    render_log_defaults as _render_log_defaults_module,
)
from components.stationary_combustion.units import (
    COMPACT_FUEL_VOLUME_UNITS,
    DEFAULT_COMPACT_FUEL_VOLUME_UNIT,
    DEFAULT_LONG_FUEL_VOLUME_UNIT,
    LONG_FUEL_VOLUME_UNITS,
    default_fuel_volume_unit,
    option_index,
)
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

_BEMS_REPORT_TYPES: dict[str, str] = {
    "Equipment Trend Report": "equipment_trend_report",
    "Time-Series Trend Export": "time_series_trend_export",
}

_CONFIG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "config")

_STATIONARY_SUPPLIERS = [
    {
        "name": "ABC Fuels Ltd",
        "code": "ABF",
        "address": "Fuel Distribution Centre\n21 River Port Way\nLiverpool L21 4AA\nUnited Kingdom",
    },
    {
        "name": "Emerald Fuel Services Ltd",
        "code": "EFS",
        "address": "Docklands Energy Park\n9 Marina Road\nCork T12 H2K8\nIreland",
    },
    {
        "name": "Nordic Industrial Fuels A/S",
        "code": "NIF",
        "address": "Generator Logistics Hub\n18 Havnevej\n2750 Ballerup\nDenmark",
    },
    {
        "name": "Continental Fuel Supply GmbH",
        "code": "CFS",
        "address": "Dieselring 14\n50858 Cologne\nGermany",
    },
]

_FUEL_CARD_PROVIDERS = [
    {
        "name": "WEX Europe Services Ltd",
        "code": "WEX",
        "address": "3rd Floor\n40 Mespil Road\nDublin 4\nIreland",
    },
    {
        "name": "Radius Payment Solutions Ltd",
        "code": "RADIUS",
        "address": "Euro House\nJunction Business Park\nDublin\nIreland",
    },
    {
        "name": "DKV Mobility Services",
        "code": "DKV",
        "address": "Balcke-Durr-Allee 3\n40882 Ratingen\nGermany",
    },
]

_FUELS = [
    "Gas Oil (Red Diesel)",
    "Diesel",
    "HVO",
    "Kerosene",
    "Heating Oil",
    "Fuel Oil",
    "Biodiesel",
    "LPG",
    "Propane",
    "Butane",
]

_DELIVERY_TANKS = [
    "Boiler Tank 1",
    "Generator Tank A",
    "Heating Oil Tank 2",
    "Standby Fuel Tank",
]

_FUEL_CARD_MERCHANTS = [
    "Fuel Depot Cork",
    "Industrial Energy Supply Dublin",
    "Harbour Fuel Services",
    "Generator Fuels NI",
]

_EMISSION_SOURCES = [
    "Backup Generator",
    "Emergency Generator",
    "Stationary Generator",
    "Standby Power Unit",
]

_BEMS_ASSET_DEFAULTS = [
    {
        "asset_tag": "BLR-01",
        "equipment_name": "Main Boiler 1",
        "emission_source": "Boiler",
        "fuel": "Natural Gas",
        "unit": "kWh",
        "sensor_name": "Gas Consumption",
        "quantity": 12850.0,
        "operating_hours": 210.0,
    },
    {
        "asset_tag": "BLR-02",
        "equipment_name": "Main Boiler 2",
        "emission_source": "Boiler",
        "fuel": "Natural Gas",
        "unit": "kWh",
        "sensor_name": "Gas Consumption",
        "quantity": 8420.0,
        "operating_hours": 144.0,
    },
    {
        "asset_tag": "GEN-01",
        "equipment_name": "Backup Generator",
        "emission_source": "Backup Generator",
        "fuel": "Diesel",
        "unit": "L",
        "sensor_name": "Fuel Consumption",
        "quantity": 180.0,
        "operating_hours": 8.2,
    },
]

def _load_json(filename: str):
    path = os.path.join(_CONFIG_DIR, filename)
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


_COMPANIES_CONFIG: list[dict] = _load_json("companies.json") or []
_SITES_CONFIG: list[dict] = _load_json("sites.json") or []

def _document_defaults(document_type: str | None) -> tuple[str, str]:
    cfg = get_document_type_config("stationary_combustion", document_type or "")
    default_title = cfg.get("default_title", "Document")
    default_subject = cfg.get("default_subject", "")
    if document_type == "bems":
        if _bems_report_type_value_module(_BEMS_REPORT_TYPES) == "time_series_trend_export":
            return "BEMS Time-Series Trend Export", "Scope 1 stationary combustion BEMS time-series export"
        return "BEMS Fuel Consumption Summary", "Scope 1 stationary combustion BEMS equipment trend report"
    return default_title, default_subject


def _sync_document_setting_defaults(document_type: str | None) -> None:
    selection_key = (
        f"stationary_combustion:{document_type or ''}:"
        f"{_bems_report_type_value_module(_BEMS_REPORT_TYPES) if document_type == 'bems' else ''}"
    )
    if st.session_state.get("_document_settings_selection") == selection_key:
        return

    default_title, default_subject = _document_defaults(document_type)
    st.session_state["doc_title"] = default_title
    st.session_state["doc_subject"] = default_subject
    st.session_state["_document_settings_selection"] = selection_key


def _rand_financial_period() -> tuple[date, date, str]:
    current_year = 2026
    year = random.randint(current_year - 4, current_year)
    month = random.randint(1, 12)
    start = date(year, month, 1)
    end = date(year, month, monthrange(year, month)[1])
    return start, end, start.strftime("%B %Y")


def _render_document_settings(document_type: str | None) -> None:
    _sync_document_setting_defaults(document_type)
    default_title, default_subject = _document_defaults(document_type)

    with st.expander("Document Settings", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            st.text_input("Document Title", value=default_title, key="doc_title")
        with col2:
            st.text_input("Document Subject", value=default_subject, key="doc_subject")
            st.number_input(
                "Random Seed",
                value=20260415,
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

        if st.session_state.get("sidebar_format", "PDF") in {"PDF", "DOCX"}:
            st.slider(
                "Scan noise level",
                min_value=0.0,
                max_value=1.0,
                value=0.0,
                step=0.05,
                key="doc_noise",
                help="Retained for consistency with other document generators.",
            )


def _render_financial_period() -> None:
    if "fp_start" not in st.session_state:
        fp_start_rand, fp_end_rand, fp_label_rand = _rand_financial_period()
        st.session_state["fp_start"] = fp_start_rand
        st.session_state["fp_end"] = fp_end_rand
        st.session_state["fp_label"] = fp_label_rand

    st.markdown("#### Reporting Period")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        st.text_input("Period Label", key="fp_label")
    with col2:
        st.date_input("Start Date", key="fp_start")
    with col3:
        st.date_input("End Date", key="fp_end")

    fp_start: date = st.session_state.get("fp_start", date(2026, 1, 1))
    fp_end: date = st.session_state.get("fp_end", date(2026, 1, 31))
    if fp_end < fp_start:
        st.error("End date must be after start date.")
        return
    st.caption(f"Period spans {(fp_end - fp_start).days + 1} day(s).")


def _config_company(i: int) -> dict:
    if not _COMPANIES_CONFIG:
        return {}
    return _COMPANIES_CONFIG[i % len(_COMPANIES_CONFIG)]


def _config_site(i: int, j: int) -> dict:
    if not _SITES_CONFIG:
        return {}
    raw = f"stationary-site:{i}:{j}".encode()
    idx = int(hashlib.sha1(raw).hexdigest()[:4], 16) % len(_SITES_CONFIG)
    return _SITES_CONFIG[idx]


def _supplier_default(i: int) -> dict:
    return _STATIONARY_SUPPLIERS[i % len(_STATIONARY_SUPPLIERS)]


def _provider_default(document_type: str | None, i: int) -> dict:
    if document_type == "fuel_card":
        return _FUEL_CARD_PROVIDERS[i % len(_FUEL_CARD_PROVIDERS)]
    return _supplier_default(i)


def _country_from_address(address: str) -> str:
    lines = [line.strip() for line in str(address).splitlines() if line.strip()]
    return lines[-1] if lines else ""


def _company_default(i: int, field: str, fallback: str = "", document_type: str | None = None) -> str:
    company = _config_company(i)
    supplier = _provider_default(document_type, i)
    if field == "label":
        return company.get("name", fallback)
    if field == "customer":
        return company.get("name", fallback)
    if field == "customer_code":
        return company.get("code", fallback)
    if field == "currency":
        return _CURRENCY_DISPLAY.get(company.get("currency", "EUR"), "EUR (€)")
    if field == "supplier":
        return supplier.get("name", fallback)
    if field == "supplier_code":
        return supplier.get("code", fallback)
    if field == "supplier_address":
        return supplier.get("address", fallback)
    return fallback


def _site_default(i: int, j: int, field: str, fallback=None):
    site = _config_site(i, j)
    if field == "label":
        return site.get("label", fallback)
    if field == "address":
        return site.get("address", fallback)
    if field == "country":
        return _country_from_address(site.get("address", ""))
    if field == "equipment":
        return f"Generator GEN-{i + 1:02d}-{j + 1:02d}"
    if field == "emission_source":
        return _EMISSION_SOURCES[(i + j) % len(_EMISSION_SOURCES)]
    if field == "fuel":
        return _FUELS[(i + j) % len(_FUELS)]
    return fallback


def _bems_asset_default(asset_idx: int, field: str, fallback=None):
    asset = _BEMS_ASSET_DEFAULTS[asset_idx % len(_BEMS_ASSET_DEFAULTS)]
    return asset.get(field, fallback)


def _equipment_default(i: int, j: int, document_type: str | None, equipment_idx: int) -> str:
    if document_type == "delivery_note":
        return _DELIVERY_TANKS[(i + j + equipment_idx) % len(_DELIVERY_TANKS)]
    if equipment_idx == 0:
        return _site_default(i, j, "equipment", "")
    return f"Generator GEN-{i + 1:02d}-{j + 1:02d}-{equipment_idx + 1:02d}"


def _emission_source_default(i: int, j: int, equipment_idx: int) -> str:
    return _EMISSION_SOURCES[(i + j + equipment_idx) % len(_EMISSION_SOURCES)]


def _equipment_label(document_type: str | None) -> str:
    if document_type == "delivery_note":
        return "Tank / Equipment"
    if document_type == "fuel_card":
        return "Equipment / Alias"
    return "Equipment"


def _equipment_help(document_type: str | None) -> str | None:
    if document_type == "delivery_note":
        return "Delivery notes sometimes name a tank or equipment location, but not always."
    if document_type == "fuel_invoice":
        return "Invoices can identify the site without naming a specific asset or tank."
    if document_type == "fuel_card":
        return "Can come from card alias, cost center, or transaction notes."
    return None


def _item_key(i: int, j: int, equipment_idx: int) -> str:
    return f"stationary_site_{i}_{j}_equipment_{equipment_idx}"


def _legacy_site_value(i: int, j: int, field: str, default=None):
    return st.session_state.get(f"stationary_site_{i}_{j}_{field}", default)


def _entry_default_rng(document_type: str, field: str, i: int, j: int, equipment_idx: int) -> random.Random:
    seed = int(st.session_state.get("doc_seed", 20260415))
    return random.Random(f"{seed}:{document_type}_default:{field}:{i}:{j}:{equipment_idx}")


def _document_default_prefix(document_type: str) -> str:
    return {
        "fuel_invoice": "stationary_invoice",
        "delivery_note": "stationary_delivery_note",
        "fuel_card": "stationary_fuel_card",
        "generator_log": "stationary_log",
    }[document_type]


def _entry_default_is_randomized(document_type: str, field: str) -> bool:
    prefix = _document_default_prefix(document_type)
    return bool(st.session_state.get(f"{prefix}_{field}_randomize", False))


def _entry_fuel_default(document_type: str, i: int, j: int, equipment_idx: int) -> str:
    prefix = _document_default_prefix(document_type)
    default_fuel = st.session_state.get(f"{prefix}_fuel", _FUELS[0])
    if _entry_default_is_randomized(document_type, "fuel") and _FUELS:
        return _entry_default_rng(document_type, "fuel", i, j, equipment_idx).choice(_FUELS)
    return default_fuel


def _entry_unit_default(
    document_type: str,
    unit_options,
    fallback_unit: str,
    i: int,
    j: int,
    equipment_idx: int,
) -> str:
    prefix = _document_default_prefix(document_type)
    default_unit = st.session_state.get(f"{prefix}_unit", fallback_unit)
    if _entry_default_is_randomized(document_type, "unit") and unit_options:
        return _entry_default_rng(document_type, "unit", i, j, equipment_idx).choice(tuple(unit_options))
    return default_unit


def _entry_quantity_default(document_type: str, i: int, j: int, equipment_idx: int) -> float:
    if document_type == "fuel_invoice":
        if _entry_default_is_randomized(document_type, "quantity"):
            return float(_entry_default_rng(document_type, "quantity", i, j, equipment_idx).randrange(1200, 6000, 50))
        return float(st.session_state.get("stationary_invoice_quantity", 2500.0))
    if document_type == "delivery_note":
        if _entry_default_is_randomized(document_type, "quantity"):
            return float(_entry_default_rng(document_type, "quantity", i, j, equipment_idx).randrange(1500, 5000, 50))
        return float(st.session_state.get("stationary_delivery_note_quantity", 3200.0))
    if document_type == "fuel_card":
        if _entry_default_is_randomized(document_type, "quantity"):
            return float(_entry_default_rng(document_type, "quantity", i, j, equipment_idx).randrange(150, 700, 10))
        return float(st.session_state.get("stationary_fuel_card_quantity", 250.0))
    return 0.0


def _entry_unit_price_default(document_type: str, i: int, j: int, equipment_idx: int) -> float:
    if document_type == "fuel_invoice":
        if _entry_default_is_randomized(document_type, "unit_price"):
            return round(_entry_default_rng(document_type, "unit_price", i, j, equipment_idx).uniform(0.88, 1.35), 2)
        return float(st.session_state.get("stationary_invoice_unit_price", 1.12))
    if document_type == "fuel_card":
        if _entry_default_is_randomized(document_type, "unit_price"):
            return round(_entry_default_rng(document_type, "unit_price", i, j, equipment_idx).uniform(0.95, 1.45), 2)
        return float(st.session_state.get("stationary_fuel_card_unit_price", 1.24))
    return 0.0


def _entry_delivery_charge_default(i: int, j: int, equipment_idx: int) -> float:
    if _entry_default_is_randomized("fuel_invoice", "delivery_charge"):
        return round(_entry_default_rng("fuel_invoice", "delivery_charge", i, j, equipment_idx).uniform(20.0, 95.0), 2)
    return float(st.session_state.get("stationary_invoice_delivery_charge", 50.0))


def _company_currency_code(i: int = 0, document_type: str | None = None) -> str:
    selected = (
        st.session_state.get(f"stationary_co_{i}_currency")
        or _company_default(i, "currency", "GBP (£)", document_type=document_type)
    )
    return _currency_code(selected)


def _entry_runs_per_month_default(i: int, j: int, equipment_idx: int) -> int:
    if _entry_default_is_randomized("generator_log", "runs_per_month"):
        return _entry_default_rng("generator_log", "runs_per_month", i, j, equipment_idx).randint(2, 5)
    return int(st.session_state.get("stationary_log_runs_per_month", 3))


def _entry_tank_capacity_default(i: int, j: int, equipment_idx: int) -> float:
    if _entry_default_is_randomized("generator_log", "tank_capacity"):
        return float(_entry_default_rng("generator_log", "tank_capacity", i, j, equipment_idx).randrange(400, 1600, 50))
    return float(st.session_state.get("stationary_log_tank_capacity", 800.0))


def _entry_fuel_used_per_hour_default(i: int, j: int, equipment_idx: int) -> float:
    if _entry_default_is_randomized("generator_log", "fuel_used_per_hour"):
        return round(_entry_default_rng("generator_log", "fuel_used_per_hour", i, j, equipment_idx).uniform(9.0, 28.0), 1)
    return float(st.session_state.get("stationary_log_fuel_used_per_hour", 15.0))


def _entry_run_hours_min_default(i: int, j: int, equipment_idx: int) -> float:
    if _entry_default_is_randomized("generator_log", "run_hours_min"):
        return _entry_default_rng("generator_log", "run_hours_min", i, j, equipment_idx).randrange(1, 9) / 4
    return float(st.session_state.get("stationary_log_run_hours_min", 0.5))


def _entry_run_hours_max_default(i: int, j: int, equipment_idx: int) -> float:
    if _entry_default_is_randomized("generator_log", "run_hours_max"):
        min_hours = _entry_run_hours_min_default(i, j, equipment_idx)
        lower_bound = max(int(min_hours * 4) + 1, 12)
        upper_bound = max(lower_bound + 1, 33)
        return _entry_default_rng("generator_log", "run_hours_max", i, j, equipment_idx).randrange(lower_bound, upper_bound) / 4
    return float(st.session_state.get("stationary_log_run_hours_max", 5.0))


def _render_fuel_unit_quantity(prefix: str, i: int, j: int, document_type: str, equipment_idx: int) -> None:
    if document_type in {"fuel_invoice", "delivery_note"}:
        unit_options = LONG_FUEL_VOLUME_UNITS
        default_unit = st.session_state.get(
            "stationary_invoice_unit" if document_type == "fuel_invoice" else "stationary_delivery_note_unit",
            DEFAULT_LONG_FUEL_VOLUME_UNIT,
        )
    else:
        unit_options = COMPACT_FUEL_VOLUME_UNITS
        default_unit = st.session_state.get(
            "stationary_fuel_card_unit" if document_type == "fuel_card" else "stationary_log_unit",
            DEFAULT_COMPACT_FUEL_VOLUME_UNIT,
        )

    quantity_steps = {
        "fuel_invoice": 50.0,
        "delivery_note": 50.0,
        "fuel_card": 10.0,
    }
    default_fuel = _entry_fuel_default(document_type, i, j, equipment_idx)
    default_unit = _entry_unit_default(document_type, unit_options, default_unit, i, j, equipment_idx)

    default_quantity = None
    if document_type != "generator_log":
        default_quantity = _entry_quantity_default(document_type, i, j, equipment_idx)

    col1, col2, col3 = st.columns(3)
    with col1:
        legacy_fuel = _legacy_site_value(i, j, "fuel", "") if equipment_idx == 0 else ""
        st.text_input(
            "Fuel",
            value=legacy_fuel or default_fuel,
            key=f"{prefix}_fuel",
        )
    with col2:
        st.selectbox(
            "Unit",
            options=unit_options,
            index=option_index(
                unit_options,
                (_legacy_site_value(i, j, "unit", None) if equipment_idx == 0 else None) or default_unit,
            ),
            key=f"{prefix}_unit",
        )
    if document_type != "generator_log":
        with col3:
            label = "Delivered Quantity" if document_type == "delivery_note" else "Quantity"
            st.number_input(
                label,
                min_value=0.0,
                step=quantity_steps[document_type],
                format="%.2f",
                value=float(
                    (_legacy_site_value(i, j, "quantity", None) if equipment_idx == 0 else None)
                    or default_quantity
                ),
                key=f"{prefix}_quantity",
            )


def _render_invoice_item_fields(prefix: str, i: int, j: int, equipment_idx: int) -> None:
    st.markdown("**Invoice Line Defaults**")
    _render_fuel_unit_quantity(prefix, i, j, "fuel_invoice", equipment_idx)
    unit_price_default = _entry_unit_price_default("fuel_invoice", i, j, equipment_idx)
    delivery_charge_default = _entry_delivery_charge_default(i, j, equipment_idx)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.number_input(
            f"Unit Price ({_company_currency_code(i, 'fuel_invoice')})",
            min_value=0.01,
            step=0.01,
            format="%.2f",
            value=float(
                (_legacy_site_value(i, j, "unit_price", None) if equipment_idx == 0 else None)
                or unit_price_default
            ),
            key=f"{prefix}_unit_price",
        )
    with col2:
        st.number_input(
            f"Delivery Charge ({_company_currency_code(i, 'fuel_invoice')})",
            min_value=0.0,
            step=5.0,
            format="%.2f",
            value=float(
                (_legacy_site_value(i, j, "delivery_charge", None) if equipment_idx == 0 else None)
                or delivery_charge_default
            ),
            key=f"{prefix}_delivery_charge",
        )
    with col3:
        st.number_input(
            "VAT Rate (%)",
            min_value=0,
            max_value=100,
            step=1,
            value=int(
                (_legacy_site_value(i, j, "vat_rate", None) if equipment_idx == 0 else None)
                or st.session_state.get("stationary_invoice_vat_rate", 20)
            ),
            key=f"{prefix}_vat_rate",
        )


def _render_delivery_note_item_fields(prefix: str, i: int, j: int, equipment_idx: int) -> None:
    st.markdown("**Delivery Details**")
    _render_fuel_unit_quantity(prefix, i, j, "delivery_note", equipment_idx)


def _render_fuel_card_item_fields(prefix: str, i: int, j: int, equipment_idx: int) -> None:
    st.markdown("**Transaction Details**")
    _render_fuel_unit_quantity(prefix, i, j, "fuel_card", equipment_idx)
    unit_price_default = _entry_unit_price_default("fuel_card", i, j, equipment_idx)
    st.number_input(
        f"Unit Price ({_company_currency_code(i, 'fuel_card')})",
        min_value=0.01,
        step=0.01,
        format="%.2f",
        value=float(
            (_legacy_site_value(i, j, "unit_price", None) if equipment_idx == 0 else None)
            or unit_price_default
        ),
        key=f"{prefix}_unit_price",
    )


def _render_generator_log_item_fields(prefix: str, i: int, j: int, equipment_idx: int) -> None:
    st.markdown("**Log Generation Defaults**")
    _render_fuel_unit_quantity(prefix, i, j, "generator_log", equipment_idx)
    runs_per_month_default = _entry_runs_per_month_default(i, j, equipment_idx)
    fuel_used_per_hour_default = _entry_fuel_used_per_hour_default(i, j, equipment_idx)
    run_hours_min_default = _entry_run_hours_min_default(i, j, equipment_idx)
    tank_capacity_default = _entry_tank_capacity_default(i, j, equipment_idx)
    run_hours_max_default = _entry_run_hours_max_default(i, j, equipment_idx)
    col1, col2 = st.columns(2)
    with col1:
        st.number_input(
            "Runs Per Month",
            min_value=1,
            max_value=31,
            step=1,
            value=int(
                (_legacy_site_value(i, j, "runs_per_month", None) if equipment_idx == 0 else None)
                or runs_per_month_default
            ),
            key=f"{prefix}_runs_per_month",
        )
        st.number_input(
            "Fuel Used Per Hour",
            min_value=0.1,
            step=0.5,
            format="%.1f",
            value=float(
                (_legacy_site_value(i, j, "fuel_used_per_hour", None) if equipment_idx == 0 else None)
                or fuel_used_per_hour_default
            ),
            key=f"{prefix}_fuel_used_per_hour",
        )
        st.number_input(
            "Minimum Run Hours",
            min_value=0.25,
            step=0.25,
            format="%.2f",
            value=float(
                (_legacy_site_value(i, j, "run_hours_min", None) if equipment_idx == 0 else None)
                or run_hours_min_default
            ),
            key=f"{prefix}_run_hours_min",
        )
    with col2:
        st.radio(
            "Equipment Quantity Mode",
            options=["Tank Level Change", "Explicit Fuel Used"],
            horizontal=True,
            index=0
            if (
                (_legacy_site_value(i, j, "quantity_mode_label", None) if equipment_idx == 0 else None)
                or st.session_state.get("stationary_log_quantity_mode_label", "Tank Level Change")
            ) == "Tank Level Change"
            else 1,
            key=f"{prefix}_quantity_mode_label",
        )
        st.number_input(
            "Tank Capacity",
            min_value=50.0,
            step=25.0,
            format="%.1f",
            value=float(
                (_legacy_site_value(i, j, "tank_capacity", None) if equipment_idx == 0 else None)
                or tank_capacity_default
            ),
            key=f"{prefix}_tank_capacity",
        )
        st.number_input(
            "Maximum Run Hours",
            min_value=0.25,
            step=0.25,
            format="%.2f",
            value=float(
                (_legacy_site_value(i, j, "run_hours_max", None) if equipment_idx == 0 else None)
                or run_hours_max_default
            ),
            key=f"{prefix}_run_hours_max",
        )


def _render_equipment_item_details(prefix: str, i: int, j: int, document_type: str | None, equipment_idx: int) -> None:
    if document_type == "fuel_invoice":
        _render_invoice_item_fields(prefix, i, j, equipment_idx)
    elif document_type == "delivery_note":
        _render_delivery_note_item_fields(prefix, i, j, equipment_idx)
    elif document_type == "fuel_card":
        _render_fuel_card_item_fields(prefix, i, j, equipment_idx)
    elif document_type == "generator_log":
        _render_generator_log_item_fields(prefix, i, j, equipment_idx)


def _fuel_card_number_default(i: int, j: int) -> str:
    return f"****{8200 + ((i * 7 + j) % 80):04d}"


def _optional_field(widget_fn, label: str, key: str, *, omit_default: bool = False, help: str | None = None, **kwargs) -> None:
    is_omitted = bool(st.session_state.get(f"{key}_omit", omit_default))
    field_col, omit_col = st.columns([8, 1])
    with field_col:
        widget_fn(label, key=key, disabled=is_omitted, help=help, **kwargs)
    with omit_col:
        st.checkbox(
            "Omit",
            value=omit_default,
            key=f"{key}_omit",
            help="Leave this field blank in the generated output.",
        )


def _render_equipment_items(i: int, j: int, document_type: str | None) -> None:
    if document_type == "bems":
        return

    st.markdown("**Equipment Items**")
    omit_key = f"stationary_site_{i}_{j}_equipment_omit"
    st.session_state[omit_key] = False
    count_key = f"stationary_site_{i}_{j}_equipment_count"

    count = st.number_input(
        "Number of equipment items",
        min_value=1,
        max_value=20,
        value=1,
        step=1,
        key=count_key,
    )

    has_emission_source = document_type in {"fuel_invoice", "fuel_card", "generator_log"}
    for equipment_idx in range(int(count)):
        if equipment_idx > 0:
            st.divider()
        equipment_key = _item_key(i, j, equipment_idx)
        legacy_equipment = st.session_state.get(f"stationary_site_{i}_{j}_equipment", "") if equipment_idx == 0 else ""
        default_equipment = legacy_equipment or _equipment_default(i, j, document_type, equipment_idx)
        equipment_label = _equipment_label(document_type)
        if int(count) > 1:
            equipment_label = f"{equipment_label} {equipment_idx + 1}"

        if not has_emission_source:
            if document_type in {"delivery_note", "fuel_invoice", "fuel_card"}:
                _optional_field(
                    st.text_input,
                    equipment_label,
                    equipment_key,
                    value=default_equipment,
                    help=_equipment_help(document_type),
                )
            else:
                st.text_input(
                    equipment_label,
                    value=default_equipment,
                    key=equipment_key,
                    help=_equipment_help(document_type),
                )
            _render_equipment_item_details(equipment_key, i, j, document_type, equipment_idx)
            continue

        equipment_col, source_col = st.columns(2)
        with equipment_col:
            if document_type in {"delivery_note", "fuel_invoice", "fuel_card"}:
                _optional_field(
                    st.text_input,
                    equipment_label,
                    equipment_key,
                    value=default_equipment,
                    help=_equipment_help(document_type),
                )
            else:
                st.text_input(
                    equipment_label,
                    value=default_equipment,
                    key=equipment_key,
                    help=_equipment_help(document_type),
                )
        with source_col:
            legacy_source = st.session_state.get(f"stationary_site_{i}_{j}_emission_source", "") if equipment_idx == 0 else ""
            source_key = f"{equipment_key}_emission_source"
            source_label = "Emission Source" if int(count) == 1 else f"Emission Source {equipment_idx + 1}"
            _optional_field(
                st.text_input,
                source_label,
                source_key,
                value=legacy_source or _emission_source_default(i, j, equipment_idx),
                omit_default=(document_type == "fuel_card"),
                help=(
                    "Usually inferred rather than explicit in fuel-card statements."
                    if document_type == "fuel_card"
                    else "Often derived during mapping rather than shown on the source document."
                ),
            )
        _render_equipment_item_details(equipment_key, i, j, document_type, equipment_idx)


def _render_companies(document_type: str | None) -> None:
    st.markdown("#### Companies")
    count = st.number_input(
        "Number of companies",
        min_value=1,
        max_value=10,
        value=1,
        step=1,
        key="stationary_n_companies",
    )
    for i in range(int(count)):
        company_label = st.session_state.get(f"stationary_co_{i}_label") or _company_default(i, "label", document_type=document_type) or f"Company {i + 1}"
        with st.expander(f"Company {i + 1}: {company_label}", expanded=(i == 0)):
            _render_company(i, document_type)


def _render_company(i: int, document_type: str | None) -> None:
    col1, col2 = st.columns(2)
    with col1:
        company_label_field = "Account Name" if document_type == "fuel_card" else "Company Label"
        supplier_label = "Fuel Card Provider" if document_type == "fuel_card" else "Supplier Name"
        supplier_code_label = "Provider Code" if document_type == "fuel_card" else "Supplier Code"
        supplier_address_label = "Provider Address" if document_type == "fuel_card" else "Supplier Address"
        st.text_input(company_label_field, value=_company_default(i, "label", document_type=document_type), key=f"stationary_co_{i}_label")
        st.text_input(supplier_label, value=_company_default(i, "supplier", document_type=document_type), key=f"stationary_co_{i}_supplier")
        st.text_input(supplier_code_label, value=_company_default(i, "supplier_code", document_type=document_type), key=f"stationary_co_{i}_supplier_code")
        if document_type != "fuel_card":
            st.text_area(supplier_address_label, value=_company_default(i, "supplier_address", document_type=document_type), height=104, key=f"stationary_co_{i}_supplier_address")
        else:
            st.text_input(
                "Merchant",
                value=_FUEL_CARD_MERCHANTS[i % len(_FUEL_CARD_MERCHANTS)],
                key=f"stationary_co_{i}_merchant",
            )
    with col2:
        customer_label = "Account Holder" if document_type == "fuel_card" else "Bill To / Customer"
        st.text_input(customer_label, value=_company_default(i, "customer", document_type=document_type), key=f"stationary_co_{i}_customer")
        st.text_input("Customer Code", value=_company_default(i, "customer_code", document_type=document_type), key=f"stationary_co_{i}_customer_code")
        if document_type != "delivery_note":
            st.selectbox(
                "Currency",
                options=_currency_options(),
                index=_currency_index(
                    _company_default(
                        i,
                        "currency",
                        _CURRENCY_DISPLAY.get(NEW_COMPANY_PLACEHOLDER["currency"], "EUR (€)"),
                        document_type=document_type,
                    )
                ),
                key=f"stationary_co_{i}_currency",
            )
        if document_type == "fuel_card":
            st.text_input(
                "Card Number",
                value=_fuel_card_number_default(i, 0),
                key=f"stationary_co_{i}_card_number",
            )

    section_label = "Sites"
    count_label = "Number of sites"
    st.markdown(f"**{section_label} for Company {i + 1}**")
    site_count = st.number_input(
        count_label,
        min_value=1,
        max_value=20,
        value=1,
        step=1,
        key=f"stationary_n_sites_{i}",
    )
    for j in range(int(site_count)):
        default_site_label = _site_default(i, j, "label") or f"Site {j + 1}"
        site_label = st.session_state.get(f"stationary_site_{i}_{j}_label") or default_site_label
        expander_label = f"Site {j + 1}: {site_label or 'Unmapped'}"
        with st.expander(expander_label, expanded=(i == 0 and j == 0)):
            _render_site(i, j, document_type)


def _render_site(i: int, j: int, document_type: str | None) -> None:
    col1, col2 = st.columns(2)
    with col1:
        if document_type == "fuel_card":
            _optional_field(
                st.text_input,
                "Site",
                f"stationary_site_{i}_{j}_label",
                value=_site_default(i, j, "label", ""),
                omit_default=True,
                help="Fuel-card statements often do not include a mapped site.",
            )
        else:
            st.text_input("Site", value=_site_default(i, j, "label", ""), key=f"stationary_site_{i}_{j}_label")
        if document_type not in {"fuel_card"}:
            st.text_area("Address", value=_site_default(i, j, "address", ""), height=104, key=f"stationary_site_{i}_{j}_address")
        if document_type == "bems":
            _optional_field(
                st.text_input,
                "Country",
                f"stationary_site_{i}_{j}_country",
                value=_site_default(i, j, "country", ""),
                help="Often present in BEMS headers, but can be absent in some exports.",
            )
        elif document_type in {"fuel_invoice", "delivery_note", "generator_log"}:
            _optional_field(
                st.text_input,
                "Country",
                f"stationary_site_{i}_{j}_country",
                value=_site_default(i, j, "country", ""),
                help="Some source files identify the location through the address only.",
            )
        elif document_type == "fuel_card":
            _optional_field(
                st.text_input,
                "Country",
                f"stationary_site_{i}_{j}_country",
                value=_site_default(i, j, "country", ""),
                omit_default=True,
                help="Country is often not explicit on fuel-card statements.",
            )
        else:
            st.text_input("Country", value=_site_default(i, j, "country", ""), key=f"stationary_site_{i}_{j}_country")
    with col2:
        _render_equipment_items(i, j, document_type)

    if document_type == "bems":
        _render_bems_site_fields_module(i, j, _BEMS_ASSET_DEFAULTS, _bems_asset_default, _optional_field)


def _collect_site_equipment_items(i: int, j: int, document_type: str | None) -> list[dict]:
    s = st.session_state
    count = int(s.get(f"stationary_site_{i}_{j}_equipment_count", 1))
    has_emission_source = document_type in {"fuel_invoice", "fuel_card", "generator_log"}
    items: list[dict] = []

    for equipment_idx in range(max(count, 1)):
        equipment_key = _item_key(i, j, equipment_idx)
        legacy_equipment = s.get(f"stationary_site_{i}_{j}_equipment", "") if equipment_idx == 0 else ""
        item: dict = {
            "equipment": s.get(equipment_key, legacy_equipment),
            "_omit": {
                "equipment": bool(s.get(f"{equipment_key}_omit", False)),
            },
        }
        if has_emission_source:
            source_key = f"{equipment_key}_emission_source"
            legacy_source = s.get(f"stationary_site_{i}_{j}_emission_source", "") if equipment_idx == 0 else ""
            item["emission_source"] = s.get(source_key, legacy_source)
            item["_omit"]["emission_source"] = bool(
                s.get(f"{source_key}_omit", document_type == "fuel_card")
            )
        if document_type in {"fuel_invoice", "delivery_note", "fuel_card", "generator_log"}:
            item.update({
                "fuel": s.get(f"{equipment_key}_fuel", s.get(f"stationary_site_{i}_{j}_fuel", "")),
                "unit": s.get(f"{equipment_key}_unit", s.get(f"stationary_site_{i}_{j}_unit", default_fuel_volume_unit(document_type))),
            })
        if document_type in {"fuel_invoice", "delivery_note", "fuel_card"}:
            item["quantity"] = str(s.get(f"{equipment_key}_quantity", s.get(f"stationary_site_{i}_{j}_quantity", 0.0)))
        if document_type in {"fuel_invoice", "fuel_card"}:
            item["unit_price"] = str(s.get(f"{equipment_key}_unit_price", s.get(f"stationary_site_{i}_{j}_unit_price", 0.0)))
        if document_type == "fuel_invoice":
            item.update({
                "delivery_charge": str(s.get(f"{equipment_key}_delivery_charge", s.get(f"stationary_site_{i}_{j}_delivery_charge", 0.0))),
                "vat_rate": str(s.get(f"{equipment_key}_vat_rate", s.get(f"stationary_site_{i}_{j}_vat_rate", 20))),
            })
        elif document_type == "generator_log":
            item.update({
                "runs_per_month": int(s.get(f"{equipment_key}_runs_per_month", s.get(f"stationary_site_{i}_{j}_runs_per_month", 3))),
                "fuel_used_per_hour": str(s.get(f"{equipment_key}_fuel_used_per_hour", s.get(f"stationary_site_{i}_{j}_fuel_used_per_hour", 15.0))),
                "quantity_mode": "tank_level_change"
                if s.get(f"{equipment_key}_quantity_mode_label", s.get(f"stationary_site_{i}_{j}_quantity_mode_label", "Tank Level Change")) == "Tank Level Change"
                else "explicit_fuel_used",
                "tank_capacity": str(s.get(f"{equipment_key}_tank_capacity", s.get(f"stationary_site_{i}_{j}_tank_capacity", 800.0))),
                "run_hours_min": str(s.get(f"{equipment_key}_run_hours_min", s.get(f"stationary_site_{i}_{j}_run_hours_min", 0.5))),
                "run_hours_max": str(s.get(f"{equipment_key}_run_hours_max", s.get(f"stationary_site_{i}_{j}_run_hours_max", 5.0))),
            })
        items.append(item)

    return items


def _collect_companies(document_type: str | None) -> list[dict]:
    s = st.session_state
    companies: list[dict] = []
    company_count = int(s.get("stationary_n_companies", 1))

    for i in range(company_count):
        site_count = int(s.get(f"stationary_n_sites_{i}", 1))
        sites: list[dict] = []
        for j in range(site_count):
            equipment_items = _collect_site_equipment_items(i, j, document_type)
            primary_equipment = equipment_items[0].get("equipment", "") if equipment_items else ""
            primary_emission_source = equipment_items[0].get("emission_source", "") if equipment_items else ""
            primary_fuel = equipment_items[0].get("fuel", "") if equipment_items else ""
            primary_unit = equipment_items[0].get("unit", default_fuel_volume_unit(document_type)) if equipment_items else default_fuel_volume_unit(document_type)
            site = {
                "label": s.get(f"stationary_site_{i}_{j}_label", "") or ("" if document_type == "fuel_card" else f"Site {j + 1}"),
                "customer_address": [
                    line
                    for line in s.get(f"stationary_site_{i}_{j}_address", "").split("\n")
                    if line.strip()
                ],
                "country": s.get(f"stationary_site_{i}_{j}_country", ""),
                "equipment": primary_equipment,
                "equipment_items": equipment_items,
                "emission_source": primary_emission_source,
                "fuel": primary_fuel or s.get(f"stationary_site_{i}_{j}_fuel", ""),
                "unit": primary_unit,
            }
            if document_type == "fuel_invoice":
                site.update({
                    "quantity": str(equipment_items[0].get("quantity", s.get(f"stationary_site_{i}_{j}_quantity", 0.0))),
                    "unit_price": str(equipment_items[0].get("unit_price", s.get(f"stationary_site_{i}_{j}_unit_price", 0.0))),
                    "delivery_charge": str(equipment_items[0].get("delivery_charge", s.get(f"stationary_site_{i}_{j}_delivery_charge", 0.0))),
                    "vat_rate": str(equipment_items[0].get("vat_rate", s.get(f"stationary_site_{i}_{j}_vat_rate", 20))),
                    "_omit": {
                        "country": bool(s.get(f"stationary_site_{i}_{j}_country_omit", False)),
                        "equipment": False,
                        "emission_source": False,
                    },
                })
            elif document_type == "delivery_note":
                site.update({
                    "quantity": str(equipment_items[0].get("quantity", s.get(f"stationary_site_{i}_{j}_quantity", 0.0))),
                    "_omit": {
                        "country": bool(s.get(f"stationary_site_{i}_{j}_country_omit", False)),
                        "equipment": False,
                    },
                })
            elif document_type == "fuel_card":
                site.update({
                    "quantity": str(equipment_items[0].get("quantity", s.get(f"stationary_site_{i}_{j}_quantity", 0.0))),
                    "unit_price": str(equipment_items[0].get("unit_price", s.get(f"stationary_site_{i}_{j}_unit_price", 0.0))),
                    "_omit": {
                        "label": bool(s.get(f"stationary_site_{i}_{j}_label_omit", True)),
                        "country": bool(s.get(f"stationary_site_{i}_{j}_country_omit", True)),
                        "equipment": False,
                        "emission_source": False,
                    },
                })
            elif document_type == "bems":
                asset_count = int(s.get(f"stationary_site_{i}_{j}_asset_count", len(_BEMS_ASSET_DEFAULTS)))
                assets: list[dict] = []
                for asset_idx in range(asset_count):
                    assets.append({
                        "asset_tag": s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_tag",
                            _bems_asset_default(asset_idx, "asset_tag", f"AST-{asset_idx + 1:02d}"),
                        ),
                        "equipment_name": s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_equipment_name",
                            _bems_asset_default(asset_idx, "equipment_name", ""),
                        ),
                        "emission_source": s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_emission_source",
                            _bems_asset_default(asset_idx, "emission_source", ""),
                        ),
                        "fuel": s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_fuel",
                            _bems_asset_default(asset_idx, "fuel", ""),
                        ),
                        "unit": s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_unit",
                            _bems_asset_default(asset_idx, "unit", "kWh"),
                        ),
                        "sensor_name": s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_sensor_name",
                            _bems_asset_default(asset_idx, "sensor_name", ""),
                        ),
                        "quantity": str(s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_quantity",
                            _bems_asset_default(asset_idx, "quantity", 0.0),
                        )),
                        "operating_hours": str(s.get(
                            f"stationary_site_{i}_{j}_asset_{asset_idx}_operating_hours",
                            _bems_asset_default(asset_idx, "operating_hours", 0.0),
                        )),
                        "_omit": {
                            "equipment_name": bool(s.get(f"stationary_site_{i}_{j}_asset_{asset_idx}_equipment_name_omit", False)),
                            "emission_source": bool(s.get(f"stationary_site_{i}_{j}_asset_{asset_idx}_emission_source_omit", False)),
                            "sensor_name": bool(s.get(f"stationary_site_{i}_{j}_asset_{asset_idx}_sensor_name_omit", False)),
                            "fuel": bool(s.get(f"stationary_site_{i}_{j}_asset_{asset_idx}_fuel_omit", False)),
                            "operating_hours": bool(s.get(f"stationary_site_{i}_{j}_asset_{asset_idx}_operating_hours_omit", False)),
                        },
                    })
                site.update({
                    "assets": assets,
                    "_omit": {
                        "country": bool(s.get(f"stationary_site_{i}_{j}_country_omit", False)),
                    },
                })
            else:
                site.update({
                    "runs_per_month": int(equipment_items[0].get("runs_per_month", s.get(f"stationary_site_{i}_{j}_runs_per_month", 3))),
                    "fuel_used_per_hour": str(equipment_items[0].get("fuel_used_per_hour", s.get(f"stationary_site_{i}_{j}_fuel_used_per_hour", 15.0))),
                    "quantity_mode": equipment_items[0].get("quantity_mode", "tank_level_change"),
                    "tank_capacity": str(equipment_items[0].get("tank_capacity", s.get(f"stationary_site_{i}_{j}_tank_capacity", 800.0))),
                    "run_hours_min": str(equipment_items[0].get("run_hours_min", s.get(f"stationary_site_{i}_{j}_run_hours_min", 0.5))),
                    "run_hours_max": str(equipment_items[0].get("run_hours_max", s.get(f"stationary_site_{i}_{j}_run_hours_max", 5.0))),
                    "_omit": {
                        "country": bool(s.get(f"stationary_site_{i}_{j}_country_omit", False)),
                        "emission_source": False,
                    },
                })
            sites.append(site)

        companies.append({
            "label": s.get(f"stationary_co_{i}_label", "") or f"Company {i + 1}",
            "supplier": s.get(f"stationary_co_{i}_supplier", ""),
            "supplier_code": s.get(f"stationary_co_{i}_supplier_code", ""),
            "supplier_address": [
                line
                for line in s.get(f"stationary_co_{i}_supplier_address", "").split("\n")
                if line.strip()
            ] if document_type != "fuel_card" else [],
            "customer": s.get(f"stationary_co_{i}_customer", ""),
            "customer_code": s.get(f"stationary_co_{i}_customer_code", ""),
            "currency": s.get(f"stationary_co_{i}_currency", "GBP (£)"),
            "merchant": s.get(f"stationary_co_{i}_merchant", ""),
            "card_number": s.get(f"stationary_co_{i}_card_number", ""),
            "sites": sites,
            "_omit": {},
        })

    return companies


def render_stationary_combustion_form(document_type: str | None) -> dict:
    st.subheader("Stationary Combustion")
    captions = {
        "fuel_invoice": "Fuel supplier invoice configuration for stationary combustion.",
        "generator_log": "Generator operation log configuration for stationary combustion.",
        "bems": "Building Energy Management System configuration for stationary combustion.",
        "delivery_note": "Fuel delivery note configuration for stationary combustion.",
        "fuel_card": "Fuel card statement configuration for stationary combustion.",
    }
    st.caption(captions.get(document_type, "Stationary combustion document configuration."))

    if document_type == "bems":
        _render_bems_report_type_selector_module(_BEMS_REPORT_TYPES)
    _render_document_settings(document_type)
    _render_financial_period()

    if document_type == "fuel_invoice":
        _render_invoice_defaults_module(_FUELS)
    elif document_type == "delivery_note":
        _render_delivery_note_defaults_module(_FUELS)
    elif document_type == "fuel_card":
        _render_fuel_card_defaults_module(_FUELS)
    elif document_type == "bems":
        _render_bems_defaults_module(_bems_report_type_value_module(_BEMS_REPORT_TYPES), _BEMS_ASSET_DEFAULTS)
    else:
        _render_log_defaults_module(_FUELS)

    _render_companies(document_type)

    s = st.session_state
    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end: date = s.get("fp_end", date(2026, 1, 31))
    default_title, default_subject = _document_defaults(document_type)

    return {
        "_category": "stationary_combustion",
        "document_type": document_type or "fuel_invoice",
        "doc_title": s.get("doc_title", default_title),
        "doc_subject": s.get("doc_subject", default_subject),
        "doc_seed": int(s.get("doc_seed", 20260415)),
        "fp_label": s.get("fp_label", "January 2026"),
        "fp_start": fp_start.isoformat(),
        "fp_end": fp_end.isoformat(),
        "doc_language": _LANGUAGE_OPTIONS.get(s.get("doc_language_label", "English"), "en"),
        "doc_noise": float(s.get("doc_noise", 0.0)),
        "doc_inject_special_chars": bool(s.get("doc_inject_special_chars", False)),
        "bems_interval_minutes": int(str(s.get("stationary_bems_interval_label", "60 minutes")).split()[0]),
        "bems_report_type": _bems_report_type_value_module(_BEMS_REPORT_TYPES),
        "companies": _collect_companies(document_type),
    }
