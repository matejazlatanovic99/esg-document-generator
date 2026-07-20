from __future__ import annotations

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
from utils.currency import (
    currency_code as _currency_code,
    currency_index as _currency_index,
    currency_options as _currency_options,
)
from utils.document_catalog import document_type_requires_company_currency

_LANGUAGE_OPTIONS: dict[str, str] = {
    "English": "en",
}

_FUEL_SUPPLIERS = [
    {
        "name": "Roadside Fuel Supplies Ltd",
        "code": "RFS",
        "address": "Fleet Fuel Terminal\n4 Carriage Way\nLeeds LS10 1DN\nUnited Kingdom",
    },
    {
        "name": "Shamrock Forecourts Ltd",
        "code": "SHF",
        "address": "Forecourt Services Park\n12 Naas Road\nDublin D12 XK52\nIreland",
    },
    {
        "name": "Autobahn Kraftstoff GmbH",
        "code": "ABK",
        "address": "Tankstellenring 8\n60486 Frankfurt am Main\nGermany",
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

_TELEMATICS_PROVIDERS = [
    {
        "name": "Geotab Fleet Analytics",
        "code": "GEO",
        "address": "Fleet Data Centre\n21 Analytics Way\nDublin\nIreland",
    },
    {
        "name": "Webfleet Solutions",
        "code": "WFS",
        "address": "Telematics House\n14 Signal Street\nAmsterdam\nNetherlands",
    },
    {
        "name": "Samsara Connected Operations",
        "code": "SAM",
        "address": "Connected Fleet Hub\n5 Motorway Park\nLondon\nUnited Kingdom",
    },
]

_VEHICLE_FUELS = [
    "Diesel",
    "Petrol",
    "Diesel (average biofuel blend)",
    "Petrol (average biofuel blend)",
    "CNG",
    "LPG",
    "Biodiesel HVO",
]

_VEHICLE_TYPES = [
    "Passenger car",
    "Van (LGV)",
    "HGV",
    "Motorcycle",
    "Forklift",
    "Other",
]

_VEHICLE_MODELS = [
    "Ford Transit 350 L3",
    "Toyota Corolla Hybrid Estate",
    "Mercedes-Benz Sprinter 315",
    "Volkswagen Caddy Cargo",
    "Toyota Hilux Active",
    "DAF LF 180 7.5t",
    "Toyota Proace Verso",
    "Skoda Octavia Estate",
]

_DRIVERS = [
    "J. Smith",
    "A. Murphy",
    "P. Novak",
    "L. Fischer",
    "M. Horvath",
    "S. Jensen",
    "K. O'Brien",
    "T. Walker",
]

_MERCHANTS = [
    "Applegreen M1 Services",
    "Shell Recharge Kings Cross",
    "Circle K Ballymount",
    "BP Connect Hammersmith",
    "Esso Severn View",
    "Texaco City North",
]

_DEPOTS = [
    "London Depot",
    "Manchester DC",
    "Dublin Yard",
    "Birmingham Hub",
]

_FUEL_UNITS = ("L", "gal")
_DISTANCE_UNITS = ("km", "mi")

_REG_PREFIXES = ["AB", "BD", "CE", "DK", "EF", "GH", "KL", "MN"]
_REG_SUFFIXES = ["CDE", "FGH", "JKL", "MNP", "RST", "VWX", "XYZ", "QRS"]


def _providers_for(document_type: str | None) -> list[dict]:
    if document_type == "fuel_card_statement":
        return _FUEL_CARD_PROVIDERS
    if document_type in {"telematics_fuel", "telematics_trips", "telematics_odometer"}:
        return _TELEMATICS_PROVIDERS
    return _FUEL_SUPPLIERS


def _document_defaults(document_type: str | None) -> tuple[str, str]:
    cfg = get_document_type_config("mobile_combustion", document_type or "")
    return cfg.get("default_title", "Document"), cfg.get("default_subject", "")


def _sync_document_setting_defaults(document_type: str | None) -> None:
    selection_key = f"mobile_combustion:{document_type or ''}"
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

        if document_type in {"fuel_invoice", "fuel_card_statement"}:
            st.checkbox(
                "Include line items from other scopes/categories",
                key="doc_cross_scope_items",
                help=(
                    "Mix stationary-combustion fuel lines (e.g. heating oil for a boiler) into the "
                    "document so extraction has to separate categories. Ground-truth entries are "
                    "tagged with their true category."
                ),
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

        _render_distractor_field_controls()
        _render_layout_controls()
        st.divider()
        _render_invalid_data_controls()


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


def _init_defaults() -> None:
    s = st.session_state
    s.setdefault("mobile_n_companies", 1)
    s.setdefault("mobile_invoice_doc_count", 1)
    s.setdefault("mobile_fuel_card_doc_count", 12)
    s.setdefault("mobile_vat_rate", 20)
    s.setdefault("mobile_trips_per_vehicle", 10)
    s.setdefault("mobile_distance_unit", "km")


def _render_defaults(document_type: str | None) -> None:
    st.markdown("#### Generation Defaults")

    if document_type == "fuel_invoice":
        col1, col2 = st.columns(2)
        with col1:
            st.number_input(
                "Number of invoices to generate",
                min_value=1,
                max_value=200,
                step=1,
                key="mobile_invoice_doc_count",
                help=(
                    "Each invoice is a separate one-time fuel purchase for one vehicle. Generate 1 "
                    "for a single file, or more to download them all as a ZIP archive."
                ),
            )
        with col2:
            st.number_input("VAT Rate (%)", min_value=0, max_value=100, step=1, key="mobile_vat_rate")
    elif document_type == "fuel_card_statement":
        col1, col2 = st.columns(2)
        with col1:
            st.number_input(
                "Number of line items (transactions)",
                min_value=1,
                max_value=500,
                step=1,
                key="mobile_fuel_card_doc_count",
                help=(
                    "How many transaction lines appear on each statement. The configured vehicles "
                    "seed the lines; extra lines reuse them with randomized dates, quantities, and prices."
                ),
            )
        with col2:
            st.number_input("VAT Rate (%)", min_value=0, max_value=100, step=1, key="mobile_vat_rate")
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.selectbox("Distance Unit", options=_DISTANCE_UNITS, key="mobile_distance_unit")
        with col2:
            if document_type == "telematics_trips":
                st.number_input(
                    "Trips per vehicle",
                    min_value=1,
                    max_value=200,
                    step=1,
                    key="mobile_trips_per_vehicle",
                    help="Number of trip rows generated per vehicle across the reporting period.",
                )
    st.divider()


def _vehicle_default(i: int, j: int, field: str):
    rng = random.Random(f"mobile_vehicle_default:{field}:{i}:{j}")
    if field == "registration":
        prefix = _REG_PREFIXES[(i * 3 + j) % len(_REG_PREFIXES)]
        suffix = _REG_SUFFIXES[(i + j * 2) % len(_REG_SUFFIXES)]
        return f"{prefix}{rng.randint(10, 72)} {suffix}"
    if field == "make_model":
        return _VEHICLE_MODELS[(i * 2 + j) % len(_VEHICLE_MODELS)]
    if field == "driver":
        return _DRIVERS[(i * 3 + j) % len(_DRIVERS)]
    if field == "vehicle_type":
        return "Van (LGV)" if (i + j) % 2 == 0 else "Passenger car"
    if field == "fuel":
        return _VEHICLE_FUELS[0] if (i + j) % 3 else _VEHICLE_FUELS[1]
    if field == "card_number":
        return f"7002 34{rng.randint(10, 99)} XXXX {rng.randint(1000, 9999)}"
    if field == "site":
        return _DEPOTS[(i + j) % len(_DEPOTS)]
    if field == "quantity":
        return float(rng.randrange(35, 80))
    if field == "unit_price":
        return round(rng.uniform(1.38, 1.72), 2)
    if field == "monthly_distance_km":
        return float(rng.randrange(1400, 4200, 100))
    if field == "efficiency_l_per_100km":
        return round(rng.uniform(6.2, 13.5), 1)
    if field == "odometer_start":
        return float(rng.randrange(18000, 140000, 500))
    return ""


def _option_index(options: tuple[str, ...] | list[str], selected) -> int:
    return options.index(selected) if selected in options else 0


def _render_vehicle(i: int, j: int, document_type: str | None) -> None:
    key = f"mobile_v_{i}_{j}"
    col1, col2 = st.columns(2)
    with col1:
        st.text_input("Registration", value=_vehicle_default(i, j, "registration"), key=f"{key}_registration")
        st.text_input("Make / Model", value=_vehicle_default(i, j, "make_model"), key=f"{key}_make_model")
        st.selectbox(
            "Vehicle Type",
            options=_VEHICLE_TYPES,
            index=_option_index(_VEHICLE_TYPES, _vehicle_default(i, j, "vehicle_type")),
            key=f"{key}_vehicle_type",
        )
        st.selectbox(
            "Fuel",
            options=_VEHICLE_FUELS,
            index=_option_index(_VEHICLE_FUELS, _vehicle_default(i, j, "fuel")),
            key=f"{key}_fuel",
        )
        omit_driver = st.checkbox("Omit driver", key=f"{key}_driver_omit")
        st.text_input("Driver", value=_vehicle_default(i, j, "driver"), key=f"{key}_driver", disabled=omit_driver)
    with col2:
        if document_type in {"fuel_invoice", "fuel_card_statement"}:
            st.selectbox("Fuel Unit", options=_FUEL_UNITS, key=f"{key}_unit")
            st.number_input(
                "Typical Fill Quantity",
                min_value=0.0,
                step=5.0,
                format="%.2f",
                value=_vehicle_default(i, j, "quantity"),
                key=f"{key}_quantity",
            )
            st.number_input(
                f"Unit Price ({_currency_code(st.session_state.get(f'mobile_co_{i}_currency', 'GBP (£)'))})",
                min_value=0.01,
                step=0.01,
                format="%.2f",
                value=_vehicle_default(i, j, "unit_price"),
                key=f"{key}_unit_price",
            )
        if document_type == "fuel_card_statement":
            st.text_input("Card Number", value=_vehicle_default(i, j, "card_number"), key=f"{key}_card_number")
        if document_type in {"telematics_fuel", "telematics_trips", "telematics_odometer"}:
            st.number_input(
                "Monthly Distance (km)",
                min_value=0.0,
                step=100.0,
                format="%.0f",
                value=_vehicle_default(i, j, "monthly_distance_km"),
                key=f"{key}_monthly_distance_km",
            )
        if document_type == "telematics_fuel":
            st.number_input(
                "Fuel Efficiency (L/100km)",
                min_value=0.0,
                step=0.1,
                format="%.1f",
                value=_vehicle_default(i, j, "efficiency_l_per_100km"),
                key=f"{key}_efficiency_l_per_100km",
            )
        if document_type in {"fuel_card_statement", "telematics_odometer"}:
            st.number_input(
                "Odometer Start (km)",
                min_value=0.0,
                step=500.0,
                format="%.0f",
                value=_vehicle_default(i, j, "odometer_start"),
                key=f"{key}_odometer_start",
            )
        st.text_input("Home Site / Depot", value=_vehicle_default(i, j, "site"), key=f"{key}_site")


def _render_companies(document_type: str | None) -> None:
    st.markdown("#### Companies & Fleet")
    st.number_input("Number of companies", min_value=1, max_value=10, step=1, key="mobile_n_companies")

    providers = _providers_for(document_type)
    provider_label = {
        "fuel_invoice": "Supplier Name",
        "fuel_card_statement": "Fuel Card Provider",
    }.get(document_type or "", "Telematics Provider")

    for i in range(int(st.session_state.get("mobile_n_companies", 1))):
        provider = providers[i % len(providers)]
        with st.expander(f"Company {i + 1}", expanded=i == 0):
            col1, col2 = st.columns(2)
            with col1:
                st.text_input("Company Label", value=f"Fleet Operations {i + 1}", key=f"mobile_co_{i}_label")
                st.text_input(provider_label, value=provider["name"], key=f"mobile_co_{i}_supplier")
                if document_type == "fuel_invoice":
                    st.text_input("Supplier Code", value=provider["code"], key=f"mobile_co_{i}_supplier_code")
                    st.text_area(
                        "Supplier Address",
                        value=provider["address"],
                        key=f"mobile_co_{i}_supplier_address",
                        height=110,
                    )
            with col2:
                st.text_input("Customer Name", value="Toyota Financial Services UK", key=f"mobile_co_{i}_customer")
                st.text_input("Customer Code", value=f"TFS{i + 1}", key=f"mobile_co_{i}_customer_code")
                st.text_input(
                    "Account Number",
                    value="",
                    key=f"mobile_co_{i}_account_number",
                    help="Optional; generated from the customer code when left blank.",
                )
                if document_type_requires_company_currency("mobile_combustion", document_type or ""):
                    st.selectbox(
                        "Currency",
                        options=_currency_options(),
                        index=_currency_index(st.session_state.get(f"mobile_co_{i}_currency", "GBP (£)")),
                        key=f"mobile_co_{i}_currency",
                    )
                if document_type == "fuel_card_statement":
                    st.text_input(
                        "Merchants (one per line)",
                        value="; ".join(_MERCHANTS[:3]),
                        key=f"mobile_co_{i}_merchants",
                        help="Semicolon-separated list of merchants/stations used for transactions.",
                    )

            st.session_state.setdefault(f"mobile_n_vehicles_{i}", 2)
            st.number_input(
                "Number of vehicles",
                min_value=1,
                max_value=25,
                step=1,
                key=f"mobile_n_vehicles_{i}",
            )
            for j in range(int(st.session_state.get(f"mobile_n_vehicles_{i}", 2))):
                st.markdown(f"**Vehicle {j + 1}**")
                _render_vehicle(i, j, document_type)
                st.divider()


def _collect_companies(document_type: str | None) -> list[dict]:
    s = st.session_state
    companies: list[dict] = []

    for i in range(int(s.get("mobile_n_companies", 1))):
        vehicles: list[dict] = []
        for j in range(int(s.get(f"mobile_n_vehicles_{i}", 2))):
            key = f"mobile_v_{i}_{j}"
            vehicles.append({
                "registration": s.get(f"{key}_registration", ""),
                "make_model": s.get(f"{key}_make_model", ""),
                "vehicle_type": s.get(f"{key}_vehicle_type", _VEHICLE_TYPES[0]),
                "fuel": s.get(f"{key}_fuel", _VEHICLE_FUELS[0]),
                "unit": s.get(f"{key}_unit", "L"),
                "driver": s.get(f"{key}_driver", ""),
                "card_number": s.get(f"{key}_card_number", ""),
                "site": s.get(f"{key}_site", ""),
                "quantity": str(s.get(f"{key}_quantity", _vehicle_default(i, j, "quantity"))),
                "unit_price": str(s.get(f"{key}_unit_price", _vehicle_default(i, j, "unit_price"))),
                "monthly_distance_km": str(s.get(f"{key}_monthly_distance_km", _vehicle_default(i, j, "monthly_distance_km"))),
                "efficiency_l_per_100km": str(s.get(f"{key}_efficiency_l_per_100km", _vehicle_default(i, j, "efficiency_l_per_100km"))),
                "odometer_start": str(s.get(f"{key}_odometer_start", _vehicle_default(i, j, "odometer_start"))),
                "_omit": {
                    "driver": bool(s.get(f"{key}_driver_omit", False)),
                },
            })

        merchants = [
            merchant.strip()
            for merchant in str(s.get(f"mobile_co_{i}_merchants", "")).split(";")
            if merchant.strip()
        ]
        companies.append({
            "label": s.get(f"mobile_co_{i}_label", "") or f"Company {i + 1}",
            "supplier": s.get(f"mobile_co_{i}_supplier", ""),
            "supplier_code": s.get(f"mobile_co_{i}_supplier_code", ""),
            "supplier_address": [
                line
                for line in s.get(f"mobile_co_{i}_supplier_address", "").split("\n")
                if line.strip()
            ],
            "customer": s.get(f"mobile_co_{i}_customer", ""),
            "customer_code": s.get(f"mobile_co_{i}_customer_code", ""),
            "account_number": s.get(f"mobile_co_{i}_account_number", ""),
            "currency": (
                s.get(f"mobile_co_{i}_currency", "GBP (£)")
                if document_type_requires_company_currency("mobile_combustion", document_type or "")
                else ""
            ),
            "merchants": merchants,
            "merchant": merchants[0] if merchants else "",
            "vehicles": vehicles,
            "_omit": {},
        })

    return companies


def render_mobile_combustion_form(document_type: str | None) -> dict:
    st.subheader("Mobile Combustion")
    captions = {
        "fuel_invoice": "Single fuel purchase invoice / receipt for a fleet vehicle.",
        "fuel_card_statement": "Monthly fuel card statement with per-transaction detail.",
        "telematics_fuel": "Telematics fuel-usage report — the preferred fuel-based source.",
        "telematics_trips": "Trip-level telematics export — distance-based fallback.",
        "telematics_odometer": "Odometer / mileage summary — distance-based fallback.",
    }
    st.caption(captions.get(document_type, "Mobile combustion document configuration."))

    _init_defaults()
    _render_document_settings(document_type)
    _render_financial_period()
    _render_defaults(document_type)
    _render_companies(document_type)

    s = st.session_state
    fp_start: date = s.get("fp_start", date(2026, 1, 1))
    fp_end: date = s.get("fp_end", date(2026, 1, 31))
    default_title, default_subject = _document_defaults(document_type)

    doc_count_keys = {
        "fuel_invoice": "mobile_invoice_doc_count",
        "fuel_card_statement": "mobile_fuel_card_doc_count",
    }
    doc_count_key = doc_count_keys.get(document_type)
    doc_count = int(s.get(doc_count_key, 1) or 1) if doc_count_key else 1

    return {
        "_category": "mobile_combustion",
        "document_type": document_type or "fuel_invoice",
        "doc_count": doc_count,
        "mobile_vat_rate": int(s.get("mobile_vat_rate", 20) or 0),
        "mobile_distance_unit": s.get("mobile_distance_unit", "km"),
        "mobile_trips_per_vehicle": int(s.get("mobile_trips_per_vehicle", 10) or 10),
        "doc_cross_scope_items": bool(s.get("doc_cross_scope_items", False)),
        "doc_title": s.get("doc_title", default_title),
        "doc_subject": s.get("doc_subject", default_subject),
        "doc_seed": int(s.get("doc_seed", 20260415)),
        "fp_label": s.get("fp_label", "January 2026"),
        "fp_start": fp_start.isoformat(),
        "fp_end": fp_end.isoformat(),
        "doc_language": _LANGUAGE_OPTIONS.get(s.get("doc_language_label", "English"), "en"),
        "doc_noise": float(s.get("doc_noise", 0.0)),
        "doc_inject_special_chars": bool(s.get("doc_inject_special_chars", False)),
        "companies": _collect_companies(document_type),
        **_collect_distractor_fields(s),
        **_collect_layout(s),
        **_collect_invalid_data(s),
    }
