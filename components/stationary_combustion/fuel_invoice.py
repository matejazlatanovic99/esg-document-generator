from __future__ import annotations

import random

import streamlit as st

from components.stationary_combustion.units import (
    DEFAULT_LONG_FUEL_VOLUME_UNIT,
    LONG_FUEL_VOLUME_UNITS,
    option_index,
)
from utils.currency import currency_code


def _default_currency_code() -> str:
    return currency_code(st.session_state.get("stationary_co_0_currency", "GBP (£)"))


def init_invoice_defaults(fuels: list[str]) -> None:
    st.session_state.setdefault("stationary_invoice_fuel_randomize", True)
    st.session_state.setdefault("stationary_invoice_unit_randomize", True)
    st.session_state.setdefault("stationary_invoice_quantity_randomize", True)
    st.session_state.setdefault("stationary_invoice_unit_price_randomize", True)
    st.session_state.setdefault("stationary_invoice_delivery_charge_randomize", True)
    if "stationary_invoice_quantity" in st.session_state:
        return
    rng = random.Random(20260415)
    st.session_state["stationary_invoice_fuel"] = rng.choice(fuels)
    st.session_state["stationary_invoice_unit"] = DEFAULT_LONG_FUEL_VOLUME_UNIT
    st.session_state["stationary_invoice_quantity"] = float(rng.randrange(1200, 6000, 50))
    st.session_state["stationary_invoice_unit_price"] = round(rng.uniform(0.88, 1.35), 2)
    st.session_state["stationary_invoice_delivery_charge"] = round(rng.uniform(20.0, 95.0), 2)
    st.session_state["stationary_invoice_vat_rate"] = 20


def render_invoice_defaults(fuels: list[str]) -> None:
    init_invoice_defaults(fuels)
    st.markdown("#### Fuel Invoice Defaults")
    st.caption("These defaults apply to every site unless you override them.")
    col1, col2 = st.columns(2)
    with col1:
        randomize_fuel = st.checkbox(
            "Randomize fuel",
            key="stationary_invoice_fuel_randomize",
            help="New equipment entries pick an initial fuel from the Default Fuel dropdown options.",
        )
        st.selectbox("Default Fuel", options=fuels, key="stationary_invoice_fuel", disabled=randomize_fuel)
        randomize_unit = st.checkbox(
            "Randomize unit",
            key="stationary_invoice_unit_randomize",
            help="New equipment entries pick an initial unit from the invoice unit options.",
        )
        st.selectbox("Default Unit", options=LONG_FUEL_VOLUME_UNITS, key="stationary_invoice_unit", disabled=randomize_unit)
    with col2:
        randomize_quantity = st.checkbox(
            "Randomize quantity",
            key="stationary_invoice_quantity_randomize",
            help="New equipment entries get an initial quantity in the invoice quantity range.",
        )
        st.number_input(
            "Default Quantity",
            min_value=0.0,
            step=50.0,
            format="%.2f",
            key="stationary_invoice_quantity",
            disabled=randomize_quantity,
        )
        randomize_unit_price = st.checkbox(
            "Randomize unit price",
            key="stationary_invoice_unit_price_randomize",
            help="New equipment entries get an initial unit price in the invoice price range.",
        )
        st.number_input(
            f"Default Unit Price ({_default_currency_code()})",
            min_value=0.01,
            step=0.01,
            format="%.2f",
            key="stationary_invoice_unit_price",
            disabled=randomize_unit_price,
        )
    col3, col4 = st.columns(2)
    with col3:
        randomize_delivery_charge = st.checkbox(
            "Randomize delivery charge",
            key="stationary_invoice_delivery_charge_randomize",
            help="New equipment entries get an initial delivery charge in the invoice delivery range.",
        )
        st.number_input(
            f"Default Delivery Charge ({_default_currency_code()})",
            min_value=0.0,
            step=5.0,
            format="%.2f",
            key="stationary_invoice_delivery_charge",
            disabled=randomize_delivery_charge,
        )
    with col4:
        st.number_input("VAT Rate (%)", min_value=0, max_value=100, step=1, key="stationary_invoice_vat_rate")
    st.divider()


def render_invoice_site_fields(i: int, j: int, site_default, fuels: list[str]) -> None:
    st.markdown("**Invoice Line Defaults**")
    col1, col2 = st.columns(2)
    with col1:
        st.text_input(
            "Fuel",
            value=site_default(i, j, "fuel", st.session_state.get("stationary_invoice_fuel", fuels[0])),
            key=f"stationary_site_{i}_{j}_fuel",
        )
        default_unit = st.session_state.get("stationary_invoice_unit", DEFAULT_LONG_FUEL_VOLUME_UNIT)
        st.selectbox(
            "Unit",
            options=LONG_FUEL_VOLUME_UNITS,
            index=option_index(LONG_FUEL_VOLUME_UNITS, default_unit),
            key=f"stationary_site_{i}_{j}_unit",
        )
        st.number_input(
            "Quantity",
            min_value=0.0,
            step=50.0,
            format="%.2f",
            value=float(st.session_state.get("stationary_invoice_quantity", 2500.0)),
            key=f"stationary_site_{i}_{j}_quantity",
        )
    with col2:
        st.number_input(
            f"Unit Price ({_default_currency_code()})",
            min_value=0.01,
            step=0.01,
            format="%.2f",
            value=float(st.session_state.get("stationary_invoice_unit_price", 1.12)),
            key=f"stationary_site_{i}_{j}_unit_price",
        )
        st.number_input(
            f"Delivery Charge ({_default_currency_code()})",
            min_value=0.0,
            step=5.0,
            format="%.2f",
            value=float(st.session_state.get("stationary_invoice_delivery_charge", 50.0)),
            key=f"stationary_site_{i}_{j}_delivery_charge",
        )
        st.number_input(
            "VAT Rate (%)",
            min_value=0,
            max_value=100,
            step=1,
            value=int(st.session_state.get("stationary_invoice_vat_rate", 20)),
            key=f"stationary_site_{i}_{j}_vat_rate",
        )
