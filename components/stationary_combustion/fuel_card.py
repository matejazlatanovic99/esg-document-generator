from __future__ import annotations

import random

import streamlit as st

from components.stationary_combustion.units import (
    COMPACT_FUEL_VOLUME_UNITS,
    DEFAULT_COMPACT_FUEL_VOLUME_UNIT,
    option_index,
)


def init_fuel_card_defaults(fuels: list[str]) -> None:
    st.session_state.setdefault("stationary_fuel_card_fuel_randomize", True)
    st.session_state.setdefault("stationary_fuel_card_unit_randomize", True)
    st.session_state.setdefault("stationary_fuel_card_quantity_randomize", True)
    st.session_state.setdefault("stationary_fuel_card_unit_price_randomize", True)
    if "stationary_fuel_card_quantity" in st.session_state:
        return
    rng = random.Random(20260418)
    st.session_state["stationary_fuel_card_fuel"] = rng.choice(fuels)
    st.session_state["stationary_fuel_card_unit"] = DEFAULT_COMPACT_FUEL_VOLUME_UNIT
    st.session_state["stationary_fuel_card_quantity"] = float(rng.randrange(150, 700, 10))
    st.session_state["stationary_fuel_card_unit_price"] = round(rng.uniform(0.95, 1.45), 2)


def render_fuel_card_defaults(fuels: list[str]) -> None:
    init_fuel_card_defaults(fuels)
    st.markdown("#### Fuel Card Defaults")
    st.caption("These defaults apply to each transaction line unless you override them.")
    col1, col2 = st.columns(2)
    with col1:
        randomize_fuel = st.checkbox(
            "Randomize fuel",
            key="stationary_fuel_card_fuel_randomize",
            help="New equipment entries pick an initial fuel from the Default Fuel dropdown options.",
        )
        st.selectbox("Default Fuel", options=fuels, key="stationary_fuel_card_fuel", disabled=randomize_fuel)
        randomize_unit = st.checkbox(
            "Randomize unit",
            key="stationary_fuel_card_unit_randomize",
            help="New equipment entries pick an initial unit from the fuel-card unit options.",
        )
        st.selectbox("Default Unit", options=COMPACT_FUEL_VOLUME_UNITS, key="stationary_fuel_card_unit", disabled=randomize_unit)
    with col2:
        randomize_quantity = st.checkbox(
            "Randomize quantity",
            key="stationary_fuel_card_quantity_randomize",
            help="New equipment entries get an initial quantity in the fuel-card quantity range.",
        )
        st.number_input(
            "Default Quantity",
            min_value=0.0,
            step=10.0,
            format="%.2f",
            key="stationary_fuel_card_quantity",
            disabled=randomize_quantity,
        )
        randomize_unit_price = st.checkbox(
            "Randomize unit price",
            key="stationary_fuel_card_unit_price_randomize",
            help="New equipment entries get an initial unit price in the fuel-card price range.",
        )
        st.number_input(
            "Default Unit Price",
            min_value=0.01,
            step=0.01,
            format="%.2f",
            key="stationary_fuel_card_unit_price",
            disabled=randomize_unit_price,
        )
    st.caption("Site, country, and emission source are omitted by default because fuel-card statements often do not provide them clearly.")
    st.divider()


def render_fuel_card_site_fields(i: int, j: int, site_default, fuels: list[str], merchants: list[str], fuel_card_number_default) -> None:
    st.markdown("**Transaction Details**")
    col1, col2 = st.columns(2)
    with col1:
        st.text_input(
            "Merchant",
            value=merchants[(i + j) % len(merchants)],
            key=f"stationary_site_{i}_{j}_merchant",
        )
        st.text_input(
            "Card Number",
            value=fuel_card_number_default(i, j),
            key=f"stationary_site_{i}_{j}_card_number",
        )
        st.text_input(
            "Fuel",
            value=site_default(i, j, "fuel", st.session_state.get("stationary_fuel_card_fuel", fuels[0])),
            key=f"stationary_site_{i}_{j}_fuel",
        )
    with col2:
        default_unit = st.session_state.get("stationary_fuel_card_unit", DEFAULT_COMPACT_FUEL_VOLUME_UNIT)
        st.selectbox(
            "Unit",
            options=COMPACT_FUEL_VOLUME_UNITS,
            index=option_index(COMPACT_FUEL_VOLUME_UNITS, default_unit),
            key=f"stationary_site_{i}_{j}_unit",
        )
        st.number_input(
            "Quantity",
            min_value=0.0,
            step=10.0,
            format="%.2f",
            value=float(st.session_state.get("stationary_fuel_card_quantity", 250.0)),
            key=f"stationary_site_{i}_{j}_quantity",
        )
        st.number_input(
            "Unit Price",
            min_value=0.01,
            step=0.01,
            format="%.2f",
            value=float(st.session_state.get("stationary_fuel_card_unit_price", 1.24)),
            key=f"stationary_site_{i}_{j}_unit_price",
        )
