from __future__ import annotations

import random

import streamlit as st

from components.stationary_combustion.units import (
    DEFAULT_LONG_FUEL_VOLUME_UNIT,
    LONG_FUEL_VOLUME_UNITS,
    option_index,
)


def init_delivery_note_defaults(fuels: list[str]) -> None:
    st.session_state.setdefault("stationary_delivery_note_fuel_randomize", True)
    st.session_state.setdefault("stationary_delivery_note_unit_randomize", True)
    st.session_state.setdefault("stationary_delivery_note_quantity_randomize", True)
    if "stationary_delivery_note_quantity" in st.session_state:
        return
    rng = random.Random(20260417)
    st.session_state["stationary_delivery_note_fuel"] = rng.choice(fuels)
    st.session_state["stationary_delivery_note_unit"] = DEFAULT_LONG_FUEL_VOLUME_UNIT
    st.session_state["stationary_delivery_note_quantity"] = float(rng.randrange(1500, 5000, 50))


def render_delivery_note_defaults(fuels: list[str]) -> None:
    init_delivery_note_defaults(fuels)
    st.markdown("#### Delivery Note Defaults")
    st.caption("These defaults apply to every delivery note unless you override them.")
    col1, col2 = st.columns(2)
    with col1:
        randomize_fuel = st.checkbox(
            "Randomize fuel",
            key="stationary_delivery_note_fuel_randomize",
            help="New equipment entries pick an initial fuel from the Default Fuel dropdown options.",
        )
        st.selectbox("Default Fuel", options=fuels, key="stationary_delivery_note_fuel", disabled=randomize_fuel)
        randomize_unit = st.checkbox(
            "Randomize unit",
            key="stationary_delivery_note_unit_randomize",
            help="New equipment entries pick an initial unit from the delivery-note unit options.",
        )
        st.selectbox("Default Unit", options=LONG_FUEL_VOLUME_UNITS, key="stationary_delivery_note_unit", disabled=randomize_unit)
    with col2:
        randomize_quantity = st.checkbox(
            "Randomize quantity",
            key="stationary_delivery_note_quantity_randomize",
            help="New equipment entries get an initial delivered quantity in the delivery-note range.",
        )
        st.number_input(
            "Default Quantity",
            min_value=0.0,
            step=50.0,
            format="%.2f",
            key="stationary_delivery_note_quantity",
            disabled=randomize_quantity,
        )
    st.caption("Delivery notes usually reflect a single delivery event inside the selected reporting period.")
    st.divider()


def render_delivery_note_site_fields(i: int, j: int, site_default, fuels: list[str]) -> None:
    st.markdown("**Delivery Details**")
    col1, col2 = st.columns(2)
    with col1:
        st.text_input(
            "Fuel",
            value=site_default(i, j, "fuel", st.session_state.get("stationary_delivery_note_fuel", fuels[0])),
            key=f"stationary_site_{i}_{j}_fuel",
        )
        default_unit = st.session_state.get("stationary_delivery_note_unit", DEFAULT_LONG_FUEL_VOLUME_UNIT)
        st.selectbox(
            "Unit",
            options=LONG_FUEL_VOLUME_UNITS,
            index=option_index(LONG_FUEL_VOLUME_UNITS, default_unit),
            key=f"stationary_site_{i}_{j}_unit",
        )
    with col2:
        st.number_input(
            "Delivered Quantity",
            min_value=0.0,
            step=50.0,
            format="%.2f",
            value=float(st.session_state.get("stationary_delivery_note_quantity", 3200.0)),
            key=f"stationary_site_{i}_{j}_quantity",
        )
