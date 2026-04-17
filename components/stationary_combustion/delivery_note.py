from __future__ import annotations

import random

import streamlit as st


def init_delivery_note_defaults(fuels: list[str]) -> None:
    if "stationary_delivery_note_quantity" in st.session_state:
        return
    rng = random.Random(20260417)
    st.session_state["stationary_delivery_note_fuel"] = rng.choice(fuels)
    st.session_state["stationary_delivery_note_unit"] = "Litres"
    st.session_state["stationary_delivery_note_quantity"] = float(rng.randrange(1500, 5000, 50))


def render_delivery_note_defaults(fuels: list[str]) -> None:
    init_delivery_note_defaults(fuels)
    st.markdown("#### Delivery Note Defaults")
    st.caption("These defaults apply to every delivery note unless you override them.")
    col1, col2 = st.columns(2)
    with col1:
        st.selectbox("Default Fuel", options=fuels, key="stationary_delivery_note_fuel")
        st.selectbox("Default Unit", options=["Litres", "Gallons"], key="stationary_delivery_note_unit")
    with col2:
        st.number_input("Default Quantity", min_value=0.0, step=50.0, format="%.2f", key="stationary_delivery_note_quantity")
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
        st.selectbox("Unit", options=["Litres", "Gallons"], index=0, key=f"stationary_site_{i}_{j}_unit")
    with col2:
        st.number_input(
            "Delivered Quantity",
            min_value=0.0,
            step=50.0,
            format="%.2f",
            value=float(st.session_state.get("stationary_delivery_note_quantity", 3200.0)),
            key=f"stationary_site_{i}_{j}_quantity",
        )

