from __future__ import annotations

import random

import streamlit as st


def init_invoice_defaults(fuels: list[str]) -> None:
    if "stationary_invoice_quantity" in st.session_state:
        return
    rng = random.Random(20260415)
    st.session_state["stationary_invoice_fuel"] = rng.choice(fuels)
    st.session_state["stationary_invoice_unit"] = "Litres"
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
        st.selectbox("Default Fuel", options=fuels, key="stationary_invoice_fuel")
        st.selectbox("Default Unit", options=["Litres", "Gallons"], key="stationary_invoice_unit")
    with col2:
        st.number_input("Default Quantity", min_value=0.0, step=50.0, format="%.2f", key="stationary_invoice_quantity")
        st.number_input("Default Unit Price", min_value=0.01, step=0.01, format="%.2f", key="stationary_invoice_unit_price")
    col3, col4 = st.columns(2)
    with col3:
        st.number_input("Default Delivery Charge", min_value=0.0, step=5.0, format="%.2f", key="stationary_invoice_delivery_charge")
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
        st.selectbox("Unit", options=["Litres", "Gallons"], index=0, key=f"stationary_site_{i}_{j}_unit")
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
            "Unit Price",
            min_value=0.01,
            step=0.01,
            format="%.2f",
            value=float(st.session_state.get("stationary_invoice_unit_price", 1.12)),
            key=f"stationary_site_{i}_{j}_unit_price",
        )
        st.number_input(
            "Delivery Charge",
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

