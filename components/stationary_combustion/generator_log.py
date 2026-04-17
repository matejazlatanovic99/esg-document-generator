from __future__ import annotations

import random

import streamlit as st


def init_log_defaults(fuels: list[str]) -> None:
    if "stationary_log_runs_per_month" in st.session_state:
        return
    rng = random.Random(20260416)
    st.session_state["stationary_log_fuel"] = rng.choice(fuels)
    st.session_state["stationary_log_unit"] = "L"
    st.session_state["stationary_log_runs_per_month"] = rng.randint(2, 5)
    st.session_state["stationary_log_quantity_mode_label"] = "Tank Level Change"
    st.session_state["stationary_log_tank_capacity"] = float(rng.randrange(400, 1600, 50))
    st.session_state["stationary_log_fuel_used_per_hour"] = round(rng.uniform(9.0, 28.0), 1)
    st.session_state["stationary_log_run_hours_min"] = 0.5
    st.session_state["stationary_log_run_hours_max"] = 5.0


def render_log_defaults(fuels: list[str]) -> None:
    init_log_defaults(fuels)
    st.markdown("#### Generator Log Defaults")
    st.caption("These defaults apply to every generator unless you override them.")
    col1, col2 = st.columns(2)
    with col1:
        st.selectbox("Default Fuel", options=fuels, key="stationary_log_fuel")
        st.selectbox("Default Unit", options=["L", "gal"], key="stationary_log_unit")
        st.number_input("Runs Per Month", min_value=1, max_value=31, step=1, key="stationary_log_runs_per_month")
    with col2:
        st.radio(
            "Fuel Quantity Mode",
            options=["Tank Level Change", "Explicit Fuel Used"],
            horizontal=True,
            key="stationary_log_quantity_mode_label",
        )
        st.number_input("Tank Capacity", min_value=50.0, step=25.0, format="%.1f", key="stationary_log_tank_capacity")
        st.number_input("Fuel Used Per Hour", min_value=0.1, step=0.5, format="%.1f", key="stationary_log_fuel_used_per_hour")
    col3, col4 = st.columns(2)
    with col3:
        st.number_input("Minimum Run Hours", min_value=0.25, step=0.25, format="%.2f", key="stationary_log_run_hours_min")
    with col4:
        st.number_input("Maximum Run Hours", min_value=0.25, step=0.25, format="%.2f", key="stationary_log_run_hours_max")
    st.divider()


def render_log_site_fields(i: int, j: int) -> None:
    st.markdown("**Log Generation Defaults**")
    col1, col2 = st.columns(2)
    with col1:
        st.selectbox("Unit", options=["L", "gal"], index=0, key=f"stationary_site_{i}_{j}_unit")
        st.number_input(
            "Runs Per Month",
            min_value=1,
            max_value=31,
            step=1,
            value=int(st.session_state.get("stationary_log_runs_per_month", 3)),
            key=f"stationary_site_{i}_{j}_runs_per_month",
        )
        st.number_input(
            "Fuel Used Per Hour",
            min_value=0.1,
            step=0.5,
            format="%.1f",
            value=float(st.session_state.get("stationary_log_fuel_used_per_hour", 15.0)),
            key=f"stationary_site_{i}_{j}_fuel_used_per_hour",
        )
    with col2:
        st.radio(
            "Site Quantity Mode",
            options=["Tank Level Change", "Explicit Fuel Used"],
            horizontal=True,
            key=f"stationary_site_{i}_{j}_quantity_mode_label",
        )
        st.number_input(
            "Tank Capacity",
            min_value=50.0,
            step=25.0,
            format="%.1f",
            value=float(st.session_state.get("stationary_log_tank_capacity", 800.0)),
            key=f"stationary_site_{i}_{j}_tank_capacity",
        )
        st.number_input(
            "Minimum Run Hours",
            min_value=0.25,
            step=0.25,
            format="%.2f",
            value=float(st.session_state.get("stationary_log_run_hours_min", 0.5)),
            key=f"stationary_site_{i}_{j}_run_hours_min",
        )
        st.number_input(
            "Maximum Run Hours",
            min_value=0.25,
            step=0.25,
            format="%.2f",
            value=float(st.session_state.get("stationary_log_run_hours_max", 5.0)),
            key=f"stationary_site_{i}_{j}_run_hours_max",
        )

