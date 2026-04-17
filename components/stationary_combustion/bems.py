from __future__ import annotations

import streamlit as st


def bems_report_type_value(report_types: dict[str, str]) -> str:
    label = st.session_state.get("stationary_bems_report_type_label", "Equipment Trend Report")
    return report_types.get(label, "equipment_trend_report")


def render_bems_report_type_selector(report_types: dict[str, str]) -> None:
    st.markdown("#### BEMS Report Type")
    st.radio(
        "Report Type",
        options=list(report_types.keys()),
        horizontal=True,
        key="stationary_bems_report_type_label",
        label_visibility="collapsed",
    )


def init_bems_defaults(asset_defaults: list[dict]) -> None:
    if "stationary_bems_interval_label" in st.session_state:
        return
    st.session_state["stationary_bems_interval_label"] = "60 minutes"
    st.session_state["stationary_bems_assets_per_site"] = len(asset_defaults)


def render_bems_defaults(report_type_value: str, asset_defaults: list[dict]) -> None:
    init_bems_defaults(asset_defaults)
    st.markdown("#### BEMS Defaults")
    if report_type_value == "time_series_trend_export":
        st.caption("These defaults control the generated time-series trend export.")
        col1, col2 = st.columns(2)
        with col1:
            st.selectbox("Time-Series Interval", options=["15 minutes", "30 minutes", "60 minutes"], key="stationary_bems_interval_label")
        with col2:
            st.number_input("Default Assets Per Site", min_value=1, max_value=10, step=1, key="stationary_bems_assets_per_site")
    else:
        st.caption("These defaults control the generated equipment trend report and summary layout.")
        st.caption("Time-series interval is hidden here because it only affects time-series trend exports.")
        st.number_input("Default Assets Per Site", min_value=1, max_value=10, step=1, key="stationary_bems_assets_per_site")
    st.caption("The selected report type is used for every output format. Format now controls only the container, not the report shape.")
    st.divider()


def render_bems_site_fields(i: int, j: int, asset_defaults: list[dict], bems_asset_default, optional_field) -> None:
    st.markdown("**BEMS Assets**")
    asset_count = st.number_input(
        "Number of assets",
        min_value=1,
        max_value=10,
        value=int(st.session_state.get("stationary_bems_assets_per_site", len(asset_defaults))),
        step=1,
        key=f"stationary_site_{i}_{j}_asset_count",
    )
    for asset_idx in range(int(asset_count)):
        asset_tag = st.session_state.get(
            f"stationary_site_{i}_{j}_asset_{asset_idx}_tag",
            bems_asset_default(asset_idx, "asset_tag", f"AST-{asset_idx + 1:02d}"),
        )
        with st.expander(f"Asset {asset_idx + 1}: {asset_tag}", expanded=(asset_idx == 0)):
            render_bems_asset_fields(i, j, asset_idx, bems_asset_default, optional_field)


def render_bems_asset_fields(i: int, j: int, asset_idx: int, bems_asset_default, optional_field) -> None:
    col1, col2 = st.columns(2)
    with col1:
        st.text_input(
            "Asset Tag",
            value=bems_asset_default(asset_idx, "asset_tag", f"AST-{asset_idx + 1:02d}"),
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_tag",
        )
        optional_field(
            st.text_input,
            "Equipment Name",
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_equipment_name",
            value=bems_asset_default(asset_idx, "equipment_name", ""),
            help="Time-series exports often only include the asset tag, not a human-readable equipment name.",
        )
        optional_field(
            st.text_input,
            "Emission Source",
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_emission_source",
            value=bems_asset_default(asset_idx, "emission_source", ""),
            help="Can be inferred from asset naming and is not always explicitly available in BEMS exports.",
        )
        optional_field(
            st.text_input,
            "Sensor Name",
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_sensor_name",
            value=bems_asset_default(asset_idx, "sensor_name", ""),
            help="Trend exports sometimes include generic sensor labels or omit them entirely.",
        )
    with col2:
        optional_field(
            st.text_input,
            "Fuel",
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_fuel",
            value=bems_asset_default(asset_idx, "fuel", ""),
            help="Fuel type may need to be inferred from the equipment or sensor context.",
        )
        st.selectbox(
            "Unit",
            options=["kWh", "L", "m3"],
            index=["kWh", "L", "m3"].index(bems_asset_default(asset_idx, "unit", "kWh")),
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_unit",
        )
        st.number_input(
            "Consumption",
            min_value=0.0,
            step=10.0,
            format="%.2f",
            value=float(bems_asset_default(asset_idx, "quantity", 0.0)),
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_quantity",
        )
        optional_field(
            st.number_input,
            "Operating Hours",
            key=f"stationary_site_{i}_{j}_asset_{asset_idx}_operating_hours",
            min_value=0.0,
            step=1.0,
            format="%.1f",
            value=float(bems_asset_default(asset_idx, "operating_hours", 0.0)),
            help="Useful in equipment summaries, but commonly missing from raw BEMS trend exports.",
        )
