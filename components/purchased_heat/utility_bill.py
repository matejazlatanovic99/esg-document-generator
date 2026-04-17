from __future__ import annotations

import streamlit as st


def render_utility_bill_form(
    render_document_settings,
    render_financial_period,
    render_heat_global_config,
    render_companies_section,
    collect_form_data,
) -> dict:
    st.subheader("Purchased Heat / Steam / Cooling")
    st.caption("District heating and cooling utility bill configuration.")

    render_document_settings("purchased_heat_steam_cooling", "utility_bill")
    fp_months = render_financial_period()
    render_heat_global_config()
    render_companies_section(fp_months)
    return collect_form_data("utility_bill")

