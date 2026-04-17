from __future__ import annotations

import streamlit as st


def render_supplier_portal_data_form(
    render_document_settings,
    render_financial_period,
    render_heat_global_config,
    render_companies_section,
    collect_form_data,
) -> dict:
    st.subheader("Purchased Heat / Steam / Cooling")
    st.caption("Supplier portal export configuration for purchased heat / steam / cooling.")

    render_document_settings("purchased_heat_steam_cooling", "supplier_portal_data")
    fp_months = render_financial_period()
    render_heat_global_config()
    render_companies_section(fp_months)
    return collect_form_data("supplier_portal_data")
