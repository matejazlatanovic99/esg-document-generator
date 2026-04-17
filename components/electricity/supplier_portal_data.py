from __future__ import annotations

import streamlit as st


def render_electricity_supplier_portal_data_form(
    render_document_settings,
    render_financial_period,
    render_electricity_global_config,
    render_electricity_companies_section,
    collect_electricity_form_data,
) -> dict:
    st.subheader("Electricity")
    st.caption("Supplier portal export configuration for Scope 2 purchased electricity.")

    render_document_settings("electricity", "supplier_portal_data")
    fp_months = render_financial_period()
    render_electricity_global_config()
    render_electricity_companies_section(fp_months)
    return collect_electricity_form_data("supplier_portal_data")
