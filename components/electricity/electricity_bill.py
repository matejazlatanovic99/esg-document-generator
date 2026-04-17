from __future__ import annotations

import streamlit as st


def render_electricity_bill_form(
    render_document_settings,
    render_financial_period,
    render_electricity_global_config,
    render_electricity_companies_section,
    collect_electricity_form_data,
) -> dict:
    st.subheader("Electricity")
    st.caption("Scope 2 purchased electricity bill configuration.")

    render_document_settings("electricity", "electricity_bill")
    fp_months = render_financial_period()
    render_electricity_global_config()
    render_electricity_companies_section(fp_months)
    return collect_electricity_form_data("electricity_bill")
