from __future__ import annotations

import streamlit as st

from utils.document_catalog import (
    DOCUMENT_TYPES,
    OUTPUT_FORMATS,
    get_allowed_formats,
    get_default_document_type,
    get_default_format,
    get_document_type_config,
    get_document_type_options,
)

COMPANIES: dict[str, dict] = {
    "Toyota Kinto": {
        "country": "Japan",
        "industry": "Automotive / Mobility Services",
        "reporting_year": 2026,
        "currency": "JPY",
        "fiscal_year_end": "March 31",
    },
    "Toyota Financial Services Corporation": {
        "country": "United States",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "USD",
        "fiscal_year_end": "March 31",
    },
    "Toyota Financial Services UK group": {
        "country": "United Kingdom",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "GBP",
        "fiscal_year_end": "December 31",
    },
    "Toyota Financial Services Slovakia": {
        "country": "Slovakia",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "EUR",
        "fiscal_year_end": "December 31",
    },
    "TEST Subsidiary": {
        "country": "Slovakia",
        "industry": "Test / Placeholder",
        "reporting_year": 2026,
        "currency": "EUR",
        "fiscal_year_end": "December 31",
    },
    "Toyota Financial Services Belgium": {
        "country": "Belgium",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "EUR",
        "fiscal_year_end": "December 31",
    },
    "Toyota Financial Services UK": {
        "country": "United Kingdom",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "GBP",
        "fiscal_year_end": "December 31",
    },
    "Toyota Financial Services Ireland": {
        "country": "Ireland",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "EUR",
        "fiscal_year_end": "December 31",
    },
    "Toyota Financial Services Danmark": {
        "country": "Denmark",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "DKK",
        "fiscal_year_end": "December 31",
    },
    "Toyota Financial Services Hungary": {
        "country": "Hungary",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "HUF",
        "fiscal_year_end": "December 31",
    },
    "Prazna": {
        "country": "Slovakia",
        "industry": "—",
        "reporting_year": 2026,
        "currency": "EUR",
        "fiscal_year_end": "December 31",
    },
    "Toyota Kreditbank GMBH": {
        "country": "Germany",
        "industry": "Financial Services / Automotive",
        "reporting_year": 2026,
        "currency": "EUR",
        "fiscal_year_end": "December 31",
    },
}

NEW_COMPANY_PLACEHOLDER: dict = {
    "country": "—",
    "industry": "—",
    "reporting_year": 2026,
    "currency": "EUR",
    "fiscal_year_end": "December 31",
}

SCOPE_CONFIG: dict[str, dict] = {
    "Scope 1: Direct Emissions": {
        "categories": {
            "Stationary Combustion": "stationary_combustion",
            "Mobile Combustion": "mobile_combustion",
            "Fugitive Emissions": "fugitive_emissions",
        },
        "implemented": {"stationary_combustion"},
    },
    "Scope 2: Indirect Energy": {
        "categories": {
            "Electricity": "electricity",
            "Purchased Heat / Steam / Cooling": "purchased_heat_steam_cooling",
        },
        "implemented": {"purchased_heat_steam_cooling", "electricity"},
    },
    "Scope 3: Upstream": {
        "categories": {
            "Purchased Goods & Services": "purchased_goods_services",
            "Capital Goods": "capital_goods",
            "Fuel & Energy Related Activities": "fuel_energy_related_activities",
            "Upstream Transport & Distribution": "upstream_transport_distribution",
            "Waste in Operations": "waste_generated_in_operations",
            "Business Travel": "business_travel",
            "Employee Commuting": "employee_commuting",
            "Upstream Leased Assets": "upstream_leased_assets",
        },
        "implemented": set(),
    },
    "Scope 3: Downstream": {
        "categories": {
            "Downstream Transport & Distribution": "downstream_transport_distribution",
            "Processing of Sold Products": "processing_of_sold_products",
            "Use of Sold Products": "use_of_sold_products",
            "End-of-Life Treatment": "end_of_life_treatment_sold_products",
            "Downstream Leased Assets": "downstream_leased_assets",
            "Franchises": "franchises",
            "Investments": "investments",
        },
        "implemented": set(),
    },
}

def render_sidebar() -> tuple[str, str, str | None, str | None]:
    """Render the sidebar and return (scope_label, category_key, document_type_key, output_format)."""
    with st.sidebar:
        st.header("Configuration")

        scope_label: str = st.selectbox(
            "GHG Scope",
            options=list(SCOPE_CONFIG.keys()),
            index=1,
            key="sidebar_scope",
        )
        scope_data = SCOPE_CONFIG[scope_label]

        cat_labels = list(scope_data["categories"].keys())
        cat_keys = list(scope_data["categories"].values())
        cat_display = [
            label
            if cat_keys[i] in scope_data["implemented"]
            else f"{label} (Coming Soon)"
            for i, label in enumerate(cat_labels)
        ]

        default_idx = next(
            (i for i, k in enumerate(cat_keys) if k in scope_data["implemented"]),
            0,
        )

        cat_display_label: str = st.selectbox(
            "Category",
            options=cat_display,
            index=default_idx,
            key="sidebar_category",
        )

        clean_label = cat_display_label.removesuffix(" (Coming Soon)")
        selected_cat_key = scope_data["categories"][clean_label]
        is_implemented = selected_cat_key in scope_data["implemented"]

        st.divider()

        document_type_options = get_document_type_options(selected_cat_key)
        selected_doc_type = None
        clean_doc_type_label = ""
        doc_type_implemented = False

        if document_type_options:
            document_type_keys = list(document_type_options.keys())
            document_type_labels = [cfg["label"] for cfg in document_type_options.values()]
            document_type_label_map = dict(zip(document_type_labels, document_type_keys))
            document_type_display = [
                label
                if document_type_options[key].get("implemented", False)
                else f"{label} (Coming Soon)"
                for label, key in zip(document_type_labels, document_type_keys)
            ]

            default_doc_type = get_default_document_type(selected_cat_key)
            current_doc_type = st.session_state.get("sidebar_document_type")
            if current_doc_type not in document_type_options:
                current_doc_type = default_doc_type
                if current_doc_type is not None:
                    st.session_state["sidebar_document_type"] = current_doc_type

            current_doc_label = document_type_options[current_doc_type]["label"] if current_doc_type in document_type_options else document_type_labels[0]
            if st.session_state.get("sidebar_document_type_display") not in document_type_display:
                st.session_state["sidebar_document_type_display"] = current_doc_label

            default_doc_idx = document_type_keys.index(current_doc_type) if current_doc_type in document_type_keys else 0
            selected_doc_type_display: str = st.selectbox(
                "Document Type",
                options=document_type_display,
                index=default_doc_idx,
                key="sidebar_document_type_display",
            )
            clean_doc_type_label = selected_doc_type_display.removesuffix(" (Coming Soon)")
            selected_doc_type = document_type_label_map.get(clean_doc_type_label)
            if selected_doc_type is not None:
                st.session_state["sidebar_document_type"] = selected_doc_type

            selected_doc_type_cfg = get_document_type_config(selected_cat_key, selected_doc_type or "")
            doc_type_implemented = bool(selected_doc_type_cfg.get("implemented", False))
        else:
            st.caption("Document types will appear here when this category is implemented.")

        allowed_formats = get_allowed_formats(selected_cat_key, selected_doc_type or "")
        format_display = [
            fmt
            if OUTPUT_FORMATS[fmt]["implemented"]
            else f"{fmt} (Coming Soon)"
            for fmt in allowed_formats
        ]

        default_format = get_default_format(selected_cat_key, selected_doc_type or "")
        current_format = st.session_state.get("sidebar_format")
        if current_format not in allowed_formats:
            current_format = default_format
            if current_format is not None:
                st.session_state["sidebar_format"] = current_format

        if format_display:
            if st.session_state.get("sidebar_format_display") not in format_display:
                st.session_state["sidebar_format_display"] = current_format or format_display[0]
            default_fmt_idx = allowed_formats.index(current_format) if current_format in allowed_formats else 0
            output_format_display: str = st.selectbox(
                "Output Format",
                options=format_display,
                index=default_fmt_idx,
                key="sidebar_format_display",
            )
            output_format = output_format_display.removesuffix(" (Coming Soon)")
            st.session_state["sidebar_format"] = output_format
        else:
            st.caption("Available formats depend on the selected document type.")
            output_format = None

        st.divider()

        if not is_implemented:
            st.warning(f"**{clean_label}** is not yet implemented.", icon="🚧")
        elif selected_doc_type and not doc_type_implemented:
            st.info(f"**{clean_doc_type_label}** is coming soon.", icon="ℹ️")
        elif output_format and not OUTPUT_FORMATS[output_format]["implemented"]:
            st.info(f"**{output_format}** format is coming soon.", icon="ℹ️")

        st.caption("ESG Document Generator v1.0")

    return scope_label, selected_cat_key, selected_doc_type, output_format
