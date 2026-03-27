from __future__ import annotations

import streamlit as st

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
        "implemented": set(),
    },
    "Scope 2: Indirect Energy": {
        "categories": {
            "Electricity": "electricity",
            "Purchased Heat / Steam / Cooling": "purchased_heat_steam_cooling",
        },
        "implemented": {"purchased_heat_steam_cooling"},
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

OUTPUT_FORMATS: dict[str, dict] = {
    "PDF": {"mime": "application/pdf", "ext": ".pdf", "implemented": True},
    "DOCX": {
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "ext": ".docx",
        "implemented": True,
    },
    "XLSX": {
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "ext": ".xlsx",
        "implemented": True,
    },
    "CSV": {"mime": "text/csv", "ext": ".csv", "implemented": True},
}


def render_sidebar() -> tuple[str, str, str]:
    """Render the sidebar and return (scope_label, category_key, output_format)."""
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

        implemented_formats = [f for f, v in OUTPUT_FORMATS.items() if v["implemented"]]
        coming_soon_formats = [f for f, v in OUTPUT_FORMATS.items() if not v["implemented"]]
        format_display = implemented_formats + [f"{f} (Coming Soon)" for f in coming_soon_formats]

        default_fmt_idx = format_display.index("XLSX") if "XLSX" in format_display else 0
        output_format_display: str = st.selectbox(
            "Output Format",
            options=format_display,
            index=default_fmt_idx,
            key="sidebar_format",
        )
        output_format = output_format_display.removesuffix(" (Coming Soon)")

        st.divider()

        if not is_implemented:
            st.warning(f"**{clean_label}** is not yet implemented.", icon="🚧")
        elif not OUTPUT_FORMATS[output_format]["implemented"]:
            st.info(f"**{output_format}** format is coming soon.", icon="ℹ️")

        st.caption("ESG Document Generator v1.0")

    return scope_label, selected_cat_key, output_format
