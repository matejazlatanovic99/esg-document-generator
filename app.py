from __future__ import annotations

import streamlit as st

from components.sidebar import OUTPUT_FORMATS, SCOPE_CONFIG, render_sidebar
from components.scope_forms import render_scope_form
from utils.config import build_raw_config, build_raw_config_electricity, validate_raw_config, validate_raw_config_electricity
from utils.generator import (
    generate_csv_document,
    generate_docx_document,
    generate_json_ground_truth,
    generate_pdf_document,
    generate_xlsx_document,
    generate_electricity_pdf_document,
    generate_electricity_xlsx_document,
    generate_electricity_csv_document,
    generate_electricity_docx_document,
)

st.set_page_config(
    page_title="ESG Document Generator",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Query-param bootstrapping — sets sidebar widget defaults from URL params.
# Supported params:
#   scope    : 1 | 2 | 3u | 3d
#   category : heat | electricity | stationary | mobile | fugitive | ...
#   format   : pdf | xlsx | csv | docx
# Example: ?scope=2&category=electricity&format=xlsx
# ---------------------------------------------------------------------------
_SCOPE_PARAM_MAP: dict[str, str] = {
    "1":  "Scope 1: Direct Emissions",
    "2":  "Scope 2: Indirect Energy",
    "3u": "Scope 3: Upstream",
    "3d": "Scope 3: Downstream",
}
_CATEGORY_PARAM_MAP: dict[str, str] = {
    "heat":         "Purchased Heat / Steam / Cooling",
    "electricity":  "Electricity",
    "stationary":   "Stationary Combustion",
    "mobile":       "Mobile Combustion",
    "fugitive":     "Fugitive Emissions",
    "goods":        "Purchased Goods & Services",
    "capital":      "Capital Goods",
    "travel":       "Business Travel",
    "commuting":    "Employee Commuting",
}
_FORMAT_PARAM_MAP: dict[str, str] = {
    "pdf":  "PDF",
    "xlsx": "XLSX",
    "csv":  "CSV",
    "docx": "DOCX",
}

# Reverse maps: label/key → short param token
_SCOPE_PARAM_REV  = {v: k for k, v in _SCOPE_PARAM_MAP.items()}
_CATEGORY_PARAM_REV = {v: k for k, v in _CATEGORY_PARAM_MAP.items()}
_FORMAT_PARAM_REV = {v: k for k, v in _FORMAT_PARAM_MAP.items()}

_qp = st.query_params
if "sidebar_scope" not in st.session_state:
    _qp_scope = _qp.get("scope", "").lower()
    if _qp_scope in _SCOPE_PARAM_MAP:
        st.session_state["sidebar_scope"] = _SCOPE_PARAM_MAP[_qp_scope]

if "sidebar_category" not in st.session_state:
    _qp_cat = _qp.get("category", "").lower()
    if _qp_cat in _CATEGORY_PARAM_MAP:
        st.session_state["sidebar_category"] = _CATEGORY_PARAM_MAP[_qp_cat]

if "sidebar_format" not in st.session_state:
    _qp_fmt = _qp.get("format", "").lower()
    if _qp_fmt in _FORMAT_PARAM_MAP:
        st.session_state["sidebar_format"] = _FORMAT_PARAM_MAP[_qp_fmt]

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.title("ESG Document Generator")
st.caption(
    "Generate compliance-ready ESG billing documentation "
    "across all GHG Protocol scopes."
)
st.divider()

# ---------------------------------------------------------------------------
# Sidebar → scope / category / format selection
# ---------------------------------------------------------------------------
scope, category, output_format = render_sidebar()

# Keep URL query params in sync with current sidebar selections.
_cat_label = next((lbl for lbl, k in SCOPE_CONFIG[scope]["categories"].items() if k == category), category)
st.query_params.update({
    "scope":    _SCOPE_PARAM_REV.get(scope, "2"),
    "category": _CATEGORY_PARAM_REV.get(_cat_label, category),
    "format":   _FORMAT_PARAM_REV.get(output_format, output_format.lower()),
})

# Determine whether the selected combination is currently implemented.
scope_data = SCOPE_CONFIG[scope]
cat_implemented = category in scope_data["implemented"]
fmt_implemented = OUTPUT_FORMATS[output_format]["implemented"]
is_ready = cat_implemented and fmt_implemented

# ---------------------------------------------------------------------------
# Main form
# ---------------------------------------------------------------------------
form_data = render_scope_form(scope, category)

if form_data is None:
    # Coming-soon branch: nothing more to do.
    st.stop()

# ---------------------------------------------------------------------------
# Format warning (form shown but format not ready)
# ---------------------------------------------------------------------------
if not fmt_implemented:
    st.warning(
        f"**{output_format}** format is not yet implemented. "
        "Switch to **PDF** in the sidebar to generate a document.",
        icon="⚠️",
    )

st.divider()

# ---------------------------------------------------------------------------
# Generate button
# ---------------------------------------------------------------------------
_, btn_col, _ = st.columns([3, 2, 3])
with btn_col:
    generate_clicked = st.button(
        "Generate Document",
        type="primary",
        use_container_width=True,
        disabled=not is_ready,
    )

# ---------------------------------------------------------------------------
# Generation logic
# ---------------------------------------------------------------------------
if generate_clicked and is_ready:
    _is_electricity = form_data.get("_category") == "electricity"

    if _is_electricity:
        raw_config = build_raw_config_electricity(form_data)
        errors = validate_raw_config_electricity(raw_config)
    else:
        raw_config = build_raw_config(form_data)
        errors = validate_raw_config(raw_config)

    if errors:
        st.error("Please fix the following issues before generating:")
        for err in errors:
            st.markdown(f"- {err}")
    else:
        # Build filename: {category}_{fp_start}_{fp_end}.{ext}
        # e.g. heat_2025-04_2025-09.xlsx  or  electricity_2026-01_2026-12.pdf
        _cat_slug = {
            "purchased_heat_steam_cooling": "heat",
            "electricity": "electricity",
        }.get(category, category.split("_")[0])
        _fp_start = form_data.get("fp_start", "")[:7]   # "YYYY-MM"
        _fp_end   = form_data.get("fp_end",   "")[:7]
        _period_slug = f"{_fp_start}_{_fp_end}" if _fp_start != _fp_end else _fp_start
        _base_name = f"{_cat_slug}_{_period_slug}"
        with st.spinner("Building document… this may take a moment."):
            try:
                fmt_cfg = OUTPUT_FORMATS[output_format]
                filename = _base_name + fmt_cfg["ext"]

                if _is_electricity:
                    if output_format == "PDF":
                        file_bytes = generate_electricity_pdf_document(raw_config)
                    elif output_format == "XLSX":
                        file_bytes = generate_electricity_xlsx_document(raw_config)
                    elif output_format == "CSV":
                        file_bytes = generate_electricity_csv_document(raw_config)
                    elif output_format == "DOCX":
                        file_bytes = generate_electricity_docx_document(raw_config)
                    else:
                        raise NotImplementedError(f"Format {output_format} is not yet supported.")
                    st.session_state["generated_ground_truth"] = b"{}"
                elif output_format == "PDF":
                    file_bytes = generate_pdf_document(raw_config)
                elif output_format == "XLSX":
                    file_bytes = generate_xlsx_document(raw_config)
                elif output_format == "CSV":
                    file_bytes = generate_csv_document(raw_config)
                elif output_format == "DOCX":
                    file_bytes = generate_docx_document(raw_config)
                else:
                    raise NotImplementedError(f"Format {output_format} is not yet supported.")

                st.session_state["generated_file"] = (file_bytes, filename, fmt_cfg["mime"])
                if not _is_electricity:
                    st.session_state["generated_ground_truth"] = generate_json_ground_truth(raw_config)
                st.session_state.pop("generation_error", None)
                st.success("Document generated successfully!")
            except Exception as exc:
                import traceback
                st.session_state["generation_error"] = (
                    str(exc) + "\n\n" + traceback.format_exc()
                )
                st.session_state.pop("generated_file", None)

if "generation_error" in st.session_state:
    st.error(f"Generation failed: {st.session_state['generation_error']}")

# ---------------------------------------------------------------------------
# Download section
# ---------------------------------------------------------------------------
if "generated_file" in st.session_state:
    file_bytes, filename, mime = st.session_state["generated_file"]
    st.divider()
    _, dl_col, _ = st.columns([3, 2, 3])
    with dl_col:
        st.download_button(
            label=f"Download {filename}",
            data=file_bytes,
            file_name=filename,
            mime=mime,
            type="primary",
            use_container_width=True,
        )
        if "generated_ground_truth" in st.session_state:
            gt_filename = filename.rsplit(".", 1)[0] + "_ground_truth.json"
            st.download_button(
                label=f"Download ground-truth JSON",
                data=st.session_state["generated_ground_truth"],
                file_name=gt_filename,
                mime="application/json",
                use_container_width=True,
            )
