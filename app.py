from __future__ import annotations

import streamlit as st

from components.sidebar import (
    OUTPUT_FORMATS,
    SCOPE_CONFIG,
    get_allowed_formats,
    get_document_type_config,
    render_sidebar,
)
from components.scope_forms import render_scope_form
from utils.category_registry import get_category_workflow
from utils.generator import (
    generate_document_bytes,
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
#   doc_type : utility_bill | supplier_portal_data | electricity_bill | smart_meter_data
#   format   : pdf | xlsx | csv | docx
# Example: ?scope=2&category=electricity&doc_type=smart_meter_data&format=xlsx
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

if "sidebar_document_type" not in st.session_state:
    _qp_doc_type = _qp.get("doc_type", "").lower()
    if _qp_doc_type:
        st.session_state["sidebar_document_type"] = _qp_doc_type

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
scope, category, document_type, output_format = render_sidebar()
workflow = get_category_workflow(category)

# Keep URL query params in sync with current sidebar selections.
_cat_label = next((lbl for lbl, k in SCOPE_CONFIG[scope]["categories"].items() if k == category), category)
query_params = {
    "scope":    _SCOPE_PARAM_REV.get(scope, "2"),
    "category": _CATEGORY_PARAM_REV.get(_cat_label, category),
}
if document_type:
    query_params["doc_type"] = document_type
if output_format:
    query_params["format"] = _FORMAT_PARAM_REV.get(output_format, output_format.lower())
st.query_params.update(query_params)

# Determine whether the selected combination is currently implemented.
scope_data = SCOPE_CONFIG[scope]
cat_implemented = category in scope_data["implemented"]
document_type_cfg = get_document_type_config(category, document_type or "")
doc_type_implemented = bool(document_type_cfg.get("implemented", False))
allowed_formats = get_allowed_formats(category, document_type or "")
format_allowed = bool(output_format and output_format in allowed_formats)
fmt_implemented = bool(output_format and OUTPUT_FORMATS[output_format]["implemented"])
is_ready = cat_implemented and doc_type_implemented and format_allowed and fmt_implemented

# ---------------------------------------------------------------------------
# Main form
# ---------------------------------------------------------------------------
form_data = render_scope_form(scope, category, document_type)

if form_data is None:
    # Coming-soon branch: nothing more to do.
    st.stop()

# ---------------------------------------------------------------------------
# Format warning (form shown but format not ready)
# ---------------------------------------------------------------------------
if document_type and not doc_type_implemented:
    st.warning(
        f"**{document_type_cfg.get('label', document_type)}** is not yet implemented.",
        icon="⚠️",
    )
elif output_format and not format_allowed:
    allowed_label = ", ".join(allowed_formats)
    st.warning(
        f"**{output_format}** is not available for **{document_type_cfg.get('label', document_type)}**. "
        f"Allowed formats: **{allowed_label}**.",
        icon="⚠️",
    )
elif output_format and not fmt_implemented:
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
    if workflow is None:
        st.error(f"No category workflow is registered for '{category}'.")
        st.stop()

    raw_config = workflow.raw_config_builder(form_data)
    errors = workflow.validator(raw_config)

    if errors:
        st.error("Please fix the following issues before generating:")
        for err in errors:
            st.markdown(f"- {err}")
    else:
        base_name = workflow.build_filename_base(document_type, form_data)
        zip_export = workflow.should_zip_export(document_type, output_format, form_data)
        with st.spinner("Building document… this may take a moment."):
            try:
                fmt_cfg = OUTPUT_FORMATS[output_format]
                filename = base_name + (".zip" if zip_export else fmt_cfg["ext"])
                mime = "application/zip" if zip_export else fmt_cfg["mime"]
                file_bytes = generate_document_bytes(raw_config, output_format)

                st.session_state["generated_file"] = (file_bytes, filename, mime)
                st.session_state["generated_ground_truth"] = workflow.ground_truth_builder(raw_config)
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
