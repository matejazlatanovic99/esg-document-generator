from __future__ import annotations

import streamlit as st

from components.sidebar import OUTPUT_FORMATS, SCOPE_CONFIG, render_sidebar
from components.scope_forms import render_scope_form
from utils.config import build_raw_config, validate_raw_config
from utils.generator import generate_csv_document, generate_docx_document, generate_json_ground_truth, generate_pdf_document, generate_xlsx_document

st.set_page_config(
    page_title="ESG Document Generator",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
    raw_config = build_raw_config(form_data)
    errors = validate_raw_config(raw_config)

    if errors:
        st.error("Please fix the following issues before generating:")
        for err in errors:
            st.markdown(f"- {err}")
    else:
        with st.spinner("Building document… this may take a moment."):
            try:
                fmt_cfg = OUTPUT_FORMATS[output_format]
                base = raw_config["document"]["pdf_filename"].rsplit(".", 1)[0]
                filename = base + fmt_cfg["ext"]

                if output_format == "PDF":
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
