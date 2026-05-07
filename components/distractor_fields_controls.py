from __future__ import annotations

import streamlit as st


def render_distractor_field_controls() -> None:
    st.checkbox(
        "Add non-relevant fields",
        key="doc_distractor_fields_enabled",
        help=(
            "Render plausible but non-relevant fields in the generated document. "
            "These fields are deterministic and excluded from ground-truth JSON."
        ),
    )


def collect_distractor_field_form_data(session_state) -> dict:
    return {
        "doc_distractor_fields_enabled": bool(
            session_state.get("doc_distractor_fields_enabled", False)
        ),
    }