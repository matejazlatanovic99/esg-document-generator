"""Shared Streamlit UI controls for the invalid-data generation feature.

Call ``render_invalid_data_controls()`` inside any ``_render_document_settings``
expander to add the three invalid-data widgets.

Call ``collect_invalid_data_form_data(session_state)`` inside any
``_collect_*_form_data`` function to pull the three values into the returned
form-data dict.
"""
from __future__ import annotations

import streamlit as st

_PRESET_OPTIONS = ["light", "mixed", "aggressive"]
_PRESET_LABELS = {
    "light": "Light – subtle truncation / swap / format damage",
    "mixed": "Mixed – light + type-mismatch and short random replacement",
    "aggressive": "Aggressive – mixed + obvious replacements and unit/currency damage",
}
_PRESET_LABEL_LIST = [_PRESET_LABELS[k] for k in _PRESET_OPTIONS]
_LABEL_TO_PRESET = {v: k for k, v in _PRESET_LABELS.items()}


def _preset_label(preset: str) -> str:
    return _PRESET_LABELS.get(preset, _PRESET_LABELS["mixed"])


def render_invalid_data_controls() -> None:
    """Render the three invalid-data widgets (checkbox + selectbox + slider).

    Widget keys written to ``st.session_state``:
    - ``doc_invalid_data_enabled``  (bool)
    - ``doc_invalid_data_preset``   (str label, not the short key)
    - ``doc_invalid_data_rate``     (int 0-100)
    """
    st.checkbox(
        "Enable invalid data",
        key="doc_invalid_data_enabled",
        help=(
            "After building all records, corrupt a percentage of visible export "
            "fields with bad values.  Validation and calculations are unaffected."
        ),
    )

    if st.session_state.get("doc_invalid_data_enabled", False):
        col_preset, col_rate = st.columns([3, 1])
        with col_preset:
            current_preset_label = _preset_label(
                st.session_state.get("doc_invalid_data_preset_key", "mixed")
            )
            st.selectbox(
                "Preset",
                options=_PRESET_LABEL_LIST,
                index=_PRESET_LABEL_LIST.index(current_preset_label),
                key="doc_invalid_data_preset",
                help="Controls how aggressively fields are corrupted.",
            )
        with col_rate:
            st.number_input(
                "Bad field %",
                min_value=1,
                max_value=100,
                value=int(st.session_state.get("doc_invalid_data_rate", 15)),
                step=1,
                key="doc_invalid_data_rate",
                help="Percentage of eligible fields to corrupt per record.",
            )


def collect_invalid_data_form_data(s) -> dict:
    """Return the three invalid-data keys ready for inclusion in form data.

    The returned dict should be unpacked/merged into the main form data dict.
    """
    enabled = bool(s.get("doc_invalid_data_enabled", False))
    preset_label = s.get("doc_invalid_data_preset", _PRESET_LABELS["mixed"])
    preset_key = _LABEL_TO_PRESET.get(str(preset_label), "mixed")
    rate = int(s.get("doc_invalid_data_rate", 15))
    return {
        "doc_invalid_data_enabled": enabled,
        "doc_invalid_data_preset": preset_key,
        "doc_invalid_data_rate": rate,
    }
