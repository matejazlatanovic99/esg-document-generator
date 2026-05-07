from __future__ import annotations

import streamlit as st


_PRESET_LABELS: dict[str, str] = {
    "realistic": "Realistic",
    "balanced": "Balanced",
    "stress": "Stress",
}


def render_layout_controls() -> None:
    st.checkbox(
        "Enable layout randomization",
        key="doc_layout_enabled",
        help=(
            "Generate deterministic layout variants for the same business data. "
            "The main random seed also controls the resolved layout when this is enabled."
        ),
    )

    preset_label = st.session_state.get("doc_layout_preset", _PRESET_LABELS["balanced"])
    if preset_label not in _PRESET_LABELS.values():
        preset_label = _PRESET_LABELS["balanced"]

    if st.session_state.get("doc_layout_enabled", False):
        options = list(_PRESET_LABELS.values())
        st.selectbox(
            "Layout preset",
            options=options,
            index=options.index(preset_label),
            key="doc_layout_preset",
            help=(
                "Realistic keeps variants close to supplier-like layouts, Balanced mixes realism with parser stress, "
                "and Stress uses the broadest allowed layout changes."
            ),
        )
    else:
        st.session_state.setdefault("doc_layout_preset", _PRESET_LABELS["balanced"])


def collect_layout_form_data(session_state) -> dict:
    preset_label = session_state.get("doc_layout_preset", _PRESET_LABELS["balanced"])
    preset_key = next((key for key, label in _PRESET_LABELS.items() if label == preset_label), "balanced")
    return {
        "doc_layout_enabled": bool(session_state.get("doc_layout_enabled", False)),
        "doc_layout_preset": preset_key,
    }