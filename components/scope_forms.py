from __future__ import annotations

import streamlit as st

from utils.category_registry import get_category_workflow


def render_scope_form(scope: str, category: str, document_type: str | None) -> dict | None:
    """Render the appropriate form. Returns form_data dict or None if not implemented."""
    workflow = get_category_workflow(category)
    if workflow is not None:
        return workflow.form_renderer(document_type)
    _render_coming_soon(scope, category)
    return None


def _render_coming_soon(scope: str, category: str) -> None:
    label = category.replace("_", " ").title()
    st.info(
        f"**{label}** ({scope}) is not yet available.\n\n"
        "This generator currently supports **Scope 2 – Purchased Heat / Steam / Cooling**. "
        "Additional scopes and categories are planned for future releases.",
        icon="🚧",
    )
