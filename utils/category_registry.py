from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional

from components.electricity_form import render_electricity_form
from components.purchased_heat_form import render_purchased_heat_form
from components.stationary_combustion_form import render_stationary_combustion_form
from utils.config import (
    build_raw_config,
    build_raw_config_electricity,
    build_raw_config_stationary,
    validate_raw_config,
    validate_raw_config_electricity,
    validate_raw_config_stationary,
)
from utils.generator import generate_json_ground_truth

FormRenderer = Callable[[Optional[str]], Optional[dict]]
RawConfigBuilder = Callable[[dict], dict]
RawConfigValidator = Callable[[dict], list[str]]
GroundTruthBuilder = Callable[[dict], bytes]
DocumentTypeSlugBuilder = Callable[[Optional[str], dict], str]

_EMPTY_GROUND_TRUTH = b"{}"


def _build_heat_raw_config(form_data: dict) -> dict:
    return build_raw_config(form_data, category="purchased_heat_steam_cooling")


def _build_empty_ground_truth(raw_config: dict) -> bytes:
    del raw_config
    return _EMPTY_GROUND_TRUTH


def _default_document_type_slug(document_type: str | None, form_data: dict) -> str:
    del form_data
    return document_type or "document"


def _stationary_document_type_slug(document_type: str | None, form_data: dict) -> str:
    if document_type != "bems":
        return _default_document_type_slug(document_type, form_data)

    report_type = form_data.get("bems_report_type", "")
    if report_type:
        return f"bems_{report_type}"
    return "bems"


@dataclass(frozen=True)
class CategoryWorkflow:
    key: str
    form_renderer: FormRenderer
    raw_config_builder: RawConfigBuilder
    validator: RawConfigValidator
    filename_category_slug: str
    ground_truth_builder: GroundTruthBuilder
    document_type_slug_builder: DocumentTypeSlugBuilder = _default_document_type_slug

    def build_filename_base(self, document_type: str | None, form_data: dict) -> str:
        doc_type_slug = self.document_type_slug_builder(document_type, form_data)
        fp_start = str(form_data.get("fp_start", ""))[:7]
        fp_end = str(form_data.get("fp_end", ""))[:7]
        period_slug = f"{fp_start}_{fp_end}" if fp_start and fp_end and fp_start != fp_end else fp_start
        return f"{self.filename_category_slug}_{doc_type_slug}_{period_slug}"

    def should_zip_export(
        self,
        document_type: str | None,
        output_format: str | None,
        form_data: dict,
    ) -> bool:
        return (
            bool(form_data.get("doc_monthly_zip", False))
            and (document_type or "") in {"utility_bill", "electricity_bill"}
            and output_format in {"PDF", "DOCX"}
        )


CATEGORY_WORKFLOWS: dict[str, CategoryWorkflow] = {
    "purchased_heat_steam_cooling": CategoryWorkflow(
        key="purchased_heat_steam_cooling",
        form_renderer=render_purchased_heat_form,
        raw_config_builder=_build_heat_raw_config,
        validator=validate_raw_config,
        filename_category_slug="heat",
        ground_truth_builder=generate_json_ground_truth,
    ),
    "electricity": CategoryWorkflow(
        key="electricity",
        form_renderer=render_electricity_form,
        raw_config_builder=build_raw_config_electricity,
        validator=validate_raw_config_electricity,
        filename_category_slug="electricity",
        ground_truth_builder=_build_empty_ground_truth,
    ),
    "stationary_combustion": CategoryWorkflow(
        key="stationary_combustion",
        form_renderer=render_stationary_combustion_form,
        raw_config_builder=build_raw_config_stationary,
        validator=validate_raw_config_stationary,
        filename_category_slug="stationary",
        ground_truth_builder=generate_json_ground_truth,
        document_type_slug_builder=_stationary_document_type_slug,
    ),
}


def get_category_workflow(category: str) -> CategoryWorkflow | None:
    return CATEGORY_WORKFLOWS.get(category)
