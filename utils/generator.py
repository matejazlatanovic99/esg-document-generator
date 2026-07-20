from __future__ import annotations

import copy
import json
import os
import tempfile
import zipfile
from calendar import monthrange
from datetime import date, datetime
from decimal import Decimal
from io import BytesIO

from generators import (
    csv_generator,
    docx_generator,
    electricity_generator,
    heat_steam_generator,
    mobile_combustion_generator,
    pdf_generator,
    stationary_combustion_generator,
    xlsx_generator,
)
from utils.bad_data import (
    FTYPE_CURRENCY_UNIT,
    FTYPE_DATE_TIME,
    FTYPE_IDENTIFIER,
    FTYPE_NUMERIC,
    FTYPE_TEXT,
    BadDataConfig,
    corrupt_records,
    get_bad_data_config,
)
from utils.distractor_fields import (
    normalize_distractor_settings,
    resolve_distractor_plan,
)
from utils.document_catalog import (
    supports_monthly_zip_export,
    supports_multi_document_export,
)
from utils.layouts import apply_layout_suffix, normalize_layout_settings, resolve_layout_plan

# ── field-type maps for corruption ───────────────────────────────────────────

_HEAT_RECORD_FIELD_TYPES: dict[str, str] = {
    "supplier": FTYPE_TEXT,
    "customer": FTYPE_TEXT,
    "site_label": FTYPE_TEXT,
    "city": FTYPE_TEXT,
    "postcode": FTYPE_IDENTIFIER,
    "period_start": FTYPE_DATE_TIME,
    "period_end": FTYPE_DATE_TIME,
    "meter_id": FTYPE_IDENTIFIER,
    "billing_period_label": FTYPE_DATE_TIME,
    "invoice_no": FTYPE_IDENTIFIER,
    "issue_date": FTYPE_DATE_TIME,
    "due_date": FTYPE_DATE_TIME,
    "prev_read": FTYPE_NUMERIC,
    "curr_read": FTYPE_NUMERIC,
    "consumption": FTYPE_NUMERIC,
    "unit_price": FTYPE_NUMERIC,
    "heat_cost": FTYPE_NUMERIC,
    "capacity_kw": FTYPE_NUMERIC,
    "capacity_rate": FTYPE_NUMERIC,
    "supplier_ef": FTYPE_NUMERIC,
    "capacity_charge": FTYPE_NUMERIC,
    "subtotal": FTYPE_NUMERIC,
    "vat": FTYPE_NUMERIC,
    "total": FTYPE_NUMERIC,
}

_ELECTRICITY_RECORD_FIELD_TYPES: dict[str, str] = {
    "supplier": FTYPE_TEXT,
    "customer": FTYPE_TEXT,
    "site_label": FTYPE_TEXT,
    "city": FTYPE_TEXT,
    "postcode": FTYPE_IDENTIFIER,
    "meter_id": FTYPE_IDENTIFIER,
    "billing_period_label": FTYPE_DATE_TIME,
    "invoice_no": FTYPE_IDENTIFIER,
    "ref_no": FTYPE_IDENTIFIER,
    "period_start": FTYPE_DATE_TIME,
    "period_end": FTYPE_DATE_TIME,
    "start_reading": FTYPE_NUMERIC,
    "end_reading": FTYPE_NUMERIC,
    "total_quantity": FTYPE_NUMERIC,
    "total_cost": FTYPE_NUMERIC,
    "supplier_ef": FTYPE_NUMERIC,
    "emissions_kg": FTYPE_NUMERIC,
    "emissions_t": FTYPE_NUMERIC,
    "unit": FTYPE_CURRENCY_UNIT,
}


def _apply_invalid_data_to_sections(
    sections: list[dict],
    field_types: dict[str, str],
    cfg: BadDataConfig,
    artifact_prefix: str,
) -> list[dict]:
    """Return new sections list with records corrupted per cfg.

    Heat sections: section["records"] is a list of record dicts.
    Electricity sections: section["site"] IS the record dict (no "records" key).
    """
    if not cfg.enabled:
        return sections
    result = []
    # Detect electricity-style sections (no "records" key, has "site" key directly)
    if sections and "records" not in sections[0]:
        # Electricity: corrupt the site dicts directly
        corrupted_sites = corrupt_records(
            [s["site"] for s in sections],
            artifact_prefix,
            field_types,
            cfg,
        )
        for section, corrupted_site in zip(sections, corrupted_sites):
            new_section = dict(section)
            new_section["site"] = corrupted_site
            result.append(new_section)
    else:
        for s_idx, section in enumerate(sections):
            new_section = dict(section)
            new_section["records"] = corrupt_records(
                section["records"],
                f"{artifact_prefix}:s{s_idx}",
                field_types,
                cfg,
            )
            result.append(new_section)
    return result


def _apply_resolved_document_features(
    config: dict,
    raw_config: dict,
    *,
    output_format: str,
    artifact_key: str = "root",
    document_type: str | None = None,
) -> None:
    doc_cfg = raw_config.get("document", {})
    random_seed = int(raw_config.get("random_seed", 42))
    category = _category_key(raw_config)
    resolved_document_type = document_type or _document_type_key(raw_config)
    context = _layout_context(raw_config)

    config["document"]["layout"] = normalize_layout_settings(doc_cfg, random_seed)
    config["document"]["distractor_fields"] = normalize_distractor_settings(doc_cfg)

    layout_plan = resolve_layout_plan(
        doc_cfg,
        random_seed=random_seed,
        category=category,
        document_type=resolved_document_type,
        output_format=output_format,
        artifact_key=artifact_key,
        context=context,
    )
    config["document"]["layout_slug"] = layout_plan.get("slug")
    config["document"]["layout_plan"] = layout_plan
    config["document"]["distractor_plan"] = resolve_distractor_plan(
        doc_cfg,
        random_seed=random_seed,
        category=category,
        document_type=resolved_document_type,
        output_format=output_format,
        artifact_key=artifact_key,
        context=context,
    )


def _exclude_site_field_when_omitted(config: dict, sections: list[dict]) -> None:
    if not sections:
        return

    if not all(section.get("site", {}).get("_omit", {}).get("label", False) for section in sections):
        return

    plan = config.get("document", {}).get("layout_plan")
    if not isinstance(plan, dict):
        return

    excluded_fields = list(plan.get("excluded_fields") or [])
    if "site" not in excluded_fields:
        excluded_fields.append("site")
    plan["excluded_fields"] = excluded_fields
    plan["column_order"] = [field_id for field_id in plan.get("column_order", []) if field_id != "site"]


def _normalize_config(
    raw_config: dict,
    category_module,
    *,
    default_title: str,
    default_subject: str,
    output_format: str | None = None,
    artifact_key: str = "root",
) -> tuple[dict, list[dict]]:
    fp_raw = raw_config["financial_period"]
    financial_period = {
        "label": fp_raw["label"],
        "start_date": datetime.strptime(fp_raw["start_date"], "%Y-%m-%d").date(),
        "end_date": datetime.strptime(fp_raw["end_date"], "%Y-%m-%d").date(),
    }

    normalized_companies = []
    for idx, company in enumerate(raw_config["companies"]):
        normalized_company = category_module.normalize_company(company, financial_period, idx)
        normalized_company["_omit"] = company.get("_omit", {})
        raw_sites = company.get("sites", [])
        for site_idx, normalized_site in enumerate(normalized_company["sites"]):
            normalized_site["_omit"] = raw_sites[site_idx].get("_omit", {}) if site_idx < len(raw_sites) else {}
        normalized_companies.append(normalized_company)

    doc_cfg = raw_config["document"]
    document_type = raw_config.get("document_type", doc_cfg.get("type", ""))
    config = {
        "random_seed": int(raw_config.get("random_seed", 42)),
        "document_type": document_type,
        "document": {
            "type": document_type,
            "title": doc_cfg.get("title", default_title),
            "subject": doc_cfg.get("subject", default_subject),
            "language": doc_cfg.get("language", "en"),
            "noise_level": float(doc_cfg.get("noise_level", 0.0)),
            "monthly_zip": bool(doc_cfg.get("monthly_zip", False)),
            "smart_meter_data_granularity": doc_cfg.get("smart_meter_data_granularity", "monthly"),
            "smart_meter_interval_minutes": int(doc_cfg.get("smart_meter_interval_minutes", 30)),
            "smart_meter_interval_value_mode": doc_cfg.get("smart_meter_interval_value_mode", "consumption_diff"),
            "smart_meter_timestamp_format": doc_cfg.get("smart_meter_timestamp_format", "iso_8601_utc"),
            "invalid_data": doc_cfg.get("invalid_data", {"enabled": False, "preset": "mixed", "rate": 0.15}),
        },
        "financial_period": financial_period,
        "companies": normalized_companies,
    }
    _apply_resolved_document_features(
        config,
        raw_config,
        output_format=output_format or "GENERIC",
        artifact_key=artifact_key,
        document_type=document_type,
    )
    sections = category_module.build_sections(config)
    _exclude_site_field_when_omitted(config, sections)
    return config, sections


def _build_heat_config(
    raw_config: dict,
    *,
    output_format: str | None = None,
    artifact_key: str = "root",
) -> tuple[dict, list[dict]]:
    return _normalize_config(
        raw_config,
        heat_steam_generator,
        default_title="Document",
        default_subject="",
        output_format=output_format,
        artifact_key=artifact_key,
    )


def _build_electricity_config(
    raw_config: dict,
    *,
    output_format: str | None = None,
    artifact_key: str = "root",
) -> tuple[dict, list[dict]]:
    return _normalize_config(
        raw_config,
        electricity_generator,
        default_title="Electricity Consumption Statement",
        default_subject="Scope 2 Electricity",
        output_format=output_format,
        artifact_key=artifact_key,
    )


_SITE_NUMERIC_OMIT_MAP: dict[str, list[str]] = {
    "capacity_kw": ["capacity_kw"],
    "capacity_rate": ["capacity_rate"],
    "supplier_ef": ["supplier_ef"],
    "base_consumption": ["consumption"],
    "unit_price_base": ["unit_price"],
    "start_reading": ["prev_read", "curr_read"],
}


def _apply_blanks(sections: list[dict]) -> tuple[list[dict], set[str]]:
    sections = copy.deepcopy(sections)
    xlsx_blank: set[str] = set()

    for section in sections:
        site_omit: dict = section["site"].get("_omit", {})
        if site_omit.get("label"):
            section["site"]["label"] = ""

        for rec in section["records"]:
            if site_omit.get("label"):
                rec["site_label"] = ""
            if site_omit.get("city"):
                rec["city"] = ""
            if site_omit.get("postcode"):
                rec["postcode"] = ""

        for form_field, record_fields in _SITE_NUMERIC_OMIT_MAP.items():
            if site_omit.get(form_field):
                xlsx_blank.update(record_fields)

    return sections, xlsx_blank


_SPECIAL_CHARS_SUFFIX = ' & < " £ € \u00a0\u2014\u200f'

_SPECIAL_CHAR_RECORD_FIELDS = [
    "supplier",
    "customer",
    "site_label",
    "city",
    "postcode",
    "meter_id",
    "billing_period_label",
    "invoice_no",
]


def _apply_special_chars(sections: list[dict]) -> list[dict]:
    sections = copy.deepcopy(sections)
    for section in sections:
        for rec in section["records"]:
            for field in _SPECIAL_CHAR_RECORD_FIELDS:
                value = rec.get(field)
                if isinstance(value, str) and value:
                    rec[field] = value + _SPECIAL_CHARS_SUFFIX
    return sections


def _category_key(raw_config: dict) -> str:
    category = raw_config.get("_category")
    if category == "electricity":
        return "electricity"
    if category == "stationary_combustion":
        return "stationary_combustion"
    if category == "mobile_combustion":
        return "mobile_combustion"
    return "heat"


def _document_type_key(raw_config: dict) -> str:
    document_type = raw_config.get("document_type") or raw_config.get("document", {}).get("type")
    if document_type:
        return str(document_type)
    if _category_key(raw_config) == "electricity":
        return "electricity_bill"
    if _category_key(raw_config) in {"stationary_combustion", "mobile_combustion"}:
        return "fuel_invoice"
    return "utility_bill"


def _document_slug_for_filename(raw_config: dict) -> str:
    document_type = _document_type_key(raw_config)
    if _category_key(raw_config) == "stationary_combustion" and document_type == "bems":
        report_type = raw_config.get("document", {}).get("bems_report_type")
        if report_type:
            return f"bems_{report_type}"
    return document_type


def _date_slug_value(value) -> str:
    text = str(value or "")
    return text[:7]


def _layout_context(raw_config: dict) -> dict:
    doc_cfg = raw_config.get("document", {})
    return {
        "language": doc_cfg.get("language", "en"),
        "period_label": raw_config.get("financial_period", {}).get("label", ""),
        "include_summary": bool(raw_config.get("xlsx_include_summary", False)),
        "split_by_company": bool(raw_config.get("xlsx_split_by_company", False)),
        "company_labels": [company.get("label", f"Company {idx + 1}") for idx, company in enumerate(raw_config.get("companies", []))],
        "smart_meter_mode": doc_cfg.get("smart_meter_data_granularity", "monthly"),
        "smart_meter_value_mode": doc_cfg.get("smart_meter_interval_value_mode", "consumption_diff"),
    }


def build_document_filename_base(raw_config: dict, *, output_format: str | None = None, artifact_key: str = "root") -> str:
    category_slug = {
        "electricity": "electricity",
        "stationary_combustion": "stationary",
        "mobile_combustion": "mobile",
    }.get(_category_key(raw_config), "heat")
    fp = raw_config.get("financial_period", {})
    fp_start = _date_slug_value(fp.get("start_date") or fp.get("start"))
    fp_end = _date_slug_value(fp.get("end_date") or fp.get("end"))
    period_slug = f"{fp_start}_{fp_end}" if fp_start and fp_end and fp_start != fp_end else fp_start
    stem = f"{category_slug}_{_document_slug_for_filename(raw_config)}_{period_slug}"

    if not output_format:
        return stem

    layout_plan = resolve_layout_plan(
        raw_config.get("document", {}),
        random_seed=int(raw_config.get("random_seed", 42)),
        category=_category_key(raw_config),
        document_type=_document_type_key(raw_config),
        output_format=output_format,
        artifact_key=artifact_key,
        context=_layout_context(raw_config),
    )
    return apply_layout_suffix(stem, layout_plan.get("slug"))


def build_document_download_filename(raw_config: dict, output_format: str) -> str:
    zip_export = _should_generate_monthly_zip(raw_config, output_format) or _should_generate_multi_document_zip(
        raw_config, output_format
    )
    extension = ".zip" if zip_export else f".{output_format.lower()}"
    artifact_key = "archive" if zip_export else "root"
    return build_document_filename_base(raw_config, output_format=output_format, artifact_key=artifact_key) + extension


def _should_generate_monthly_zip(raw_config: dict, output_format: str) -> bool:
    return bool(raw_config.get("document", {}).get("monthly_zip", False)) and supports_monthly_zip_export(
        _category_key(raw_config),
        _document_type_key(raw_config),
        output_format,
    )


def _should_generate_multi_document_zip(raw_config: dict, output_format: str) -> bool:
    return int(raw_config.get("document", {}).get("doc_count", 1) or 1) > 1 and supports_multi_document_export(
        _category_key(raw_config),
        _document_type_key(raw_config),
        output_format,
    )


def _month_slug(year: int, month: int) -> str:
    return f"{year:04d}-{month:02d}"


def _month_config(config: dict, year: int, month: int) -> dict:
    month_start = date(year, month, 1)
    month_end = date(year, month, monthrange(year, month)[1])
    month_config = copy.deepcopy(config)
    month_config["financial_period"] = {
        "label": month_start.strftime("%B %Y"),
        "start_date": month_start,
        "end_date": month_end,
    }
    return month_config


def _group_heat_sections_by_month(sections: list[dict]) -> dict[tuple[int, int], list[dict]]:
    grouped: dict[tuple[int, int], list[dict]] = {}
    for section in sections:
        company = section["company"]
        site = section["site"]
        for record in section["records"]:
            key = (record["period_start"].year, record["period_start"].month)
            grouped.setdefault(key, []).append({
                "company": company,
                "site": site,
                "records": [record],
            })
    return grouped


def _group_electricity_sections_by_month(sections: list[dict]) -> dict[tuple[int, int], list[dict]]:
    grouped: dict[tuple[int, int], list[dict]] = {}
    for section in sections:
        site = section["site"]
        key = (site["period_start"].year, site["period_start"].month)
        grouped.setdefault(key, []).append(section)
    return grouped


def _build_zip_archive(documents: list[tuple[str, bytes]]) -> bytes:
    archive_buffer = BytesIO()
    with zipfile.ZipFile(archive_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for filename, content in documents:
            archive.writestr(filename, content)
    return archive_buffer.getvalue()


def _prepare_heat_sections(
    raw_config: dict,
    *,
    output_format: str | None = None,
    artifact_key: str = "root",
) -> tuple[dict, list[dict], set[str]]:
    config, sections = _build_heat_config(raw_config, output_format=output_format, artifact_key=artifact_key)
    blanked_sections, blank_fields = _apply_blanks(sections)
    if raw_config["document"].get("inject_special_chars", False):
        blanked_sections = _apply_special_chars(blanked_sections)
    bad_cfg = get_bad_data_config(raw_config)
    blanked_sections = _apply_invalid_data_to_sections(
        blanked_sections, _HEAT_RECORD_FIELD_TYPES, bad_cfg, "heat"
    )
    return config, blanked_sections, blank_fields


def _prepare_electricity_sections(
    raw_config: dict,
    *,
    output_format: str | None = None,
    artifact_key: str = "root",
) -> tuple[dict, list[dict]]:
    config, sections = _build_electricity_config(raw_config, output_format=output_format, artifact_key=artifact_key)
    bad_cfg = get_bad_data_config(raw_config)
    sections = _apply_invalid_data_to_sections(
        sections, _ELECTRICITY_RECORD_FIELD_TYPES, bad_cfg, "electricity"
    )
    return config, sections


def _prepare_pdf_output(config: dict, default_filename: str) -> str:
    tmpdir = config["document"]["output_dir"]
    bg_dir = os.path.join(tmpdir, "_backgrounds")
    os.makedirs(bg_dir)

    pdf_filename = config["document"].get("pdf_filename", default_filename)
    output_path = os.path.join(tmpdir, pdf_filename)
    config["document"].update({
        "output_dir": tmpdir,
        "pdf_filename": pdf_filename,
        "pdf_path": output_path,
        "background_dir": bg_dir,
    })
    return output_path


def _render_heat_pdf_bytes(config: dict, sections: list[dict], default_filename: str = "output.pdf") -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        config = copy.deepcopy(config)
        config["document"]["output_dir"] = tmpdir
        output_path = _prepare_pdf_output(config, default_filename)
        noise_level = config["document"].get("noise_level", 0.0)
        pdf_generator.render_pdf(config, sections, output_path, category="heat", noise_level=noise_level)
        with open(output_path, "rb") as fh:
            return fh.read()


def _render_electricity_pdf_bytes(
    config: dict,
    sections: list[dict],
    default_filename: str = "electricity_statement.pdf",
) -> bytes:
    with tempfile.TemporaryDirectory() as tmpdir:
        config = copy.deepcopy(config)
        config["document"]["output_dir"] = tmpdir
        output_path = _prepare_pdf_output(config, default_filename)
        noise_level = float(config["document"].get("noise_level", 0.0))
        pdf_generator.render_pdf(config, sections, output_path, category="electricity", noise_level=noise_level)
        with open(output_path, "rb") as fh:
            return fh.read()


def _generate_heat_utility_bill_pdf(raw_config: dict) -> bytes:
    config, sections, _ = _prepare_heat_sections(raw_config, output_format="PDF")
    if _should_generate_monthly_zip(raw_config, "PDF"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_heat_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            month_raw_config = copy.deepcopy(raw_config)
            month_raw_config["financial_period"] = {
                "label": month_config["financial_period"]["label"],
                "start_date": month_config["financial_period"]["start_date"],
                "end_date": month_config["financial_period"]["end_date"],
            }
            _apply_resolved_document_features(month_config, month_raw_config, output_format="PDF", artifact_key=month_slug)
            filename = build_document_filename_base(month_raw_config, output_format="PDF", artifact_key=month_slug) + ".pdf"
            documents.append((filename, _render_heat_pdf_bytes(month_config, month_sections, filename)))
        return _build_zip_archive(documents)
    return _render_heat_pdf_bytes(config, sections, "heat_utility_bill.pdf")


def _json_default(obj):
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    if isinstance(obj, Decimal):
        return str(obj)
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")


def generate_json_ground_truth(raw_config: dict) -> bytes:
    category = _category_key(raw_config)
    if category == "stationary_combustion":
        return stationary_combustion_generator.generate_ground_truth_json(raw_config)
    if category == "mobile_combustion":
        return mobile_combustion_generator.generate_ground_truth_json(raw_config)

    output = []
    if category == "electricity":
        document_type = str(raw_config.get("document_type") or raw_config.get("document", {}).get("type") or "")
        if document_type == "smart_meter_data":
            from generators.electricity_generator import build_smart_meter_rows

            config, sections = _build_electricity_config(raw_config)
            output = build_smart_meter_rows(config, sections)
        else:
            _, sections = _prepare_electricity_sections(raw_config)
            for section in sections:
                company_label = section["company"].get("label", "")
                site = section.get("site", {})
                entry = {
                    "scope": "Scope 2",
                    "category": "electricity",
                    "company_label": company_label,
                    "site_label": site.get("label", ""),
                }
                entry.update(site)
                output.append(entry)
    else:
        _, sections, _ = _prepare_heat_sections(raw_config)
        for section in sections:
            company_label = section["company"].get("label", "")
            site_label = section["site"].get("label", "")
            for rec in section["records"]:
                entry: dict = {
                    "scope": "Scope 2",
                    "category": "purchased_heat_steam_cooling",
                    "company_label": company_label,
                    "site_label": site_label,
                }
                entry.update(rec)
                output.append(entry)

    return json.dumps(output, indent=2, default=_json_default).encode("utf-8")


def _generate_heat_supplier_portal_xlsx(raw_config: dict) -> bytes:
    config, sections, xlsx_blank = _prepare_heat_sections(raw_config, output_format="XLSX")
    split_by_company = raw_config.get("xlsx_split_by_company", False)
    include_summary = raw_config.get("xlsx_include_summary", False)
    return xlsx_generator.generate_xlsx(
        config,
        sections,
        blank_fields=xlsx_blank,
        split_by_company=split_by_company,
        include_summary=include_summary,
        category="heat",
    )


def _generate_heat_utility_bill_docx(raw_config: dict) -> bytes:
    config, sections, docx_blank = _prepare_heat_sections(raw_config, output_format="DOCX")
    if _should_generate_monthly_zip(raw_config, "DOCX"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_heat_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            month_raw_config = copy.deepcopy(raw_config)
            month_raw_config["financial_period"] = {
                "label": month_config["financial_period"]["label"],
                "start_date": month_config["financial_period"]["start_date"],
                "end_date": month_config["financial_period"]["end_date"],
            }
            _apply_resolved_document_features(month_config, month_raw_config, output_format="DOCX", artifact_key=month_slug)
            filename = build_document_filename_base(month_raw_config, output_format="DOCX", artifact_key=month_slug) + ".docx"
            documents.append((
                filename,
                docx_generator.generate_docx(month_config, month_sections, blank_fields=docx_blank, category="heat"),
            ))
        return _build_zip_archive(documents)
    return docx_generator.generate_docx(config, sections, blank_fields=docx_blank, category="heat")


def _generate_heat_supplier_portal_csv(raw_config: dict) -> bytes:
    config, sections, csv_blank = _prepare_heat_sections(raw_config, output_format="CSV")
    return csv_generator.generate_csv(config, sections, blank_fields=csv_blank, category="heat")


def _generate_electricity_bill_pdf(raw_config: dict) -> bytes:
    config, sections = _prepare_electricity_sections(raw_config, output_format="PDF")
    if _should_generate_monthly_zip(raw_config, "PDF"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_electricity_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            month_raw_config = copy.deepcopy(raw_config)
            month_raw_config["financial_period"] = {
                "label": month_config["financial_period"]["label"],
                "start_date": month_config["financial_period"]["start_date"],
                "end_date": month_config["financial_period"]["end_date"],
            }
            _apply_resolved_document_features(month_config, month_raw_config, output_format="PDF", artifact_key=month_slug)
            filename = build_document_filename_base(month_raw_config, output_format="PDF", artifact_key=month_slug) + ".pdf"
            documents.append((filename, _render_electricity_pdf_bytes(month_config, month_sections, filename)))
        return _build_zip_archive(documents)
    return _render_electricity_pdf_bytes(config, sections, "electricity_electricity_bill.pdf")


def _generate_electricity_supplier_portal_xlsx(raw_config: dict) -> bytes:
    config, sections = _prepare_electricity_sections(raw_config, output_format="XLSX")
    include_summary = raw_config.get("xlsx_include_summary", False)
    return xlsx_generator.generate_xlsx(config, sections, include_summary=include_summary, category="electricity")


def _generate_electricity_smart_meter_xlsx(raw_config: dict) -> bytes:
    # Smart meter rows are derived/computed from period readings; skip corruption
    # to avoid breaking interval computation (same approach as BEMS time series).
    config, sections = _build_electricity_config(raw_config, output_format="XLSX")
    return xlsx_generator.generate_xlsx(config, sections, category="electricity")


def _generate_electricity_supplier_portal_csv(raw_config: dict) -> bytes:
    config, sections = _prepare_electricity_sections(raw_config, output_format="CSV")
    return csv_generator.generate_csv(config, sections, category="electricity")


def _generate_electricity_smart_meter_csv(raw_config: dict) -> bytes:
    # Smart meter rows are derived/computed from period readings; skip corruption.
    config, sections = _build_electricity_config(raw_config, output_format="CSV")
    return csv_generator.generate_csv(config, sections, category="electricity")


def _generate_electricity_bill_docx(raw_config: dict) -> bytes:
    config, sections = _prepare_electricity_sections(raw_config, output_format="DOCX")
    if _should_generate_monthly_zip(raw_config, "DOCX"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_electricity_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            month_raw_config = copy.deepcopy(raw_config)
            month_raw_config["financial_period"] = {
                "label": month_config["financial_period"]["label"],
                "start_date": month_config["financial_period"]["start_date"],
                "end_date": month_config["financial_period"]["end_date"],
            }
            _apply_resolved_document_features(month_config, month_raw_config, output_format="DOCX", artifact_key=month_slug)
            filename = build_document_filename_base(month_raw_config, output_format="DOCX", artifact_key=month_slug) + ".docx"
            documents.append((filename, docx_generator.generate_docx(month_config, month_sections, category="electricity")))
        return _build_zip_archive(documents)
    return docx_generator.generate_docx(config, sections, category="electricity")


def _generate_stationary_fuel_invoice_pdf(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_fuel_invoice_pdf(raw_config)


def _generate_stationary_fuel_invoice_docx(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_fuel_invoice_docx(raw_config)


def _generate_stationary_delivery_note_pdf(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_delivery_note_pdf(raw_config)


def _generate_stationary_delivery_note_docx(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_delivery_note_docx(raw_config)


def _generate_stationary_fuel_card_pdf(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_fuel_card_pdf(raw_config)


def _generate_stationary_fuel_card_docx(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_fuel_card_docx(raw_config)


def _generate_stationary_fuel_card_xlsx(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_fuel_card_xlsx(raw_config)


def _generate_stationary_fuel_card_csv(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_fuel_card_csv(raw_config)


def _generate_stationary_generator_log_xlsx(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_generator_log_xlsx(raw_config)


def _generate_stationary_generator_log_csv(raw_config: dict) -> bytes:
    return stationary_combustion_generator.generate_generator_log_csv(raw_config)


def _generate_stationary_bems_pdf(raw_config: dict) -> bytes:
    report_type = raw_config.get("document", {}).get("bems_report_type", "equipment_trend_report")
    if report_type == "time_series_trend_export":
        return stationary_combustion_generator.generate_bems_time_series_pdf(raw_config)
    return stationary_combustion_generator.generate_bems_equipment_report_pdf(raw_config)


def _generate_stationary_bems_xlsx(raw_config: dict) -> bytes:
    report_type = raw_config.get("document", {}).get("bems_report_type", "equipment_trend_report")
    if report_type == "time_series_trend_export":
        return stationary_combustion_generator.generate_bems_time_series_xlsx(raw_config)
    return stationary_combustion_generator.generate_bems_equipment_report_xlsx(raw_config)


def _generate_stationary_bems_csv(raw_config: dict) -> bytes:
    report_type = raw_config.get("document", {}).get("bems_report_type", "equipment_trend_report")
    if report_type == "time_series_trend_export":
        return stationary_combustion_generator.generate_bems_time_series_csv(raw_config)
    return stationary_combustion_generator.generate_bems_equipment_report_csv(raw_config)


def _generate_stationary_bems_docx(raw_config: dict) -> bytes:
    report_type = raw_config.get("document", {}).get("bems_report_type", "equipment_trend_report")
    if report_type == "time_series_trend_export":
        return stationary_combustion_generator.generate_bems_time_series_docx(raw_config)
    return stationary_combustion_generator.generate_bems_equipment_report_docx(raw_config)


def _generate_mobile_fuel_invoice_pdf(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_fuel_invoice_pdf(raw_config)


def _generate_mobile_fuel_invoice_docx(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_fuel_invoice_docx(raw_config)


def _generate_mobile_fuel_card_statement_pdf(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_fuel_card_statement_pdf(raw_config)


def _generate_mobile_fuel_card_statement_docx(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_fuel_card_statement_docx(raw_config)


def _generate_mobile_fuel_card_statement_xlsx(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_fuel_card_statement_xlsx(raw_config)


def _generate_mobile_fuel_card_statement_csv(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_fuel_card_statement_csv(raw_config)


def _generate_mobile_telematics_fuel_xlsx(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_telematics_fuel_xlsx(raw_config)


def _generate_mobile_telematics_fuel_csv(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_telematics_fuel_csv(raw_config)


def _generate_mobile_telematics_trips_xlsx(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_telematics_trips_xlsx(raw_config)


def _generate_mobile_telematics_trips_csv(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_telematics_trips_csv(raw_config)


def _generate_mobile_telematics_odometer_xlsx(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_telematics_odometer_xlsx(raw_config)


def _generate_mobile_telematics_odometer_csv(raw_config: dict) -> bytes:
    return mobile_combustion_generator.generate_telematics_odometer_csv(raw_config)


_DOCUMENT_GENERATOR_DISPATCH = {
    ("heat", "utility_bill", "PDF"): _generate_heat_utility_bill_pdf,
    ("heat", "utility_bill", "DOCX"): _generate_heat_utility_bill_docx,
    ("heat", "supplier_portal_data", "XLSX"): _generate_heat_supplier_portal_xlsx,
    ("heat", "supplier_portal_data", "CSV"): _generate_heat_supplier_portal_csv,
    ("electricity", "electricity_bill", "PDF"): _generate_electricity_bill_pdf,
    ("electricity", "electricity_bill", "DOCX"): _generate_electricity_bill_docx,
    ("electricity", "supplier_portal_data", "XLSX"): _generate_electricity_supplier_portal_xlsx,
    ("electricity", "supplier_portal_data", "CSV"): _generate_electricity_supplier_portal_csv,
    ("electricity", "smart_meter_data", "XLSX"): _generate_electricity_smart_meter_xlsx,
    ("electricity", "smart_meter_data", "CSV"): _generate_electricity_smart_meter_csv,
    ("stationary_combustion", "fuel_invoice", "PDF"): _generate_stationary_fuel_invoice_pdf,
    ("stationary_combustion", "fuel_invoice", "DOCX"): _generate_stationary_fuel_invoice_docx,
    ("stationary_combustion", "delivery_note", "PDF"): _generate_stationary_delivery_note_pdf,
    ("stationary_combustion", "delivery_note", "DOCX"): _generate_stationary_delivery_note_docx,
    ("stationary_combustion", "fuel_card", "PDF"): _generate_stationary_fuel_card_pdf,
    ("stationary_combustion", "fuel_card", "DOCX"): _generate_stationary_fuel_card_docx,
    ("stationary_combustion", "fuel_card", "XLSX"): _generate_stationary_fuel_card_xlsx,
    ("stationary_combustion", "fuel_card", "CSV"): _generate_stationary_fuel_card_csv,
    ("stationary_combustion", "generator_log", "XLSX"): _generate_stationary_generator_log_xlsx,
    ("stationary_combustion", "generator_log", "CSV"): _generate_stationary_generator_log_csv,
    ("stationary_combustion", "bems", "PDF"): _generate_stationary_bems_pdf,
    ("stationary_combustion", "bems", "DOCX"): _generate_stationary_bems_docx,
    ("stationary_combustion", "bems", "XLSX"): _generate_stationary_bems_xlsx,
    ("stationary_combustion", "bems", "CSV"): _generate_stationary_bems_csv,
    ("mobile_combustion", "fuel_invoice", "PDF"): _generate_mobile_fuel_invoice_pdf,
    ("mobile_combustion", "fuel_invoice", "DOCX"): _generate_mobile_fuel_invoice_docx,
    ("mobile_combustion", "fuel_card_statement", "PDF"): _generate_mobile_fuel_card_statement_pdf,
    ("mobile_combustion", "fuel_card_statement", "DOCX"): _generate_mobile_fuel_card_statement_docx,
    ("mobile_combustion", "fuel_card_statement", "XLSX"): _generate_mobile_fuel_card_statement_xlsx,
    ("mobile_combustion", "fuel_card_statement", "CSV"): _generate_mobile_fuel_card_statement_csv,
    ("mobile_combustion", "telematics_fuel", "XLSX"): _generate_mobile_telematics_fuel_xlsx,
    ("mobile_combustion", "telematics_fuel", "CSV"): _generate_mobile_telematics_fuel_csv,
    ("mobile_combustion", "telematics_trips", "XLSX"): _generate_mobile_telematics_trips_xlsx,
    ("mobile_combustion", "telematics_trips", "CSV"): _generate_mobile_telematics_trips_csv,
    ("mobile_combustion", "telematics_odometer", "XLSX"): _generate_mobile_telematics_odometer_xlsx,
    ("mobile_combustion", "telematics_odometer", "CSV"): _generate_mobile_telematics_odometer_csv,
}


def generate_document_bytes(raw_config: dict, output_format: str) -> bytes:
    category = _category_key(raw_config)
    document_type = _document_type_key(raw_config)
    dispatch_key = (category, document_type, output_format)
    generator = _DOCUMENT_GENERATOR_DISPATCH.get(dispatch_key)
    if generator is None:
        raise NotImplementedError(
            f"Document type '{document_type}' for category '{category}' does not support format '{output_format}'."
        )
    return generator(raw_config)
