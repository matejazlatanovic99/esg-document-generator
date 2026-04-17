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
    pdf_generator,
    stationary_combustion_generator,
    xlsx_generator,
)


def _normalize_config(
    raw_config: dict,
    category_module,
    *,
    default_title: str,
    default_subject: str,
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
        },
        "financial_period": financial_period,
        "companies": normalized_companies,
    }
    sections = category_module.build_sections(config)
    return config, sections


def _build_heat_config(raw_config: dict) -> tuple[dict, list[dict]]:
    return _normalize_config(
        raw_config,
        heat_steam_generator,
        default_title="Document",
        default_subject="",
    )


def _build_electricity_config(raw_config: dict) -> tuple[dict, list[dict]]:
    return _normalize_config(
        raw_config,
        electricity_generator,
        default_title="Electricity Consumption Statement",
        default_subject="Scope 2 Electricity",
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
    return "heat"


def _document_type_key(raw_config: dict) -> str:
    document_type = raw_config.get("document_type") or raw_config.get("document", {}).get("type")
    if document_type:
        return str(document_type)
    if _category_key(raw_config) == "electricity":
        return "electricity_bill"
    if _category_key(raw_config) == "stationary_combustion":
        return "fuel_invoice"
    return "utility_bill"


def _should_generate_monthly_zip(raw_config: dict, output_format: str) -> bool:
    if output_format not in {"PDF", "DOCX"}:
        return False
    if _document_type_key(raw_config) not in {"utility_bill", "electricity_bill"}:
        return False
    return bool(raw_config.get("document", {}).get("monthly_zip", False))


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


def _prepare_heat_sections(raw_config: dict) -> tuple[dict, list[dict], set[str]]:
    config, sections = _build_heat_config(raw_config)
    blanked_sections, blank_fields = _apply_blanks(sections)
    if raw_config["document"].get("inject_special_chars", False):
        blanked_sections = _apply_special_chars(blanked_sections)
    return config, blanked_sections, blank_fields


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
    config, sections, _ = _prepare_heat_sections(raw_config)
    if _should_generate_monthly_zip(raw_config, "PDF"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_heat_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            filename = f"heat_utility_bill_{month_slug}.pdf"
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
    if _category_key(raw_config) == "stationary_combustion":
        return stationary_combustion_generator.generate_ground_truth_json(raw_config)

    _, sections = _build_heat_config(raw_config)
    blanked_sections, _ = _apply_blanks(sections)

    output = []
    for section in blanked_sections:
        company_label = section["company"].get("label", "")
        site_label = section["site"].get("label", "")
        for rec in section["records"]:
            entry: dict = {"company_label": company_label, "site_label": site_label}
            entry.update(rec)
            output.append(entry)

    return json.dumps(output, indent=2, default=_json_default).encode("utf-8")


def _generate_heat_supplier_portal_xlsx(raw_config: dict) -> bytes:
    config, sections, xlsx_blank = _prepare_heat_sections(raw_config)
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
    config, sections, docx_blank = _prepare_heat_sections(raw_config)
    if _should_generate_monthly_zip(raw_config, "DOCX"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_heat_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            filename = f"heat_utility_bill_{month_slug}.docx"
            documents.append((
                filename,
                docx_generator.generate_docx(month_config, month_sections, blank_fields=docx_blank, category="heat"),
            ))
        return _build_zip_archive(documents)
    return docx_generator.generate_docx(config, sections, blank_fields=docx_blank, category="heat")


def _generate_heat_supplier_portal_csv(raw_config: dict) -> bytes:
    config, sections, csv_blank = _prepare_heat_sections(raw_config)
    return csv_generator.generate_csv(config, sections, blank_fields=csv_blank, category="heat")


def _generate_electricity_bill_pdf(raw_config: dict) -> bytes:
    config, sections = _build_electricity_config(raw_config)
    if _should_generate_monthly_zip(raw_config, "PDF"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_electricity_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            filename = f"electricity_electricity_bill_{month_slug}.pdf"
            documents.append((filename, _render_electricity_pdf_bytes(month_config, month_sections, filename)))
        return _build_zip_archive(documents)
    return _render_electricity_pdf_bytes(config, sections, "electricity_electricity_bill.pdf")


def _generate_electricity_supplier_portal_xlsx(raw_config: dict) -> bytes:
    config, sections = _build_electricity_config(raw_config)
    include_summary = raw_config.get("xlsx_include_summary", False)
    return xlsx_generator.generate_xlsx(config, sections, include_summary=include_summary, category="electricity")


def _generate_electricity_smart_meter_xlsx(raw_config: dict) -> bytes:
    config, sections = _build_electricity_config(raw_config)
    return xlsx_generator.generate_xlsx(config, sections, category="electricity")


def _generate_electricity_supplier_portal_csv(raw_config: dict) -> bytes:
    config, sections = _build_electricity_config(raw_config)
    return csv_generator.generate_csv(config, sections, category="electricity")


def _generate_electricity_smart_meter_csv(raw_config: dict) -> bytes:
    config, sections = _build_electricity_config(raw_config)
    return csv_generator.generate_csv(config, sections, category="electricity")


def _generate_electricity_bill_docx(raw_config: dict) -> bytes:
    config, sections = _build_electricity_config(raw_config)
    if _should_generate_monthly_zip(raw_config, "DOCX"):
        documents: list[tuple[str, bytes]] = []
        for (year, month), month_sections in sorted(_group_electricity_sections_by_month(sections).items()):
            month_slug = _month_slug(year, month)
            month_config = _month_config(config, year, month)
            filename = f"electricity_electricity_bill_{month_slug}.docx"
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
