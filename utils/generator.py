from __future__ import annotations

import importlib.util
import json
import os
import tempfile
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

_GENERATORS_DIR = Path(__file__).parent.parent / "generators"


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, str(path.resolve()))
    module = importlib.util.module_from_spec(spec)  # type: ignore[arg-type]
    spec.loader.exec_module(module)  # type: ignore[union-attr]
    return module


# Load once at import time so fonts are registered only once.
_pdf_gen = _load_module("pdf_generator", _GENERATORS_DIR / "pdf-generator.py")
_xlsx_gen = _load_module("xlsx_generator", _GENERATORS_DIR / "xlsx-generator.py")
_csv_gen  = _load_module("csv_generator",  _GENERATORS_DIR / "csv-generator.py")
_docx_gen = _load_module("docx_generator", _GENERATORS_DIR / "docx-generator.py")


def _build_normalized_config(raw_config: dict) -> tuple[dict, list[dict]]:
    """Parse dates, normalise companies, build sections. Shared by all generators."""
    fp_raw = raw_config["financial_period"]
    financial_period = {
        "label": fp_raw["label"],
        "start_date": datetime.strptime(fp_raw["start_date"], "%Y-%m-%d").date(),
        "end_date": datetime.strptime(fp_raw["end_date"], "%Y-%m-%d").date(),
    }
    normalized_companies = []
    for idx, company in enumerate(raw_config["companies"]):
        norm = _pdf_gen.normalize_company(company, financial_period, idx)
        norm["_omit"] = company.get("_omit", {})
        raw_sites = company.get("sites", [])
        for sj, norm_site in enumerate(norm["sites"]):
            norm_site["_omit"] = raw_sites[sj].get("_omit", {}) if sj < len(raw_sites) else {}
        normalized_companies.append(norm)

    doc_cfg = raw_config["document"]
    config = {
        "random_seed": int(raw_config.get("random_seed", 42)),
        "document": {
            "title":       doc_cfg.get("title", "Document"),
            "subject":     doc_cfg.get("subject", ""),
            "language":    doc_cfg.get("language", "en"),
            "noise_level": float(doc_cfg.get("noise_level", 1.0)),
        },
        "financial_period": financial_period,
        "companies": normalized_companies,
    }
    sections = _pdf_gen.build_sections(config)
    return config, sections


# Maps form-level site omit keys → record field names they blank in XLSX
_SITE_NUMERIC_OMIT_MAP: dict[str, list[str]] = {
    "capacity_kw":       ["capacity_kw"],
    "capacity_rate":     ["capacity_rate"],
    "base_consumption":  ["consumption"],
    "unit_price_base":   ["unit_price"],
    "start_reading":     ["prev_read", "curr_read"],
}


def _apply_blanks(sections: list[dict]) -> tuple[list[dict], set[str]]:
    """Blank omitted fields in sections for output rendering.

    Text fields are blanked in-place (affects both PDF and XLSX).
    Returns a set of numeric record field names to blank in XLSX only.
    """
    import copy
    sections = copy.deepcopy(sections)
    xlsx_blank: set[str] = set()

    for section in sections:
        co_omit: dict = section["company"].get("_omit", {})
        site_omit: dict = section["site"].get("_omit", {})

        # Company-level display objects
        if co_omit.get("supplier"):
            section["company"]["supplier"] = ""
        if co_omit.get("supplier_address"):
            section["company"]["supplier_address"] = []
        if co_omit.get("label"):
            section["company"]["label"] = ""

        # Site-level address info box
        if site_omit.get("address"):
            section["site"]["customer_address"] = []
        if site_omit.get("label"):
            section["site"]["label"] = ""

        # Billing record text fields
        for rec in section["records"]:
            if co_omit.get("supplier"):    rec["supplier"] = ""
            if co_omit.get("customer"):    rec["customer"] = ""
            if site_omit.get("label"):     rec["site_label"] = ""
            if site_omit.get("city"):      rec["city"] = ""
            if site_omit.get("postcode"):  rec["postcode"] = ""
            if site_omit.get("meter_id"):  rec["meter_id"] = ""

        # Collect numeric fields for XLSX-only blanking
        for form_field, record_fields in _SITE_NUMERIC_OMIT_MAP.items():
            if site_omit.get(form_field):
                xlsx_blank.update(record_fields)

    return sections, xlsx_blank


_SPECIAL_CHARS_SUFFIX = ' & < " £ € \u00a0\u2014\u200f'

_SPECIAL_CHAR_RECORD_FIELDS = [
    "supplier", "customer", "site_label", "city", "postcode",
    "meter_id", "billing_period_label", "invoice_no",
]


def _apply_special_chars(sections: list[dict]) -> list[dict]:
    """Append special characters to every non-empty string record field."""
    import copy
    sections = copy.deepcopy(sections)
    for section in sections:
        for rec in section["records"]:
            for field in _SPECIAL_CHAR_RECORD_FIELDS:
                value = rec.get(field)
                if isinstance(value, str) and value:
                    rec[field] = value + _SPECIAL_CHARS_SUFFIX
    return sections


def generate_pdf_document(raw_config: dict) -> bytes:
    """Generate a PDF from a raw config dict and return the PDF bytes."""
    config, sections = _build_normalized_config(raw_config)

    with tempfile.TemporaryDirectory() as tmpdir:
        bg_dir = os.path.join(tmpdir, "_backgrounds")
        os.makedirs(bg_dir)

        doc_cfg = raw_config["document"]
        pdf_filename = doc_cfg.get("pdf_filename", "output.pdf")
        output_path = os.path.join(tmpdir, pdf_filename)

        config["document"].update({
            "output_dir": tmpdir,
            "pdf_filename": pdf_filename,
            "pdf_path": output_path,
            "background_dir": bg_dir,
        })

        blanked_sections, _ = _apply_blanks(sections)
        if raw_config["document"].get("inject_special_chars", False):
            blanked_sections = _apply_special_chars(blanked_sections)
        noise_level = config["document"].get("noise_level", 1.0)
        _pdf_gen.render_pdf(config, blanked_sections, output_path, noise_level=noise_level)

        with open(output_path, "rb") as fh:
            return fh.read()


def _json_default(obj):
    if isinstance(obj, (date, datetime)):
        return obj.isoformat()
    if isinstance(obj, Decimal):
        return str(obj)
    raise TypeError(f"Object of type {type(obj).__name__} is not JSON serializable")


def generate_json_ground_truth(raw_config: dict) -> bytes:
    """Return a JSON bytes representation of the normalised, blanked sections."""
    _, sections = _build_normalized_config(raw_config)
    blanked_sections, _ = _apply_blanks(sections)

    output = []
    for section in blanked_sections:
        company_label = section["company"].get("label", "")
        site_label = section["site"].get("label", "")
        records = []
        for rec in section["records"]:
            entry: dict = {"company_label": company_label, "site_label": site_label}
            for k, v in rec.items():
                entry[k] = v
            records.append(entry)
        output.extend(records)

    return json.dumps(output, indent=2, default=_json_default).encode("utf-8")


def generate_xlsx_document(raw_config: dict) -> bytes:
    """Generate an XLSX workbook from a raw config dict and return the bytes."""
    config, sections = _build_normalized_config(raw_config)
    blanked_sections, xlsx_blank = _apply_blanks(sections)
    if raw_config["document"].get("inject_special_chars", False):
        blanked_sections = _apply_special_chars(blanked_sections)
    split_by_company = raw_config.get("xlsx_split_by_company", False)
    return _xlsx_gen.generate_xlsx(config, blanked_sections, xlsx_blank, split_by_company=split_by_company)


def generate_docx_document(raw_config: dict) -> bytes:
    """Generate a DOCX billing document from a raw config dict and return the bytes."""
    config, sections = _build_normalized_config(raw_config)
    blanked_sections, docx_blank = _apply_blanks(sections)
    if raw_config["document"].get("inject_special_chars", False):
        blanked_sections = _apply_special_chars(blanked_sections)
    return _docx_gen.generate_docx(config, blanked_sections, blank_fields=docx_blank)


def generate_csv_document(raw_config: dict) -> bytes:
    """Generate a CSV of billing detail rows from a raw config dict and return the bytes."""
    config, sections = _build_normalized_config(raw_config)
    blanked_sections, csv_blank = _apply_blanks(sections)
    if raw_config["document"].get("inject_special_chars", False):
        blanked_sections = _apply_special_chars(blanked_sections)
    return _csv_gen.generate_csv(config, blanked_sections, blank_fields=csv_blank)
