from __future__ import annotations


_CATEGORY_ALIASES: dict[str, str] = {
    "heat": "purchased_heat_steam_cooling",
}


OUTPUT_FORMATS: dict[str, dict] = {
    "PDF": {"mime": "application/pdf", "ext": ".pdf", "implemented": True},
    "DOCX": {
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "ext": ".docx",
        "implemented": True,
    },
    "XLSX": {
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "ext": ".xlsx",
        "implemented": True,
    },
    "CSV": {"mime": "text/csv", "ext": ".csv", "implemented": True},
}


DOCUMENT_TYPES: dict[str, dict[str, dict]] = {
    "stationary_combustion": {
        "fuel_invoice": {
            "label": "Fuel Invoice",
            "formats": ["PDF", "DOCX"],
            "implemented": True,
            "default_title": "Stationary Combustion Fuel Invoice",
            "default_subject": "Scope 1 stationary combustion fuel purchase",
            "company_currency_required": True,
        },
        "generator_log": {
            "label": "Generator Log",
            "formats": ["XLSX", "CSV"],
            "implemented": True,
            "default_title": "Generator Operation Log",
            "default_subject": "Scope 1 stationary combustion generator operations",
        },
        "bems": {
            "label": "Building Energy Management System (BEMS)",
            "formats": ["PDF", "DOCX", "XLSX", "CSV"],
            "implemented": True,
            "default_title": "BEMS Fuel Consumption Summary",
            "default_subject": "Scope 1 stationary combustion BEMS export",
        },
        "delivery_note": {
            "label": "Delivery Note",
            "formats": ["PDF", "DOCX"],
            "implemented": True,
            "default_title": "Fuel Delivery Note",
            "default_subject": "Scope 1 stationary combustion fuel delivery",
        },
        "fuel_card": {
            "label": "Fuel Card Statement",
            "formats": ["PDF", "DOCX", "XLSX", "CSV"],
            "implemented": True,
            "default_title": "Fuel Card Statement",
            "default_subject": "Scope 1 stationary combustion fuel card transactions",
            "company_currency_required": True,
        },
    },
    "purchased_heat_steam_cooling": {
        "utility_bill": {
            "label": "Utility Bill",
            "formats": ["PDF", "DOCX"],
            "implemented": True,
            "default_title": "District Heating Billing Statement",
            "default_subject": "Purchased heat billing statements",
            "monthly_zip_formats": ["PDF", "DOCX"],
        },
        "supplier_portal_data": {
            "label": "Supplier Portal Data",
            "formats": ["XLSX", "CSV"],
            "implemented": True,
            "default_title": "District Heating Supplier Portal Data Export",
            "default_subject": "Purchased heat supplier portal data",
        },
    },
    "electricity": {
        "electricity_bill": {
            "label": "Electricity Bill",
            "formats": ["PDF", "DOCX"],
            "implemented": True,
            "default_title": "Electricity Consumption Statement",
            "default_subject": "Scope 2 purchased electricity",
            "monthly_zip_formats": ["PDF", "DOCX"],
        },
        "smart_meter_data": {
            "label": "Smart Meter Data",
            "formats": ["XLSX", "CSV"],
            "implemented": True,
            "default_title": "Electricity Smart Meter Data Export",
            "default_subject": "Scope 2 smart meter data",
        },
        "supplier_portal_data": {
            "label": "Supplier Portal Data",
            "formats": ["XLSX", "CSV"],
            "implemented": True,
            "default_title": "Electricity Supplier Portal Data Export",
            "default_subject": "Scope 2 supplier portal data",
        },
    },
}


def _normalize_category_key(category_key: str) -> str:
    return _CATEGORY_ALIASES.get(category_key, category_key)


def get_document_type_options(category_key: str) -> dict[str, dict]:
    return DOCUMENT_TYPES.get(_normalize_category_key(category_key), {})


def get_document_type_config(category_key: str, document_type_key: str) -> dict:
    options = get_document_type_options(category_key)
    return options.get(document_type_key, {})


def get_default_document_type(category_key: str) -> str | None:
    for key, cfg in get_document_type_options(category_key).items():
        if cfg.get("implemented", False):
            return key
    return next(iter(get_document_type_options(category_key)), None)


def get_allowed_formats(category_key: str, document_type_key: str) -> list[str]:
    config = get_document_type_config(category_key, document_type_key)
    return list(config.get("formats", []))


def get_default_format(category_key: str, document_type_key: str) -> str | None:
    for fmt in get_allowed_formats(category_key, document_type_key):
        if OUTPUT_FORMATS.get(fmt, {}).get("implemented", False):
            return fmt
    allowed = get_allowed_formats(category_key, document_type_key)
    return allowed[0] if allowed else None


def document_type_requires_company_currency(category_key: str, document_type_key: str) -> bool:
    config = get_document_type_config(category_key, document_type_key)
    return bool(config.get("company_currency_required", False))


def supports_monthly_zip_export(
    category_key: str,
    document_type_key: str,
    output_format: str | None = None,
) -> bool:
    config = get_document_type_config(category_key, document_type_key)
    monthly_zip_formats = set(config.get("monthly_zip_formats", []))
    if output_format is None:
        return bool(monthly_zip_formats)
    return output_format in monthly_zip_formats