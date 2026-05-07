from __future__ import annotations

import csv
import io
from datetime import date
from decimal import Decimal
from typing import Any

from utils.currency import currency_code
from utils.distractor_fields import resolve_tabular_value

# ── translations ───────────────────────────────────────────────────────────────
# Reuses the same column-header keys as the XLSX generator.

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "col_invoice_no":   "Invoice No",
        "col_company":      "Company",
        "col_currency":     "Currency",
        "col_site":         "Site",
        "col_city":         "City",
        "col_postcode":     "Postcode",
        "col_meter_id":     "Meter ID",
        "col_period":       "Billing Period",
        "col_period_start": "Period Start",
        "col_period_end":   "Period End",
        "col_issue_date":   "Issue Date",
        "col_due_date":     "Due Date",
        "col_prev_read":    "Prev Reading (kWh)",
        "col_curr_read":    "Curr Reading (kWh)",
        "col_consumption":  "Consumption (kWh)",
        "col_unit_price":   "Unit Price (£/kWh)",
        "col_heat_cost":    "Heat Cost (£)",
        "col_capacity":     "Capacity (kW)",
        "col_cap_rate":     "Cap. Rate (£/kW/mo)",
        "col_supplier_ef": "Supplier EF (kg CO\u2082e/kWh)",
        "col_cap_charge":   "Capacity Charge (£)",
        "col_subtotal":     "Subtotal (£)",
        "col_vat":          "VAT (£)",
        "col_total":        "Total (£)",
    },
    "fr": {
        "col_invoice_no":   "N° de facture",
        "col_company":      "Entreprise",
        "col_currency":     "Devise",
        "col_site":         "Site",
        "col_city":         "Ville",
        "col_postcode":     "Code postal",
        "col_meter_id":     "ID compteur",
        "col_period":       "Période de facturation",
        "col_period_start": "Début de période",
        "col_period_end":   "Fin de période",
        "col_issue_date":   "Date d'émission",
        "col_due_date":     "Date d'échéance",
        "col_prev_read":    "Relevé préc. (kWh)",
        "col_curr_read":    "Relevé actuel (kWh)",
        "col_consumption":  "Consommation (kWh)",
        "col_unit_price":   "Prix unit. (£/kWh)",
        "col_heat_cost":    "Coût thermique (£)",
        "col_capacity":     "Capacité (kW)",
        "col_cap_rate":     "Taux cap. (£/kW/mois)",
        "col_supplier_ef": "FE fournisseur (kg CO\u2082e/kWh)",
        "col_cap_charge":   "Frais de cap. (£)",
        "col_subtotal":     "Sous-total (£)",
        "col_vat":          "TVA (£)",
        "col_total":        "Total (£)",
    },
    "de": {
        "col_invoice_no":   "Rechnungsnr.",
        "col_company":      "Unternehmen",
        "col_currency":     "Währung",
        "col_site":         "Standort",
        "col_city":         "Stadt",
        "col_postcode":     "Postleitzahl",
        "col_meter_id":     "Zähler-ID",
        "col_period":       "Abrechnungszeitraum",
        "col_period_start": "Zeitraum Beginn",
        "col_period_end":   "Zeitraum Ende",
        "col_issue_date":   "Ausstellungsdatum",
        "col_due_date":     "Fälligkeitsdatum",
        "col_prev_read":    "Vorh. Stand (kWh)",
        "col_curr_read":    "Akt. Stand (kWh)",
        "col_consumption":  "Verbrauch (kWh)",
        "col_unit_price":   "Einheitspreis (£/kWh)",
        "col_heat_cost":    "Wärmekosten (£)",
        "col_capacity":     "Leistung (kW)",
        "col_cap_rate":     "Leistungssatz (£/kW/Mo)",
        "col_supplier_ef": "EF Lieferant (kg CO\u2082e/kWh)",
        "col_cap_charge":   "Leistungsgebühr (£)",
        "col_subtotal":     "Zwischensumme (£)",
        "col_vat":          "MwSt. (£)",
        "col_total":        "Gesamt (£)",
    },
    "nl": {
        "col_invoice_no":   "Factuurnr.",
        "col_company":      "Bedrijf",
        "col_currency":     "Valuta",
        "col_site":         "Locatie",
        "col_city":         "Stad",
        "col_postcode":     "Postcode",
        "col_meter_id":     "Meter-ID",
        "col_period":       "Facturatieperiode",
        "col_period_start": "Periode begin",
        "col_period_end":   "Periode einde",
        "col_issue_date":   "Uitgiftedatum",
        "col_due_date":     "Vervaldatum",
        "col_prev_read":    "Vorige stand (kWh)",
        "col_curr_read":    "Huidige stand (kWh)",
        "col_consumption":  "Verbruik (kWh)",
        "col_unit_price":   "Eenheidsprijs (£/kWh)",
        "col_heat_cost":    "Warmtekosten (£)",
        "col_capacity":     "Vermogen (kW)",
        "col_cap_rate":     "Vermogenstarief (£/kW/mnd)",
        "col_supplier_ef": "EF leverancier (kg CO\u2082e/kWh)",
        "col_cap_charge":   "Vermogenstoeslag (£)",
        "col_subtotal":     "Subtotaal (£)",
        "col_vat":          "BTW (£)",
        "col_total":        "Totaal (£)",
    },
}

ELECTRICITY_TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "xl_col_ref": "Reference",
        "xl_col_company": "Company",
        "xl_col_currency": "Currency",
        "xl_col_site": "Site",
        "xl_col_period": "Billing Period",
        "meta_period_start": "Period Start",
        "meta_period_end": "Period End",
        "xl_col_city": "City",
        "xl_col_postcode": "Postcode",
        "xl_col_meter_id": "Meter ID",
        "xl_col_supplier_ef": "Supplier EF (kg CO\u2082e/kWh)",
        "xl_col_unit": "Unit",
        "xl_col_start_read": "Start Reading",
        "xl_col_end_read": "End Reading",
        "xl_col_total_qty": "Total Quantity",
        "xl_col_total_cost": "Total Cost",
        "xl_col_emissions_kg": "Emissions (kg CO\u2082e)",
        "xl_col_emissions_t": "Emissions (tCO\u2082e)",
        "xl_tariff_name": "Tariff Name",
        "xl_tariff_qty": "Quantity",
        "xl_tariff_rate": "Unit Cost",
        "xl_tariff_cost": "Cost",
        "sm_col_consumption": "Consumption",
        "sm_col_tariff_type": "Tariff Type",
        "sm_col_timestamp": "Timestamp",
        "sm_col_import_kwh": "Import kWh",
        "sm_col_export_kwh": "Export kWh",
        "sm_col_end_reading": "End Reading",
    },
    "fr": {
        "xl_col_ref": "R\u00e9f\u00e9rence",
        "xl_col_company": "Entreprise",
        "xl_col_currency": "Devise",
        "xl_col_site": "Site",
        "xl_col_period": "P\u00e9riode de facturation",
        "meta_period_start": "D\u00e9but de p\u00e9riode",
        "meta_period_end": "Fin de p\u00e9riode",
        "xl_col_city": "Ville",
        "xl_col_postcode": "Code postal",
        "xl_col_meter_id": "ID compteur",
        "xl_col_supplier_ef": "FE fournisseur (kg CO\u2082e/kWh)",
        "xl_col_unit": "Unit\u00e9",
        "xl_col_start_read": "Relev\u00e9 initial",
        "xl_col_end_read": "Relev\u00e9 final",
        "xl_col_total_qty": "Quantit\u00e9 totale",
        "xl_col_total_cost": "Co\u00fbt total",
        "xl_col_emissions_kg": "\u00c9missions (kg CO\u2082e)",
        "xl_col_emissions_t": "\u00c9missions (tCO\u2082e)",
        "xl_tariff_name": "Nom du tarif",
        "xl_tariff_qty": "Quantit\u00e9",
        "xl_tariff_rate": "Co\u00fbt unitaire",
        "xl_tariff_cost": "Co\u00fbt",
        "sm_col_consumption": "Consommation",
        "sm_col_tariff_type": "Type de tarif",
        "sm_col_timestamp": "Horodatage",
        "sm_col_import_kwh": "kWh import\u00e9s",
        "sm_col_export_kwh": "kWh export\u00e9s",
        "sm_col_end_reading": "Relev\u00e9 final",
    },
    "de": {
        "xl_col_ref": "Referenz",
        "xl_col_company": "Unternehmen",
        "xl_col_currency": "Währung",
        "xl_col_site": "Standort",
        "xl_col_period": "Abrechnungszeitraum",
        "meta_period_start": "Zeitraum Beginn",
        "meta_period_end": "Zeitraum Ende",
        "xl_col_city": "Stadt",
        "xl_col_postcode": "Postleitzahl",
        "xl_col_meter_id": "Z\u00e4hler-ID",
        "xl_col_supplier_ef": "EF Lieferant (kg CO\u2082e/kWh)",
        "xl_col_unit": "Einheit",
        "xl_col_start_read": "Anfangsz\u00e4hlerstand",
        "xl_col_end_read": "Endz\u00e4hlerstand",
        "xl_col_total_qty": "Gesamtmenge",
        "xl_col_total_cost": "Gesamtkosten",
        "xl_col_emissions_kg": "Emissionen (kg CO\u2082e)",
        "xl_col_emissions_t": "Emissionen (tCO\u2082e)",
        "xl_tariff_name": "Tarifname",
        "xl_tariff_qty": "Menge",
        "xl_tariff_rate": "Einheitspreis",
        "xl_tariff_cost": "Kosten",
        "sm_col_consumption": "Verbrauch",
        "sm_col_tariff_type": "Tariftyp",
        "sm_col_timestamp": "Zeitstempel",
        "sm_col_import_kwh": "Import kWh",
        "sm_col_export_kwh": "Export kWh",
        "sm_col_end_reading": "Endstand",
    },
    "nl": {
        "xl_col_ref": "Referentie",
        "xl_col_company": "Bedrijf",
        "xl_col_currency": "Valuta",
        "xl_col_site": "Locatie",
        "xl_col_period": "Facturatieperiode",
        "meta_period_start": "Periode begin",
        "meta_period_end": "Periode einde",
        "xl_col_city": "Stad",
        "xl_col_postcode": "Postcode",
        "xl_col_meter_id": "Meter-ID",
        "xl_col_supplier_ef": "EF leverancier (kg CO\u2082e/kWh)",
        "xl_col_unit": "Eenheid",
        "xl_col_start_read": "Beginmeterstand",
        "xl_col_end_read": "Eindmeterstand",
        "xl_col_total_qty": "Totale hoeveelheid",
        "xl_col_total_cost": "Totale kosten",
        "xl_col_emissions_kg": "Emissies (kg CO\u2082e)",
        "xl_col_emissions_t": "Emissies (tCO\u2082e)",
        "xl_tariff_name": "Tariefnaam",
        "xl_tariff_qty": "Hoeveelheid",
        "xl_tariff_rate": "Eenheidsprijs",
        "xl_tariff_cost": "Kosten",
        "sm_col_consumption": "Verbruik",
        "sm_col_tariff_type": "Tarieftype",
        "sm_col_timestamp": "Tijdstempel",
        "sm_col_import_kwh": "Import kWh",
        "sm_col_export_kwh": "Export kWh",
        "sm_col_end_reading": "Eindstand",
    },
}

_HEAT_COLUMN_SPECS: list[tuple[str, str, str | None, Any]] = [
    ("invoice_no", "col_invoice_no", "invoice_no", lambda co, si, r: r["invoice_no"]),
    ("company", "col_company", None, lambda co, si, r: co["label"]),
    ("site", "col_site", "site_label", lambda co, si, r: si["label"]),
    ("city", "col_city", "city", lambda co, si, r: r["city"]),
    ("postcode", "col_postcode", "postcode", lambda co, si, r: r["postcode"]),
    ("meter_id", "col_meter_id", "meter_id", lambda co, si, r: r["meter_id"]),
    ("period", "col_period", None, lambda co, si, r: r["billing_period_label"]),
    ("period_start", "col_period_start", None, lambda co, si, r: _fmt_date(r["period_start"])),
    ("period_end", "col_period_end", None, lambda co, si, r: _fmt_date(r["period_end"])),
    ("issue_date", "col_issue_date", None, lambda co, si, r: _fmt_date(r["issue_date"])),
    ("due_date", "col_due_date", None, lambda co, si, r: _fmt_date(r["due_date"])),
    ("prev_read", "col_prev_read", "prev_read", lambda co, si, r: r["prev_read"]),
    ("curr_read", "col_curr_read", "curr_read", lambda co, si, r: r["curr_read"]),
    ("consumption", "col_consumption", "consumption", lambda co, si, r: r["consumption"]),
    ("unit_price", "col_unit_price", "unit_price", lambda co, si, r: _fmt_decimal(r["unit_price"], 4)),
    ("heat_cost", "col_heat_cost", "heat_cost", lambda co, si, r: _fmt_decimal(r["heat_cost"], 2)),
    ("capacity", "col_capacity", "capacity_kw", lambda co, si, r: r["capacity_kw"]),
    ("cap_rate", "col_cap_rate", "capacity_rate", lambda co, si, r: _fmt_decimal(r["capacity_rate"], 2)),
    ("supplier_ef", "col_supplier_ef", "supplier_ef", lambda co, si, r: _fmt_decimal(r["supplier_ef"], 4)),
    ("cap_charge", "col_cap_charge", "capacity_charge", lambda co, si, r: _fmt_decimal(r["capacity_charge"], 2)),
    ("subtotal", "col_subtotal", None, lambda co, si, r: _fmt_decimal(r["subtotal"], 2)),
    ("vat", "col_vat", None, lambda co, si, r: _fmt_decimal(r["vat"], 2)),
    ("total", "col_total", None, lambda co, si, r: _fmt_decimal(r["total"], 2)),
    ("currency", "col_currency", None, lambda co, si, r: currency_code(co.get("currency"))),
]

_HEAT_COLUMN_MAP = {field_id: (header_key, blank_field, accessor) for field_id, header_key, blank_field, accessor in _HEAT_COLUMN_SPECS}

_ELECTRICITY_COLUMN_SPECS: list[tuple[str, str, Any]] = [
    ("reference", "xl_col_ref", lambda co, si: si["ref_no"]),
    ("company", "xl_col_company", lambda co, si: co["label"]),
    ("site", "xl_col_site", lambda co, si: si["label"]),
    ("period", "xl_col_period", lambda co, si: si["billing_period_label"]),
    ("period_start", "meta_period_start", lambda co, si: _fmt_date(si["period_start"])),
    ("period_end", "meta_period_end", lambda co, si: _fmt_date(si["period_end"])),
    ("city", "xl_col_city", lambda co, si: si["city"]),
    ("postcode", "xl_col_postcode", lambda co, si: si["postcode"]),
    ("meter_id", "xl_col_meter_id", lambda co, si: si["meter_id"]),
    ("supplier_ef", "xl_col_supplier_ef", lambda co, si: si["supplier_ef"] if isinstance(si["supplier_ef"], str) else f"{float(si['supplier_ef']):.4f}"),
    ("unit", "xl_col_unit", lambda co, si: si["unit"]),
    ("start_read", "xl_col_start_read", lambda co, si: si["start_reading"]),
    ("end_read", "xl_col_end_read", lambda co, si: si["end_reading"]),
    ("total_qty", "xl_col_total_qty", lambda co, si: si["total_quantity"] if isinstance(si["total_quantity"], str) else f"{float(si['total_quantity']):.2f}"),
    ("total_cost", "xl_col_total_cost", lambda co, si: si["total_cost"] if isinstance(si["total_cost"], str) else f"{float(si['total_cost']):.2f}"),
    ("currency", "xl_col_currency", lambda co, si: currency_code(co.get("currency"))),
    ("emissions_kg", "xl_col_emissions_kg", lambda co, si: si["emissions_kg"] if isinstance(si.get("emissions_kg"), str) else f"{float(si['emissions_kg']):.2f}"),
    ("emissions_t", "xl_col_emissions_t", lambda co, si: si["emissions_t"] if isinstance(si.get("emissions_t"), str) else f"{float(si['emissions_t']):.3f}"),
]

_ELECTRICITY_COLUMN_MAP = {field_id: (header_key, accessor) for field_id, header_key, accessor in _ELECTRICITY_COLUMN_SPECS}

_SMART_METER_MONTHLY_SPEC: list[tuple[str, str, Any]] = [
    ("meter_id", "xl_col_meter_id", lambda row: row["meter_id"]),
    ("site", "xl_col_site", lambda row: row["site_label"]),
    ("period", "xl_col_period", lambda row: row["period_label"]),
    ("start_read", "xl_col_start_read", lambda row: row["start_reading"]),
    ("end_read", "xl_col_end_read", lambda row: row["end_reading"]),
    ("consumption", "sm_col_consumption", lambda row: f"{float(row['consumption']):.2f}"),
    ("unit", "xl_col_unit", lambda row: row["unit"]),
    ("tariff_type", "sm_col_tariff_type", lambda row: row["tariff_type"]),
    ("tariff_cost", "xl_tariff_cost", lambda row: "" if row["cost"] in ("", None) else f"{float(row['cost']):.2f}"),
    ("currency", "xl_col_currency", lambda row: row["currency"]),
]

_SMART_METER_INTERVAL_SPECS: dict[str, list[tuple[str, str, Any]]] = {
    "consumption_diff": [
        ("meter_id", "xl_col_meter_id", lambda row: row["meter_id"]),
        ("timestamp", "sm_col_timestamp", lambda row: row["timestamp"]),
        ("import_kwh", "sm_col_import_kwh", lambda row: f"{float(row['import_kwh']):.4f}"),
        ("export_kwh", "sm_col_export_kwh", lambda row: f"{float(row['export_kwh']):.4f}"),
        ("unit", "xl_col_unit", lambda row: row["unit"]),
    ],
    "cumulative_end_reading": [
        ("meter_id", "xl_col_meter_id", lambda row: row["meter_id"]),
        ("timestamp", "sm_col_timestamp", lambda row: row["timestamp"]),
        ("end_read", "sm_col_end_reading", lambda row: f"{float(row['end_reading']):.4f}"),
        ("unit", "xl_col_unit", lambda row: row["unit"]),
    ],
}


def _fmt_date(value) -> str:
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def _fmt_decimal(value, places: int) -> str:
    if isinstance(value, str):
        return value
    try:
        if not isinstance(value, Decimal):
            value = Decimal(str(value))
        quantizer = Decimal("1." + "0" * places)
        return str(value.quantize(quantizer))
    except Exception:
        return str(value)


def _currency_label_for_sections(sections: list[dict]) -> str:
    codes = {currency_code(section["company"].get("currency")) for section in sections}
    return next(iter(codes)) if len(codes) == 1 else "Currency"


def _replace_currency_labels(strings: dict[str, str], currency_label: str) -> dict[str, str]:
    return {key: value.replace("£", currency_label) for key, value in strings.items()}


def _layout_plan(config: dict) -> dict:
    return config.get("document", {}).get("layout_plan", {})


def _distractor_plan(config: dict):
    return config.get("document", {}).get("distractor_plan")


def _ordered_ids(plan: dict, available_ids: list[str]) -> list[str]:
    excluded = set(plan.get("excluded_fields") or [])
    available_ids = [field_id for field_id in available_ids if field_id not in excluded]
    requested = list(plan.get("column_order") or [])
    if not requested:
        return available_ids
    ordered = [field_id for field_id in requested if field_id in available_ids]
    ordered.extend(field_id for field_id in available_ids if field_id not in ordered)
    return ordered


def _header_text(strings: dict[str, str], plan: dict, field_id: str, header_key: str) -> str:
    return str((plan.get("header_aliases") or {}).get(field_id, strings[header_key]))


def _write_layout_prefix(writer, plan: dict) -> None:
    for row in plan.get("preamble_rows") or []:
        writer.writerow(row)
    for _ in range(int(plan.get("header_row_offset", 0))):
        writer.writerow([])


def _ordered_tabular_ids(base_field_ids: list[str], distractor_plan) -> list[str]:
    ordered = list(base_field_ids)
    if not distractor_plan or not getattr(distractor_plan, "enabled", False):
        return ordered
    for field in distractor_plan.tabular_fields:
        insert_at = len(ordered)
        if field.anchor in ordered:
            anchor_index = ordered.index(field.anchor)
            insert_at = anchor_index + 1 if field.position == "after" else anchor_index
        ordered.insert(insert_at, field.field_id)
    return ordered


def _distractor_field_map(distractor_plan) -> dict[str, Any]:
    if not distractor_plan or not getattr(distractor_plan, "enabled", False):
        return {}
    return {field.field_id: field for field in distractor_plan.tabular_fields}


def _heat_row_key(company: dict, site: dict, rec: dict) -> str:
    del company, site
    return str(rec.get("invoice_no") or rec.get("billing_period_label") or "heat-row")


def _electricity_row_key(company: dict, site: dict) -> str:
    del company
    return str(site.get("ref_no") or site.get("meter_id") or site.get("billing_period_label") or "electricity-row")


def _smart_meter_row_key(row: dict) -> str:
    return str(row.get("timestamp") or row.get("period_label") or row.get("meter_id") or "smart-meter-row")


def _electricity_tariff_headers(strings: dict[str, str], max_tariffs: int) -> list[str]:
    headers: list[str] = []
    for idx in range(max_tariffs):
        prefix = f"Tariff {idx + 1}"
        headers.extend([
            f"{prefix}: {strings['xl_tariff_name']}",
            f"{prefix}: {strings['xl_tariff_qty']}",
            f"{prefix}: {strings['xl_tariff_rate']}",
            f"{prefix}: {strings['xl_tariff_cost']}",
        ])
    return headers


def _electricity_tariff_row(site: dict, max_tariffs: int) -> list[str]:
    row: list[str] = []
    tariffs = site.get("tariffs", [])
    for idx in range(max_tariffs):
        if idx < len(tariffs):
            tariff = tariffs[idx]
            row.extend([
                tariff["name"],
                tariff["quantity"] if isinstance(tariff["quantity"], str) else f"{float(tariff['quantity']):.2f}",
                tariff["unit_cost"] if isinstance(tariff["unit_cost"], str) else f"{float(tariff['unit_cost']):.4f}",
                tariff["cost"] if isinstance(tariff["cost"], str) else f"{float(tariff['cost']):.2f}",
            ])
        else:
            row.extend(["", "", "", ""])
    return row


def _generate_heat_csv(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
) -> bytes:
    """Build a UTF-8 CSV of billing detail rows and return bytes."""
    lang = config["document"].get("language", "en")
    strings = _replace_currency_labels(TRANSLATIONS.get(lang, TRANSLATIONS["en"]), _currency_label_for_sections(sections))
    omit = blank_fields or set()
    plan = _layout_plan(config)
    distractor_plan = _distractor_plan(config)
    distractor_fields = _distractor_field_map(distractor_plan)
    ordered_field_ids = _ordered_tabular_ids(_ordered_ids(plan, list(_HEAT_COLUMN_MAP)), distractor_plan)

    headers = [
        distractor_fields[field_id].label if field_id in distractor_fields else _header_text(strings, plan, field_id, _HEAT_COLUMN_MAP[field_id][0])
        for field_id in ordered_field_ids
    ]

    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\n")
    _write_layout_prefix(writer, plan)
    writer.writerow(headers)

    for section in sections:
        company = section["company"]
        site = section["site"]
        for rec in section["records"]:
            row = []
            for field_id in ordered_field_ids:
                if field_id in distractor_fields:
                    row.append(resolve_tabular_value(distractor_plan, distractor_fields[field_id], _heat_row_key(company, site, rec)))
                    continue
                _, blank_field, accessor = _HEAT_COLUMN_MAP[field_id]
                if blank_field and blank_field in omit:
                    row.append("")
                else:
                    row.append(accessor(company, site, rec))
            writer.writerow(row)

    return buf.getvalue().encode("utf-8-sig")  # BOM for Excel compatibility


def _generate_electricity_csv(config: dict, sections: list[dict]) -> bytes:
    lang = config["document"].get("language", "en")
    strings = ELECTRICITY_TRANSLATIONS.get(lang, ELECTRICITY_TRANSLATIONS["en"])
    plan = _layout_plan(config)
    distractor_plan = _distractor_plan(config)
    distractor_fields = _distractor_field_map(distractor_plan)

    buf = io.StringIO()
    writer = csv.writer(buf)
    _write_layout_prefix(writer, plan)

    max_tariffs = max((len(section["site"].get("tariffs", [])) for section in sections), default=0)
    ordered_field_ids = _ordered_tabular_ids(_ordered_ids(plan, list(_ELECTRICITY_COLUMN_MAP)), distractor_plan)
    headers: list[str] = []
    tariff_inserted = False
    for field_id in ordered_field_ids:
        if field_id in distractor_fields:
            headers.append(distractor_fields[field_id].label)
        else:
            headers.append(_header_text(strings, plan, field_id, _ELECTRICITY_COLUMN_MAP[field_id][0]))
        if field_id == "total_qty" and plan.get("tariff_block_position") == "after_total_qty":
            headers.extend(_electricity_tariff_headers(strings, max_tariffs))
            tariff_inserted = True
    if not tariff_inserted:
        headers.extend(_electricity_tariff_headers(strings, max_tariffs))
    writer.writerow(headers)

    for section in sections:
        company = section["company"]
        site = section["site"]
        row: list[str] = []
        tariff_inserted = False
        for field_id in ordered_field_ids:
            if field_id in distractor_fields:
                row.append(resolve_tabular_value(distractor_plan, distractor_fields[field_id], _electricity_row_key(company, site)))
            else:
                _, accessor = _ELECTRICITY_COLUMN_MAP[field_id]
                row.append(accessor(company, site))
            if field_id == "total_qty" and plan.get("tariff_block_position") == "after_total_qty":
                row.extend(_electricity_tariff_row(site, max_tariffs))
                tariff_inserted = True
        if not tariff_inserted:
            row.extend(_electricity_tariff_row(site, max_tariffs))
        writer.writerow(row)

    return buf.getvalue().encode("utf-8-sig")


def _generate_smart_meter_csv(config: dict, sections: list[dict]) -> bytes:
    from generators.electricity_generator import build_smart_meter_rows

    lang = config["document"].get("language", "en")
    strings = ELECTRICITY_TRANSLATIONS.get(lang, ELECTRICITY_TRANSLATIONS["en"])
    mode = str(config["document"].get("smart_meter_data_granularity", "monthly")).lower()
    value_mode = str(config["document"].get("smart_meter_interval_value_mode", "consumption_diff")).lower()
    plan = _layout_plan(config)
    distractor_plan = _distractor_plan(config)
    distractor_fields = _distractor_field_map(distractor_plan)
    rows = build_smart_meter_rows(config, sections)

    buf = io.StringIO()
    writer = csv.writer(buf)
    _write_layout_prefix(writer, plan)

    if mode == "interval":
        interval_spec = _SMART_METER_INTERVAL_SPECS.get(value_mode, _SMART_METER_INTERVAL_SPECS["consumption_diff"])
        interval_map = {field_id: (header_key, accessor) for field_id, header_key, accessor in interval_spec}
        ordered_field_ids = _ordered_tabular_ids(_ordered_ids(plan, list(interval_map)), distractor_plan)
        writer.writerow([
            distractor_fields[field_id].label if field_id in distractor_fields else _header_text(strings, plan, field_id, interval_map[field_id][0])
            for field_id in ordered_field_ids
        ])
        for row in rows:
            writer.writerow([
                resolve_tabular_value(distractor_plan, distractor_fields[field_id], _smart_meter_row_key(row))
                if field_id in distractor_fields
                else interval_map[field_id][1](row)
                for field_id in ordered_field_ids
            ])
    else:
        monthly_map = {field_id: (header_key, accessor) for field_id, header_key, accessor in _SMART_METER_MONTHLY_SPEC}
        ordered_field_ids = _ordered_tabular_ids(_ordered_ids(plan, list(monthly_map)), distractor_plan)
        writer.writerow([
            distractor_fields[field_id].label if field_id in distractor_fields else _header_text(strings, plan, field_id, monthly_map[field_id][0])
            for field_id in ordered_field_ids
        ])
        for row in rows:
            writer.writerow([
                resolve_tabular_value(distractor_plan, distractor_fields[field_id], _smart_meter_row_key(row))
                if field_id in distractor_fields
                else monthly_map[field_id][1](row)
                for field_id in ordered_field_ids
            ])

    return buf.getvalue().encode("utf-8-sig")


def _generate_heat_supplier_portal_csv(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
) -> bytes:
    return _generate_heat_csv(config, sections, blank_fields=blank_fields)


def _generate_electricity_supplier_portal_csv(config: dict, sections: list[dict]) -> bytes:
    return _generate_electricity_csv(config, sections)


def generate_csv(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
    category: str = "heat",
) -> bytes:
    document_type = config["document"].get("type")
    if category == "electricity":
        if document_type == "smart_meter_data":
            return _generate_smart_meter_csv(config, sections)
        if document_type == "supplier_portal_data":
            return _generate_electricity_supplier_portal_csv(config, sections)
        raise NotImplementedError(f"CSV generation is not supported for electricity document type '{document_type}'.")
    if document_type == "supplier_portal_data":
        return _generate_heat_supplier_portal_csv(config, sections, blank_fields=blank_fields)
    raise NotImplementedError(f"CSV generation is not supported for heat document type '{document_type}'.")
