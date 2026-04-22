from __future__ import annotations

import csv
import io
from datetime import date
from decimal import Decimal
from typing import Any

from utils.currency import currency_code

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

# Column spec: (header_key, record_accessor)
# accessor is a callable (company, site, rec) -> value
_COLUMN_SPEC: list[tuple[str, Any]] = [
    ("col_invoice_no",   lambda co, si, r: r["invoice_no"]),
    ("col_company",      lambda co, si, r: co["label"]),
    ("col_site",         lambda co, si, r: si["label"]),
    ("col_city",         lambda co, si, r: r["city"]),
    ("col_postcode",     lambda co, si, r: r["postcode"]),
    ("col_meter_id",     lambda co, si, r: r["meter_id"]),
    ("col_period",       lambda co, si, r: r["billing_period_label"]),
    ("col_period_start", lambda co, si, r: _fmt_date(r["period_start"])),
    ("col_period_end",   lambda co, si, r: _fmt_date(r["period_end"])),
    ("col_issue_date",   lambda co, si, r: _fmt_date(r["issue_date"])),
    ("col_due_date",     lambda co, si, r: _fmt_date(r["due_date"])),
    ("col_prev_read",    lambda co, si, r: r["prev_read"]),
    ("col_curr_read",    lambda co, si, r: r["curr_read"]),
    ("col_consumption",  lambda co, si, r: r["consumption"]),
    ("col_unit_price",   lambda co, si, r: _fmt_decimal(r["unit_price"], 4)),
    ("col_heat_cost",    lambda co, si, r: _fmt_decimal(r["heat_cost"], 2)),
    ("col_capacity",     lambda co, si, r: r["capacity_kw"]),
    ("col_cap_rate",     lambda co, si, r: _fmt_decimal(r["capacity_rate"], 2)),
    ("col_supplier_ef",  lambda co, si, r: _fmt_decimal(r["supplier_ef"], 4)),
    ("col_cap_charge",   lambda co, si, r: _fmt_decimal(r["capacity_charge"], 2)),
    ("col_subtotal",     lambda co, si, r: _fmt_decimal(r["subtotal"], 2)),
    ("col_vat",          lambda co, si, r: _fmt_decimal(r["vat"], 2)),
    ("col_total",        lambda co, si, r: _fmt_decimal(r["total"], 2)),
    ("col_currency",     lambda co, si, r: currency_code(co.get("currency"))),
]

# Maps column header key → record field name for blank_fields checking
_BLANK_FIELD_MAP: dict[str, str] = {
    "col_invoice_no":   "invoice_no",
    "col_site":         "site_label",
    "col_city":         "city",
    "col_postcode":     "postcode",
    "col_meter_id":     "meter_id",
    "col_prev_read":    "prev_read",
    "col_curr_read":    "curr_read",
    "col_consumption":  "consumption",
    "col_unit_price":   "unit_price",
    "col_heat_cost":    "heat_cost",
    "col_capacity":     "capacity_kw",
    "col_cap_rate":     "capacity_rate",
    "col_supplier_ef":  "supplier_ef",
    "col_cap_charge":   "capacity_charge",
}


def _fmt_date(value: date) -> str:
    return value.isoformat()


def _fmt_decimal(value, places: int) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    quantizer = Decimal("1." + "0" * places)
    return str(value.quantize(quantizer))


def _currency_label_for_sections(sections: list[dict]) -> str:
    codes = {currency_code(section["company"].get("currency")) for section in sections}
    return next(iter(codes)) if len(codes) == 1 else "Currency"


def _replace_currency_labels(strings: dict[str, str], currency_label: str) -> dict[str, str]:
    return {key: value.replace("£", currency_label) for key, value in strings.items()}


def _generate_heat_csv(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
) -> bytes:
    """Build a UTF-8 CSV of billing detail rows and return bytes."""
    lang = config["document"].get("language", "en")
    strings = _replace_currency_labels(TRANSLATIONS.get(lang, TRANSLATIONS["en"]), _currency_label_for_sections(sections))
    omit = blank_fields or set()

    headers = [strings[key] for key, _ in _COLUMN_SPEC]

    buf = io.StringIO()
    writer = csv.writer(buf, lineterminator="\n")
    writer.writerow(headers)

    for section in sections:
        company = section["company"]
        site = section["site"]
        for rec in section["records"]:
            row = []
            for key, accessor in _COLUMN_SPEC:
                blank_field = _BLANK_FIELD_MAP.get(key)
                if blank_field and blank_field in omit:
                    row.append("")
                else:
                    row.append(accessor(company, site, rec))
            writer.writerow(row)

    return buf.getvalue().encode("utf-8-sig")  # BOM for Excel compatibility


def _generate_electricity_csv(config: dict, sections: list[dict]) -> bytes:
    lang = config["document"].get("language", "en")
    strings = ELECTRICITY_TRANSLATIONS.get(lang, ELECTRICITY_TRANSLATIONS["en"])

    buf = io.StringIO()
    writer = csv.writer(buf)

    max_tariffs = max((len(section["site"].get("tariffs", [])) for section in sections), default=0)

    headers = [
        strings["xl_col_ref"],
        strings["xl_col_company"],
        strings["xl_col_site"],
        strings["xl_col_period"],
        strings["meta_period_start"],
        strings["meta_period_end"],
        strings["xl_col_city"],
        strings["xl_col_postcode"],
        strings["xl_col_meter_id"],
        strings["xl_col_supplier_ef"],
        strings["xl_col_unit"],
        strings["xl_col_start_read"],
        strings["xl_col_end_read"],
        strings["xl_col_total_qty"],
        strings["xl_col_total_cost"],
        strings["xl_col_currency"],
        strings["xl_col_emissions_kg"],
        strings["xl_col_emissions_t"],
    ]
    for idx in range(max_tariffs):
        prefix = f"Tariff {idx + 1}"
        headers += [
            f"{prefix}: {strings['xl_tariff_name']}",
            f"{prefix}: {strings['xl_tariff_qty']}",
            f"{prefix}: {strings['xl_tariff_rate']}",
            f"{prefix}: {strings['xl_tariff_cost']}",
        ]
    writer.writerow(headers)

    for section in sections:
        company = section["company"]
        site = section["site"]
        period_start = site["period_start"]
        period_end = site["period_end"]
        row = [
            site["ref_no"],
            company["label"],
            site["label"],
            site["billing_period_label"],
            period_start.strftime("%Y-%m-%d") if hasattr(period_start, "strftime") else str(period_start),
            period_end.strftime("%Y-%m-%d") if hasattr(period_end, "strftime") else str(period_end),
            site["city"],
            site["postcode"],
            site["meter_id"],
            f"{float(site['supplier_ef']):.4f}",
            site["unit"],
            site["start_reading"],
            site["end_reading"],
            f"{float(site['total_quantity']):.2f}",
            f"{float(site['total_cost']):.2f}",
            currency_code(company.get("currency")),
            f"{float(site['emissions_kg']):.2f}",
            f"{float(site['emissions_t']):.3f}",
        ]
        tariffs = site.get("tariffs", [])
        for idx in range(max_tariffs):
            if idx < len(tariffs):
                tariff = tariffs[idx]
                row += [
                    tariff["name"],
                    f"{float(tariff['quantity']):.2f}",
                    f"{float(tariff['unit_cost']):.4f}",
                    f"{float(tariff['cost']):.2f}",
                ]
            else:
                row += ["", "", "", ""]
        writer.writerow(row)

    return buf.getvalue().encode("utf-8-sig")


def _generate_smart_meter_csv(config: dict, sections: list[dict]) -> bytes:
    from generators.electricity_generator import build_smart_meter_rows

    lang = config["document"].get("language", "en")
    strings = ELECTRICITY_TRANSLATIONS.get(lang, ELECTRICITY_TRANSLATIONS["en"])
    mode = str(config["document"].get("smart_meter_data_granularity", "monthly")).lower()
    value_mode = str(config["document"].get("smart_meter_interval_value_mode", "consumption_diff")).lower()
    rows = build_smart_meter_rows(config, sections)

    buf = io.StringIO()
    writer = csv.writer(buf)

    if mode == "interval":
        if value_mode == "cumulative_end_reading":
            writer.writerow([
                strings["xl_col_meter_id"],
                strings["sm_col_timestamp"],
                strings["sm_col_end_reading"],
                strings["xl_col_unit"],
            ])
            for row in rows:
                writer.writerow([
                    row["meter_id"],
                    row["timestamp"],
                    f"{float(row['end_reading']):.4f}",
                    row["unit"],
                ])
        else:
            writer.writerow([
                strings["xl_col_meter_id"],
                strings["sm_col_timestamp"],
                strings["sm_col_import_kwh"],
                strings["sm_col_export_kwh"],
                strings["xl_col_unit"],
            ])
            for row in rows:
                writer.writerow([
                    row["meter_id"],
                    row["timestamp"],
                    f"{float(row['import_kwh']):.4f}",
                    f"{float(row['export_kwh']):.4f}",
                    row["unit"],
                ])
    else:
        writer.writerow([
            strings["xl_col_meter_id"],
            strings["xl_col_site"],
            strings["xl_col_period"],
            strings["xl_col_start_read"],
            strings["xl_col_end_read"],
            strings["sm_col_consumption"],
            strings["xl_col_unit"],
            strings["sm_col_tariff_type"],
            strings["xl_tariff_cost"],
            strings["xl_col_currency"],
        ])
        for row in rows:
            writer.writerow([
                row["meter_id"],
                row["site_label"],
                row["period_label"],
                row["start_reading"],
                row["end_reading"],
                f"{float(row['consumption']):.2f}",
                row["unit"],
                row["tariff_type"],
                "" if row["cost"] in ("", None) else f"{float(row['cost']):.2f}",
                row["currency"],
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
