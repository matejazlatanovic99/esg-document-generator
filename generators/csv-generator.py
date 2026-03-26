from __future__ import annotations

import csv
import io
from datetime import date
from decimal import Decimal
from typing import Any

# ── translations ───────────────────────────────────────────────────────────────
# Reuses the same column-header keys as the XLSX generator.

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "col_invoice_no":   "Invoice No",
        "col_company":      "Company",
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
        "col_cap_charge":   "Capacity Charge (£)",
        "col_subtotal":     "Subtotal (£)",
        "col_vat":          "VAT (£)",
        "col_total":        "Total (£)",
    },
    "fr": {
        "col_invoice_no":   "N° de facture",
        "col_company":      "Entreprise",
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
        "col_cap_charge":   "Frais de cap. (£)",
        "col_subtotal":     "Sous-total (£)",
        "col_vat":          "TVA (£)",
        "col_total":        "Total (£)",
    },
    "de": {
        "col_invoice_no":   "Rechnungsnr.",
        "col_company":      "Unternehmen",
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
        "col_cap_charge":   "Leistungsgebühr (£)",
        "col_subtotal":     "Zwischensumme (£)",
        "col_vat":          "MwSt. (£)",
        "col_total":        "Gesamt (£)",
    },
    "nl": {
        "col_invoice_no":   "Factuurnr.",
        "col_company":      "Bedrijf",
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
        "col_cap_charge":   "Vermogenstoeslag (£)",
        "col_subtotal":     "Subtotaal (£)",
        "col_vat":          "BTW (£)",
        "col_total":        "Totaal (£)",
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
    ("col_cap_charge",   lambda co, si, r: _fmt_decimal(r["capacity_charge"], 2)),
    ("col_subtotal",     lambda co, si, r: _fmt_decimal(r["subtotal"], 2)),
    ("col_vat",          lambda co, si, r: _fmt_decimal(r["vat"], 2)),
    ("col_total",        lambda co, si, r: _fmt_decimal(r["total"], 2)),
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
    "col_cap_charge":   "capacity_charge",
}


def _fmt_date(value: date) -> str:
    return value.isoformat()


def _fmt_decimal(value, places: int) -> str:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    quantizer = Decimal("1." + "0" * places)
    return str(value.quantize(quantizer))


def generate_csv(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
) -> bytes:
    """Build a UTF-8 CSV of billing detail rows and return bytes."""
    lang = config["document"].get("language", "en")
    strings = TRANSLATIONS.get(lang, TRANSLATIONS["en"])
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
