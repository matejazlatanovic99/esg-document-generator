from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal
from io import BytesIO
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from utils.currency import currency_code

# ── translations ──────────────────────────────────────────────────────────────

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        # summary metadata
        "meta_period":      "Financial Period",
        "meta_start":       "Period Start",
        "meta_end":         "Period End",
        "meta_generated":   "Generated",
        # summary table headers
        "sum_company":      "Company",
        "sum_currency":     "Currency",
        "sum_sites":        "Sites",
        "sum_invoices":     "Invoices",
        "sum_heat_cost":    "Heat Cost (£)",
        "sum_cap_charge":   "Capacity Charge (£)",
        "sum_subtotal":     "Subtotal (£)",
        "sum_vat":          "VAT (£)",
        "sum_total":        "Total Due (£)",
        "sum_grand":        "TOTAL",
        # detail column headers
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
        "meta_period":      "Période financière",
        "meta_start":       "Début de période",
        "meta_end":         "Fin de période",
        "meta_generated":   "Généré",
        "sum_company":      "Entreprise",
        "sum_currency":     "Devise",
        "sum_sites":        "Sites",
        "sum_invoices":     "Factures",
        "sum_heat_cost":    "Coût thermique (£)",
        "sum_cap_charge":   "Frais de capacité (£)",
        "sum_subtotal":     "Sous-total (£)",
        "sum_vat":          "TVA (£)",
        "sum_total":        "Total dû (£)",
        "sum_grand":        "TOTAL",
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
        "meta_period":      "Finanzzeitraum",
        "meta_start":       "Zeitraum Beginn",
        "meta_end":         "Zeitraum Ende",
        "meta_generated":   "Erstellt",
        "sum_company":      "Unternehmen",
        "sum_currency":     "Währung",
        "sum_sites":        "Standorte",
        "sum_invoices":     "Rechnungen",
        "sum_heat_cost":    "Wärmekosten (£)",
        "sum_cap_charge":   "Leistungsgebühr (£)",
        "sum_subtotal":     "Zwischensumme (£)",
        "sum_vat":          "MwSt. (£)",
        "sum_total":        "Gesamtbetrag (£)",
        "sum_grand":        "GESAMT",
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
        "meta_period":      "Financiële periode",
        "meta_start":       "Periode begin",
        "meta_end":         "Periode einde",
        "meta_generated":   "Gegenereerd",
        "sum_company":      "Bedrijf",
        "sum_currency":     "Valuta",
        "sum_sites":        "Locaties",
        "sum_invoices":     "Facturen",
        "sum_heat_cost":    "Warmtekosten (£)",
        "sum_cap_charge":   "Vermogenstoeslag (£)",
        "sum_subtotal":     "Subtotaal (£)",
        "sum_vat":          "BTW (£)",
        "sum_total":        "Totaal (£)",
        "sum_grand":        "TOTAAL",
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
        "xl_meta_period": "Financial Period",
        "xl_meta_start": "Period Start",
        "xl_meta_end": "Period End",
        "xl_meta_generated": "Generated",
        "xl_sum_company": "Company",
        "xl_sum_sites": "Sites",
        "xl_sum_qty": "Total Consumption",
        "xl_sum_cost": "Total Cost",
        "xl_sum_emissions_t": "Total tCO\u2082e",
        "xl_sum_grand": "TOTAL",
        "xl_col_ref": "Reference",
        "xl_col_company": "Company",
        "xl_col_currency": "Currency",
        "xl_col_site": "Site",
        "xl_col_period": "Billing Period",
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
        "xl_meta_period": "P\u00e9riode financi\u00e8re",
        "xl_meta_start": "D\u00e9but de p\u00e9riode",
        "xl_meta_end": "Fin de p\u00e9riode",
        "xl_meta_generated": "G\u00e9n\u00e9r\u00e9",
        "xl_sum_company": "Entreprise",
        "xl_sum_sites": "Sites",
        "xl_sum_qty": "Consommation totale",
        "xl_sum_cost": "Co\u00fbt total",
        "xl_sum_emissions_t": "Total tCO\u2082e",
        "xl_sum_grand": "TOTAL",
        "xl_col_ref": "R\u00e9f\u00e9rence",
        "xl_col_company": "Entreprise",
        "xl_col_currency": "Devise",
        "xl_col_site": "Site",
        "xl_col_period": "P\u00e9riode de facturation",
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
        "xl_meta_period": "Finanzzeitraum",
        "xl_meta_start": "Zeitraum Beginn",
        "xl_meta_end": "Zeitraum Ende",
        "xl_meta_generated": "Erstellt",
        "xl_sum_company": "Unternehmen",
        "xl_sum_sites": "Standorte",
        "xl_sum_qty": "Gesamtverbrauch",
        "xl_sum_cost": "Gesamtkosten",
        "xl_sum_emissions_t": "Gesamt tCO\u2082e",
        "xl_sum_grand": "GESAMT",
        "xl_col_ref": "Referenz",
        "xl_col_company": "Unternehmen",
        "xl_col_currency": "Währung",
        "xl_col_site": "Standort",
        "xl_col_period": "Abrechnungszeitraum",
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
        "xl_meta_period": "Financi\u00eble periode",
        "xl_meta_start": "Periode begin",
        "xl_meta_end": "Periode einde",
        "xl_meta_generated": "Gegenereerd",
        "xl_sum_company": "Bedrijf",
        "xl_sum_sites": "Locaties",
        "xl_sum_qty": "Totaal verbruik",
        "xl_sum_cost": "Totale kosten",
        "xl_sum_emissions_t": "Totaal tCO\u2082e",
        "xl_sum_grand": "TOTAAL",
        "xl_col_ref": "Referentie",
        "xl_col_company": "Bedrijf",
        "xl_col_currency": "Valuta",
        "xl_col_site": "Locatie",
        "xl_col_period": "Facturatieperiode",
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


# ── palette ───────────────────────────────────────────────────────────────────
_WHITE = "FFFFFF"
_DARK = "1F2328"
_MID = "5A6066"
_ALT = "F7F9FC"
_SOFT = "DCEBF5"
_BORDER = "C9CDD2"

# ── style primitives ──────────────────────────────────────────────────────────

def _font(bold: bool = False, color: str = _DARK, size: int = 10) -> Font:
    return Font(name="Calibri", bold=bold, color=color, size=size)


def _fill(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def _border() -> Border:
    side = Side(style="thin", color=_BORDER)
    return Border(left=side, right=side, top=side, bottom=side)


def _align(horizontal: str = "left") -> Alignment:
    return Alignment(horizontal=horizontal, vertical="center")


def _header_cell(cell, text: str, accent: str) -> None:
    cell.value = text
    cell.font = _font(bold=True, color=_WHITE)
    cell.fill = _fill(accent)
    cell.border = _border()
    cell.alignment = _align("center")


def _data_cell(cell, value: Any, fmt: str | None = None, alt: bool = False, bold: bool = False) -> None:
    cell.value = value
    cell.font = _font(bold=bold)
    cell.fill = _fill(_ALT if alt else _WHITE)
    cell.border = _border()
    cell.alignment = _align("right" if isinstance(value, (int, float, Decimal)) else "left")
    if fmt:
        cell.number_format = fmt


# ── detail sheet column spec: (header_key, width, number_format) ─────────────
_DETAIL_COL_SPEC: list[tuple[str, int, str | None]] = [
    ("col_invoice_no",   20, None),
    ("col_company",      30, None),
    ("col_site",         20, None),
    ("col_city",         14, None),
    ("col_postcode",     11, None),
    ("col_meter_id",     22, None),
    ("col_period",       18, None),
    ("col_period_start", 14, "DD MMM YYYY"),
    ("col_period_end",   14, "DD MMM YYYY"),
    ("col_issue_date",   14, "DD MMM YYYY"),
    ("col_due_date",     14, "DD MMM YYYY"),
    ("col_prev_read",    20, "#,##0"),
    ("col_curr_read",    20, "#,##0"),
    ("col_consumption",  20, "#,##0"),
    ("col_unit_price",   20, "0.000"),
    ("col_heat_cost",    16, "#,##0.00"),
    ("col_capacity",     14, "#,##0"),
    ("col_cap_rate",     20, "0.00"),
    ("col_supplier_ef",  24, "0.0000"),
    ("col_cap_charge",   20, "#,##0.00"),
    ("col_subtotal",     16, "#,##0.00"),
    ("col_vat",          13, "#,##0.00"),
    ("col_total",        16, "#,##0.00"),
    ("col_currency",     10, None),
]


def _detail_cols(strings: dict) -> list[tuple[str, int, str | None]]:
    return [(strings[key], width, fmt) for key, width, fmt in _DETAIL_COL_SPEC]


# ── helpers ───────────────────────────────────────────────────────────────────

def _safe_sheet_name(name: str, max_len: int = 31) -> str:
    """Return a valid Excel sheet name (max 31 chars, no forbidden characters)."""
    sanitised = re.sub(r'[\\/?*\[\]:]', '', name).strip()
    return (sanitised or "Company")[:max_len]


def _sections_by_company(sections: list[dict]) -> dict[str, list[dict]]:
    """Group sections by company label, preserving insertion order."""
    grouped: dict[str, list[dict]] = {}
    for section in sections:
        label = section["company"]["label"]
        grouped.setdefault(label, []).append(section)
    return grouped


def _currency_label_for_sections(sections: list[dict]) -> str:
    codes = {currency_code(section["company"].get("currency")) for section in sections}
    return next(iter(codes)) if len(codes) == 1 else "Currency"


def _replace_currency_labels(strings: dict[str, str], currency_label: str) -> dict[str, str]:
    return {key: value.replace("£", currency_label) for key, value in strings.items()}


# ── public API ────────────────────────────────────────────────────────────────

def _generate_heat_xlsx(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
    split_by_company: bool = False,
    include_summary: bool = True,
) -> bytes:
    """Build a styled XLSX workbook from billing sections and return bytes.

    blank_fields: record field names whose cells should be left empty (QA testing).
    split_by_company: when True, create one detail sheet per company instead of
                      a single combined "Billing Detail" sheet.
    """
    lang = config["document"].get("language", "en")
    strings = _replace_currency_labels(TRANSLATIONS.get(lang, TRANSLATIONS["en"]), _currency_label_for_sections(sections))

    default_accent = sections[0]["company"]["accent"].lstrip("#") if sections else "1E5B88"
    omit = blank_fields or set()

    wb = openpyxl.Workbook()

    if include_summary:
        wb.active.title = "Summary"
        _build_summary_sheet(wb.active, config, sections, default_accent, strings)
        detail_seed_sheet = None
    else:
        detail_seed_sheet = wb.active

    if split_by_company:
        by_company = list(_sections_by_company(sections).items())
        for idx, (label, co_sections) in enumerate(by_company):
            accent = co_sections[0]["company"]["accent"].lstrip("#")
            sheet_name = _safe_sheet_name(label)
            target_sheet = detail_seed_sheet if idx == 0 and detail_seed_sheet is not None else wb.create_sheet(sheet_name)
            target_sheet.title = sheet_name
            _build_detail_sheet(target_sheet, co_sections, accent, omit, strings)
    else:
        target_sheet = detail_seed_sheet if detail_seed_sheet is not None else wb.create_sheet("Billing Detail")
        target_sheet.title = "Billing Detail"
        _build_detail_sheet(target_sheet, sections, default_accent, omit, strings)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _generate_electricity_xlsx(config: dict, sections: list[dict], include_summary: bool = True) -> bytes:
    lang = config["document"].get("language", "en")
    strings = ELECTRICITY_TRANSLATIONS.get(lang, ELECTRICITY_TRANSLATIONS["en"])
    financial_period = config["financial_period"]

    accent_hex = sections[0]["company"]["accent"] if sections else "#1E5B88"
    accent_r, accent_g, accent_b = (int(accent_hex[i:i + 2], 16) for i in (1, 3, 5))

    def header_fill(hex_color: str) -> PatternFill:
        red, green, blue = (int(hex_color[i:i + 2], 16) for i in (1, 3, 5))
        return PatternFill("solid", fgColor=f"{red:02X}{green:02X}{blue:02X}")

    def header_font(white_text: bool = True) -> Font:
        return Font(name="Calibri", bold=True, color="FFFFFF" if white_text else "1F2328", size=9)

    def thin_border() -> Border:
        side = Side(style="thin", color="D5DADF")
        return Border(left=side, right=side, top=side, bottom=side)

    workbook = openpyxl.Workbook()
    hdr_fill = header_fill(accent_hex)
    hdr_font = header_font()
    border = thin_border()

    if include_summary:
        summary = workbook.active
        summary.title = "Summary"

        summary["A1"] = strings["xl_meta_period"]
        summary["B1"] = financial_period["label"]
        summary["A2"] = strings["xl_meta_start"]
        summary["B2"] = financial_period["start_date"].strftime("%d %b %Y") if hasattr(financial_period["start_date"], "strftime") else str(financial_period["start_date"])
        summary["A3"] = strings["xl_meta_end"]
        summary["B3"] = financial_period["end_date"].strftime("%d %b %Y") if hasattr(financial_period["end_date"], "strftime") else str(financial_period["end_date"])
        summary["A4"] = strings["xl_meta_generated"]
        summary["B4"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        for row in range(1, 5):
            summary.cell(row, 1).font = Font(name="Calibri", bold=True, size=9)

        summary_headers = [
            strings["xl_sum_company"],
            strings["xl_sum_sites"],
            strings["xl_sum_qty"],
            strings["xl_sum_cost"],
            strings["xl_col_currency"],
            strings["xl_sum_emissions_t"],
        ]
        header_row = 6
        for col_idx, header in enumerate(summary_headers, start=1):
            cell = summary.cell(header_row, col_idx, header)
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

        from collections import defaultdict

        by_company: dict[str, list] = defaultdict(list)
        for section in sections:
            by_company[section["company"]["label"]].append(section)

        grand_qty = Decimal("0")
        grand_cost = Decimal("0")
        grand_emissions = Decimal("0")

        data_row = header_row + 1
        for company_label, company_sections in by_company.items():
            site_count = len({section["site"]["_site_uid"] for section in company_sections})
            total_qty = sum(section["site"]["total_quantity"] for section in company_sections)
            total_cost = sum(section["site"]["total_cost"] for section in company_sections)
            total_emissions = sum(section["site"]["emissions_t"] for section in company_sections)

            row_values = [
                company_label,
                site_count,
                float(total_qty),
                float(total_cost),
                currency_code(company_sections[0]["company"].get("currency")),
                float(total_emissions),
            ]
            for col_idx, value in enumerate(row_values, start=1):
                cell = summary.cell(data_row, col_idx, value)
                cell.border = border
                cell.font = Font(name="Calibri", size=9)
                if col_idx in {2, 3}:
                    cell.number_format = "#,##0"
                elif col_idx in {4, 6}:
                    cell.number_format = "#,##0.00"
            data_row += 1
            grand_qty += total_qty
            grand_cost += total_cost
            grand_emissions += total_emissions

        grand_fill = PatternFill("solid", fgColor=f"{accent_r:02X}{accent_g:02X}{accent_b:02X}")
        grand_values = [
            strings["xl_sum_grand"],
            len({section["site"]["_site_uid"] for section in sections}),
            float(grand_qty),
            float(grand_cost),
            "",
            float(grand_emissions),
        ]
        for col_idx, value in enumerate(grand_values, start=1):
            cell = summary.cell(data_row, col_idx, value)
            cell.fill = grand_fill
            cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
            cell.border = border
            if col_idx in {2, 3}:
                cell.number_format = "#,##0"
            elif col_idx in {4, 6}:
                cell.number_format = "#,##0.00"

        for col_idx, width in enumerate([30, 8, 18, 18, 10, 14], start=1):
            summary.column_dimensions[get_column_letter(col_idx)].width = width

        detail = workbook.create_sheet("Detail")
    else:
        detail = workbook.active
        detail.title = "Detail"
    max_tariffs = max((len(section["site"].get("tariffs", [])) for section in sections), default=0)

    detail_headers = [
        strings["xl_col_ref"],
        strings["xl_col_company"],
        strings["xl_col_site"],
        strings["xl_col_period"],
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
        detail_headers += [
            f"{prefix}: {strings['xl_tariff_name']}",
            f"{prefix}: {strings['xl_tariff_qty']}",
            f"{prefix}: {strings['xl_tariff_rate']}",
            f"{prefix}: {strings['xl_tariff_cost']}",
        ]

    for col_idx, header in enumerate(detail_headers, start=1):
        cell = detail.cell(1, col_idx, header)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for row_idx, section in enumerate(sections, start=2):
        company = section["company"]
        site = section["site"]
        row_values = [
            site["ref_no"],
            company["label"],
            site["label"],
            site["billing_period_label"],
            site["city"],
            site["postcode"],
            site["meter_id"],
            float(site["supplier_ef"]),
            site["unit"],
            site["start_reading"],
            site["end_reading"],
            float(site["total_quantity"]),
            float(site["total_cost"]),
            currency_code(company.get("currency")),
            float(site["emissions_kg"]),
            float(site["emissions_t"]),
        ]
        tariffs = site.get("tariffs", [])
        for idx in range(max_tariffs):
            if idx < len(tariffs):
                tariff = tariffs[idx]
                row_values += [
                    tariff["name"],
                    float(tariff["quantity"]),
                    float(tariff["unit_cost"]),
                    float(tariff["cost"]),
                ]
            else:
                row_values += ["", "", "", ""]

        numeric_cols = {8, 12, 13, 15, 16}
        rate_cols = set()
        for idx in range(max_tariffs):
            base = 17 + idx * 4
            numeric_cols |= {base + 1, base + 2, base + 3}
            rate_cols.add(base + 2)

        for col_idx, value in enumerate(row_values, start=1):
            cell = detail.cell(row_idx, col_idx, value)
            cell.border = border
            cell.font = Font(name="Calibri", size=9)
            if col_idx in numeric_cols and isinstance(value, float):
                cell.number_format = "#,##0.0000" if col_idx in rate_cols else "#,##0.00"

    base_widths = [30, 28, 22, 18, 16, 10, 24, 16, 8, 14, 14, 14, 14, 10, 16, 16]
    detail_widths = base_widths + [26, 12, 12, 12] * max_tariffs
    for col_idx, width in enumerate(detail_widths, start=1):
        detail.column_dimensions[get_column_letter(col_idx)].width = width

    buf = BytesIO()
    workbook.save(buf)
    return buf.getvalue()


def _generate_smart_meter_xlsx(config: dict, sections: list[dict]) -> bytes:
    from generators.electricity_generator import build_smart_meter_rows

    lang = config["document"].get("language", "en")
    strings = ELECTRICITY_TRANSLATIONS.get(lang, ELECTRICITY_TRANSLATIONS["en"])
    mode = str(config["document"].get("smart_meter_data_granularity", "monthly")).lower()
    value_mode = str(config["document"].get("smart_meter_interval_value_mode", "consumption_diff")).lower()
    rows = build_smart_meter_rows(config, sections)

    accent = config["companies"][0].get("accent", "#1E5B88") if config.get("companies") else "#1E5B88"
    accent_fill = accent.lstrip("#").upper()
    if len(accent_fill) == 6:
        accent_fill = f"FF{accent_fill}"

    workbook = openpyxl.Workbook()
    ws = workbook.active
    ws.title = "Smart Meter Data"
    ws.freeze_panes = "A2"

    if mode == "interval":
        if value_mode == "cumulative_end_reading":
            headers = [
                strings["xl_col_meter_id"],
                strings["sm_col_timestamp"],
                strings["sm_col_end_reading"],
                strings["xl_col_unit"],
            ]
            widths = [22, 24, 16, 10]
        else:
            headers = [
                strings["xl_col_meter_id"],
                strings["sm_col_timestamp"],
                strings["sm_col_import_kwh"],
                strings["sm_col_export_kwh"],
                strings["xl_col_unit"],
            ]
            widths = [22, 24, 14, 14, 10]
    else:
        headers = [
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
        ]
        widths = [22, 24, 18, 14, 14, 14, 10, 18, 12, 10]

    for col_idx, header in enumerate(headers, start=1):
        _header_cell(ws.cell(row=1, column=col_idx), header, accent_fill)
        ws.column_dimensions[get_column_letter(col_idx)].width = widths[col_idx - 1]

    for row_idx, row in enumerate(rows, start=2):
        alt = row_idx % 2 == 0
        if mode == "interval":
            if value_mode == "cumulative_end_reading":
                row_values = [
                    (row["meter_id"], None),
                    (row["timestamp"], None),
                    (row["end_reading"], "0.0000"),
                    (row["unit"], None),
                ]
            else:
                row_values = [
                    (row["meter_id"], None),
                    (row["timestamp"], None),
                    (row["import_kwh"], "0.0000"),
                    (row["export_kwh"], "0.0000"),
                    (row["unit"], None),
                ]
        else:
            row_values = [
                (row["meter_id"], None),
                (row["site_label"], None),
                (row["period_label"], None),
                (row["start_reading"], "#,##0"),
                (row["end_reading"], "#,##0"),
                (row["consumption"], "#,##0.00"),
                (row["unit"], None),
                (row["tariff_type"], None),
                (row["cost"], "#,##0.00"),
                (row["currency"], None),
            ]

        for col_idx, (value, fmt) in enumerate(row_values, start=1):
            _data_cell(ws.cell(row=row_idx, column=col_idx), value, fmt=fmt, alt=alt)

    buf = BytesIO()
    workbook.save(buf)
    return buf.getvalue()


def _generate_heat_supplier_portal_xlsx(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
    split_by_company: bool = False,
    include_summary: bool = True,
) -> bytes:
    return _generate_heat_xlsx(
        config,
        sections,
        blank_fields=blank_fields,
        split_by_company=split_by_company,
        include_summary=include_summary,
    )


def _generate_electricity_supplier_portal_xlsx(
    config: dict,
    sections: list[dict],
    include_summary: bool = True,
) -> bytes:
    return _generate_electricity_xlsx(config, sections, include_summary=include_summary)


# ── summary sheet ─────────────────────────────────────────────────────────────

def _build_summary_sheet(ws, config: dict, sections: list[dict], accent: str, strings: dict) -> None:
    # Title banner
    ws.merge_cells("A1:I1")
    ws["A1"].value = config["document"]["title"]
    ws["A1"].font = _font(bold=True, color=_WHITE, size=14)
    ws["A1"].fill = _fill(accent)
    ws["A1"].alignment = _align("center")
    ws.row_dimensions[1].height = 30

    # Metadata
    meta = [
        (strings["meta_period"],    config["financial_period"]["label"]),
        (strings["meta_start"],     config["financial_period"]["start_date"].strftime("%d %b %Y")),
        (strings["meta_end"],       config["financial_period"]["end_date"].strftime("%d %b %Y")),
        (strings["meta_generated"], datetime.now().strftime("%d %b %Y %H:%M")),
    ]
    for offset, (label, value) in enumerate(meta):
        row = 2 + offset
        ws.cell(row=row, column=1, value=label).font = _font(bold=True, color=_MID)
        ws.cell(row=row, column=2, value=value).font = _font()

    # Company summary table
    tbl_row = 2 + len(meta) + 1
    summary_headers = [
        strings["sum_company"], strings["sum_sites"], strings["sum_invoices"],
        strings["sum_heat_cost"], strings["sum_cap_charge"],
        strings["sum_subtotal"], strings["sum_vat"],     strings["sum_total"], strings["sum_currency"],
    ]
    for col, h in enumerate(summary_headers, start=1):
        _header_cell(ws.cell(row=tbl_row, column=col), h, accent)

    # Aggregate per company
    totals: dict[str, dict] = {}
    for section in sections:
        key = section["company"]["label"]
        if key not in totals:
            totals[key] = {
                "sites": set(),
                "currency": currency_code(section["company"].get("currency")),
                "invoices": 0,
                "heat_cost": Decimal("0"),
                "capacity_charge": Decimal("0"),
                "subtotal": Decimal("0"),
                "vat": Decimal("0"),
                "total": Decimal("0"),
            }
        t = totals[key]
        t["sites"].add(section["site"]["label"])
        for rec in section["records"]:
            t["invoices"] += 1
            t["heat_cost"] += rec["heat_cost"]
            t["capacity_charge"] += rec["capacity_charge"]
            t["subtotal"] += rec["subtotal"]
            t["vat"] += rec["vat"]
            t["total"] += rec["total"]

    money_fmt = "#,##0.00"
    for i, (company, t) in enumerate(totals.items()):
        row = tbl_row + 1 + i
        alt = i % 2 == 1
        row_values = [
            (company,               None),
            (len(t["sites"]),       "#,##0"),
            (t["invoices"],         "#,##0"),
            (float(t["heat_cost"]), money_fmt),
            (float(t["capacity_charge"]), money_fmt),
            (float(t["subtotal"]),  money_fmt),
            (float(t["vat"]),       money_fmt),
            (float(t["total"]),     money_fmt),
            (t["currency"],         None),
        ]
        for col, (val, fmt) in enumerate(row_values, start=1):
            _data_cell(ws.cell(row=row, column=col), val, fmt=fmt, alt=alt)

    # Grand total row
    grand_row = tbl_row + 1 + len(totals)
    grand_values = [
        (strings["sum_grand"], None),
        ("",      None),
        (sum(t["invoices"] for t in totals.values()), "#,##0"),
        (float(sum(t["heat_cost"]        for t in totals.values())), money_fmt),
        (float(sum(t["capacity_charge"]  for t in totals.values())), money_fmt),
        (float(sum(t["subtotal"]         for t in totals.values())), money_fmt),
        (float(sum(t["vat"]              for t in totals.values())), money_fmt),
        (float(sum(t["total"]            for t in totals.values())), money_fmt),
        ("",      None),
    ]
    for col, (val, fmt) in enumerate(grand_values, start=1):
        c = ws.cell(row=grand_row, column=col, value=val)
        c.font = _font(bold=True)
        c.fill = _fill(_SOFT)
        c.border = _border()
        c.alignment = _align("right" if col > 2 else "left")
        if fmt:
            c.number_format = fmt

    # Column widths
    for col, width in enumerate([30, 8, 10, 20, 22, 18, 14, 20, 10], start=1):
        ws.column_dimensions[get_column_letter(col)].width = width

    ws.freeze_panes = "A2"


# ── detail sheet ──────────────────────────────────────────────────────────────

def _build_detail_sheet(ws, sections: list[dict], accent: str, blank_fields: set[str], strings: dict) -> None:
    # Headers
    for col, (label, width, _) in enumerate(_detail_cols(strings), start=1):
        _header_cell(ws.cell(row=1, column=col), label, accent)
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    # Each entry: (value, record_field_name_or_None, fmt)
    # record_field_name is checked against blank_fields to omit the cell value.
    row = 2
    for section in sections:
        company = section["company"]
        site = section["site"]
        for rec in section["records"]:
            alt = row % 2 == 0
            row_data: list[tuple[Any, str | None, str | None]] = [
                (rec["invoice_no"],          "invoice_no",    None),
                (company["label"],           "company_label", None),
                (site["label"],              "site_label",    None),
                (rec["city"],                "city",          None),
                (rec["postcode"],            "postcode",      None),
                (rec["meter_id"],            "meter_id",      None),
                (rec["billing_period_label"],"period_label",  None),
                (rec["period_start"],        None,            "DD MMM YYYY"),
                (rec["period_end"],          None,            "DD MMM YYYY"),
                (rec["issue_date"],          None,            "DD MMM YYYY"),
                (rec["due_date"],            None,            "DD MMM YYYY"),
                (rec["prev_read"],           "prev_read",     "#,##0"),
                (rec["curr_read"],           "curr_read",     "#,##0"),
                (rec["consumption"],         "consumption",   "#,##0"),
                (float(rec["unit_price"]),   "unit_price",    "0.000"),
                (float(rec["heat_cost"]),    "heat_cost",     "#,##0.00"),
                (rec["capacity_kw"],         "capacity_kw",   "#,##0"),
                (float(rec["capacity_rate"]),"capacity_rate", "0.00"),
                (float(rec["supplier_ef"]),  "supplier_ef",   "0.0000"),
                (float(rec["capacity_charge"]), "capacity_charge", "#,##0.00"),
                (float(rec["subtotal"]),     "subtotal",      "#,##0.00"),
                (float(rec["vat"]),          "vat",           "#,##0.00"),
                (float(rec["total"]),        "total",         "#,##0.00"),
                (currency_code(company.get("currency")), "currency", None),
            ]
            for col, (value, field_name, fmt) in enumerate(row_data, start=1):
                cell_value = None if (field_name and field_name in blank_fields) else value
                _data_cell(ws.cell(row=row, column=col), cell_value, fmt=fmt, alt=alt)
            ws.row_dimensions[row].height = 18
            row += 1


def generate_xlsx(
    config: dict,
    sections: list[dict],
    blank_fields: set[str] | None = None,
    split_by_company: bool = False,
    include_summary: bool = True,
    category: str = "heat",
) -> bytes:
    document_type = config["document"].get("type")
    if category == "electricity":
        if document_type == "smart_meter_data":
            return _generate_smart_meter_xlsx(config, sections)
        if document_type == "supplier_portal_data":
            return _generate_electricity_supplier_portal_xlsx(config, sections, include_summary=include_summary)
        raise NotImplementedError(f"XLSX generation is not supported for electricity document type '{document_type}'.")
    if document_type == "supplier_portal_data":
        return _generate_heat_supplier_portal_xlsx(
            config,
            sections,
            blank_fields=blank_fields,
            split_by_company=split_by_company,
            include_summary=include_summary,
        )
    raise NotImplementedError(f"XLSX generation is not supported for heat document type '{document_type}'.")
