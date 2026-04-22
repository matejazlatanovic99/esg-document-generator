from __future__ import annotations

import csv
import random
from calendar import monthrange
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from io import BytesIO, StringIO

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from components.stationary_combustion.units import default_fuel_volume_unit

TWOPLACES = Decimal("0.01")
PAGE_W, PAGE_H = A4
_SPECIAL_CHARS_SUFFIX = ' & < " £ € \u00a0\u2014\u200f'

STATIONARY_TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "fuel_invoice_title": "Fuel Invoice",
        "bill_to": "Bill To",
        "invoice_details": "Invoice Details",
        "delivery_site": "Delivery Site",
        "invoice_no": "Invoice No",
        "invoice_date": "Invoice Date",
        "billing_period": "Billing Period",
        "currency": "Currency",
        "country": "Country",
        "product": "Product",
        "quantity": "Quantity",
        "unit": "Unit",
        "unit_price": "Unit Price",
        "amount": "Amount",
        "delivery_charge": "Delivery Charge",
        "each": "Each",
        "subtotal": "Subtotal",
        "total": "Total",
        "generator_operation_log_title": "Generator Operation Log",
        "litres": "Litres",
        "received": "Received",
        "stationary_equipment": "Stationary Equipment",
        "test_run": "Test run",
        "power_outage": "Power outage",
        "maintenance_test": "Maintenance test",
        "load_bank_test": "Load bank test",
        "fuel_invoice_footer": "Generated for Scope 1 stationary combustion. Delivery-site details are illustrative.",
        "delivery_note_title": "Fuel Delivery Note",
        "supplier": "Supplier",
        "delivery_note_no": "Delivery Note No",
        "delivery_date": "Delivery Date",
        "customer": "Customer",
        "delivery_address": "Delivery Address",
        "delivery_details": "Delivery Details",
        "delivery_confirmation": "Delivery Confirmation",
        "product_delivered": "Product Delivered",
        "tank_equipment": "Tank / Equipment",
        "delivered_quantity": "Delivered Quantity",
        "driver_ref": "Driver Ref",
        "customer_signature": "Customer Signature",
        "delivery_note_footer": "Generated for Scope 1 stationary combustion delivery-note testing.",
        "fuel_card_title": "Fuel Card Statement",
        "account_name": "Account Name",
        "provider": "Provider",
        "statement_period": "Statement Period",
        "card_no": "Card No",
        "date": "Date",
        "merchant": "Merchant",
        "reference": "Reference",
        "qty": "Qty",
        "statement_total": "Statement Total",
        "fuel_card_footer": "Statement generated for stationary-equipment fuel-card QA.",
        "account_details": "Account Details",
        "statement_details": "Statement Details",
        "company": "Company",
        "site": "Site",
        "equipment": "Equipment",
        "emission_source": "Emission Source",
        "generator_log_sheet_title": "Generator Log",
        "period": "Period",
        "start_time": "Start Time",
        "end_time": "End Time",
        "run_hours": "Run Hours",
        "start_fuel": "Start Fuel",
        "end_fuel": "End Fuel",
        "fuel_used": "Fuel Used",
        "notes": "Notes",
        "bems_equipment_title": "BEMS Fuel Consumption Summary",
        "assets": "Assets",
        "operating_hours": "Operating Hours",
        "top_asset": "Top Asset",
        "equipment_trend_snapshot": "Equipment Trend Snapshot",
        "equipment_tag": "Equipment Tag",
        "equipment_name": "Equipment Name",
        "fuel_type": "Fuel Type",
        "consumption": "Consumption",
        "dashboard_summary_footer": "Dashboard summary generated from BEMS trend data.",
        "bems_summary_sheet_title": "BEMS Summary",
        "reporting_period": "Reporting Period",
        "bems_time_series_title": "BEMS Time-Series Trend Export",
        "interval": "Interval",
        "rows": "Rows",
        "timestamp": "Timestamp",
        "sensor_name": "Sensor Name",
        "value": "Value",
        "time_series_footer": "Time-series export rendered as PDF preview.",
        "time_series_word_footer": "Time-series export rendered as Word preview.",
        "site_fallback": "Site",
    },
    "fr": {
        "fuel_invoice_title": "Facture de carburant",
        "bill_to": "Facturer a",
        "invoice_details": "Details de facture",
        "delivery_site": "Site de livraison",
        "invoice_no": "No de facture",
        "invoice_date": "Date de facture",
        "billing_period": "Periode de facturation",
        "currency": "Devise",
        "country": "Pays",
        "product": "Produit",
        "quantity": "Quantite",
        "unit": "Unite",
        "unit_price": "Prix unitaire",
        "amount": "Montant",
        "delivery_charge": "Frais de livraison",
        "each": "Chaque",
        "subtotal": "Sous-total",
        "total": "Total",
        "generator_operation_log_title": "Journal d'exploitation generateur",
        "litres": "Litres",
        "received": "Recu",
        "stationary_equipment": "Equipement stationnaire",
        "test_run": "Essai",
        "power_outage": "Coupure de courant",
        "maintenance_test": "Essai de maintenance",
        "load_bank_test": "Essai au banc de charge",
        "fuel_invoice_footer": "Genere pour la combustion stationnaire du scope 1. Les details du site de livraison sont indicatifs.",
        "delivery_note_title": "Bon de livraison carburant",
        "supplier": "Fournisseur",
        "delivery_note_no": "No de bon de livraison",
        "delivery_date": "Date de livraison",
        "customer": "Client",
        "delivery_address": "Adresse de livraison",
        "delivery_details": "Details de livraison",
        "delivery_confirmation": "Confirmation de livraison",
        "product_delivered": "Produit livre",
        "tank_equipment": "Cuve / Equipement",
        "delivered_quantity": "Quantite livree",
        "driver_ref": "Ref chauffeur",
        "customer_signature": "Signature client",
        "delivery_note_footer": "Genere pour les tests de bons de livraison de combustion stationnaire du scope 1.",
        "fuel_card_title": "Releve de carte carburant",
        "account_name": "Nom du compte",
        "provider": "Fournisseur",
        "statement_period": "Periode du releve",
        "card_no": "No carte",
        "date": "Date",
        "merchant": "Commercant",
        "reference": "Reference",
        "qty": "Qt",
        "statement_total": "Total du releve",
        "fuel_card_footer": "Releve genere pour l'assurance qualite des cartes carburant d'equipements stationnaires.",
        "account_details": "Details du compte",
        "statement_details": "Details du releve",
        "company": "Entreprise",
        "site": "Site",
        "equipment": "Equipement",
        "emission_source": "Source d'emission",
        "generator_log_sheet_title": "Journal generateur",
        "period": "Periode",
        "start_time": "Heure debut",
        "end_time": "Heure fin",
        "run_hours": "Heures de marche",
        "start_fuel": "Carburant debut",
        "end_fuel": "Carburant fin",
        "fuel_used": "Carburant utilise",
        "notes": "Notes",
        "bems_equipment_title": "Resume BEMS de consommation de carburant",
        "assets": "Actifs",
        "operating_hours": "Heures de fonctionnement",
        "top_asset": "Actif principal",
        "equipment_trend_snapshot": "Apercu de tendance des equipements",
        "equipment_tag": "Identifiant equipement",
        "equipment_name": "Nom equipement",
        "fuel_type": "Type de carburant",
        "consumption": "Consommation",
        "dashboard_summary_footer": "Resume du tableau de bord genere a partir des donnees de tendance BEMS.",
        "bems_summary_sheet_title": "Resume BEMS",
        "reporting_period": "Periode de reporting",
        "bems_time_series_title": "Export BEMS de tendance chronologique",
        "interval": "Intervalle",
        "rows": "Lignes",
        "timestamp": "Horodatage",
        "sensor_name": "Nom du capteur",
        "value": "Valeur",
        "time_series_footer": "Export chronologique restitue comme apercu PDF.",
        "time_series_word_footer": "Export chronologique restitue comme apercu Word.",
        "site_fallback": "Site",
    },
    "de": {
        "fuel_invoice_title": "Kraftstoffrechnung",
        "bill_to": "Rechnung an",
        "invoice_details": "Rechnungsdetails",
        "delivery_site": "Lieferstandort",
        "invoice_no": "Rechnungsnr.",
        "invoice_date": "Rechnungsdatum",
        "billing_period": "Abrechnungszeitraum",
        "currency": "Wahrung",
        "country": "Land",
        "product": "Produkt",
        "quantity": "Menge",
        "unit": "Einheit",
        "unit_price": "Stuckpreis",
        "amount": "Betrag",
        "delivery_charge": "Lieferkosten",
        "each": "Je",
        "subtotal": "Zwischensumme",
        "total": "Gesamt",
        "generator_operation_log_title": "Generatorbetriebsprotokoll",
        "litres": "Liter",
        "received": "Erhalten",
        "stationary_equipment": "Stationare Anlage",
        "test_run": "Testlauf",
        "power_outage": "Stromausfall",
        "maintenance_test": "Wartungstest",
        "load_bank_test": "Lastbanktest",
        "fuel_invoice_footer": "Erstellt fur Scope-1-Standverbrennung. Lieferstandortdetails dienen nur zur Veranschaulichung.",
        "delivery_note_title": "Kraftstofflieferschein",
        "supplier": "Lieferant",
        "delivery_note_no": "Lieferscheinnr.",
        "delivery_date": "Lieferdatum",
        "customer": "Kunde",
        "delivery_address": "Lieferadresse",
        "delivery_details": "Lieferdetails",
        "delivery_confirmation": "Lieferbestatigung",
        "product_delivered": "Geliefertes Produkt",
        "tank_equipment": "Tank / Anlage",
        "delivered_quantity": "Gelieferte Menge",
        "driver_ref": "Fahrer-Ref",
        "customer_signature": "Kundenunterschrift",
        "delivery_note_footer": "Erstellt fur Tests von Lieferscheinen zur stationaren Verbrennung in Scope 1.",
        "fuel_card_title": "Tankkartenabrechnung",
        "account_name": "Kontoname",
        "provider": "Anbieter",
        "statement_period": "Abrechnungszeitraum",
        "card_no": "Kartennr.",
        "date": "Datum",
        "merchant": "Handler",
        "reference": "Referenz",
        "qty": "Menge",
        "statement_total": "Abrechnungssumme",
        "fuel_card_footer": "Abrechnung fur die QS stationarer Tankkartenvorgange erstellt.",
        "account_details": "Kontodetails",
        "statement_details": "Abrechnungsdetails",
        "company": "Unternehmen",
        "site": "Standort",
        "equipment": "Anlage",
        "emission_source": "Emissionsquelle",
        "generator_log_sheet_title": "Generatorprotokoll",
        "period": "Zeitraum",
        "start_time": "Startzeit",
        "end_time": "Endzeit",
        "run_hours": "Betriebsstunden",
        "start_fuel": "Kraftstoff Start",
        "end_fuel": "Kraftstoff Ende",
        "fuel_used": "Verbrauchter Kraftstoff",
        "notes": "Hinweise",
        "bems_equipment_title": "BEMS-Kraftstoffverbrauchsbericht",
        "assets": "Anlagen",
        "operating_hours": "Betriebsstunden",
        "top_asset": "Top-Anlage",
        "equipment_trend_snapshot": "Anlagen-Trendubersicht",
        "equipment_tag": "Anlagenkennzeichen",
        "equipment_name": "Anlagenname",
        "fuel_type": "Kraftstoffart",
        "consumption": "Verbrauch",
        "dashboard_summary_footer": "Dashboard-Zusammenfassung aus BEMS-Trenddaten erstellt.",
        "bems_summary_sheet_title": "BEMS Ubersicht",
        "reporting_period": "Berichtszeitraum",
        "bems_time_series_title": "BEMS-Zeitreihenexport",
        "interval": "Intervall",
        "rows": "Zeilen",
        "timestamp": "Zeitstempel",
        "sensor_name": "Sensorname",
        "value": "Wert",
        "time_series_footer": "Zeitreihenexport als PDF-Vorschau dargestellt.",
        "time_series_word_footer": "Zeitreihenexport als Word-Vorschau dargestellt.",
        "site_fallback": "Standort",
    },
    "nl": {
        "fuel_invoice_title": "Brandstoffactuur",
        "bill_to": "Factureren aan",
        "invoice_details": "Factuurgegevens",
        "delivery_site": "Leveringslocatie",
        "invoice_no": "Factuurnr.",
        "invoice_date": "Factuurdatum",
        "billing_period": "Facturatieperiode",
        "currency": "Valuta",
        "country": "Land",
        "product": "Product",
        "quantity": "Hoeveelheid",
        "unit": "Eenheid",
        "unit_price": "Eenheidsprijs",
        "amount": "Bedrag",
        "delivery_charge": "Leveringskosten",
        "each": "Per stuk",
        "subtotal": "Subtotaal",
        "total": "Totaal",
        "generator_operation_log_title": "Generatorbedrijfslog",
        "litres": "Liter",
        "received": "Ontvangen",
        "stationary_equipment": "Stationaire installatie",
        "test_run": "Testrun",
        "power_outage": "Stroomstoring",
        "maintenance_test": "Onderhoudstest",
        "load_bank_test": "Belastingbanktest",
        "fuel_invoice_footer": "Gegenereerd voor Scope 1 stationaire verbranding. Details van de leveringslocatie zijn illustratief.",
        "delivery_note_title": "Brandstofleverbon",
        "supplier": "Leverancier",
        "delivery_note_no": "Leverbonnr.",
        "delivery_date": "Leverdatum",
        "customer": "Klant",
        "delivery_address": "Leveradres",
        "delivery_details": "Leveringsdetails",
        "delivery_confirmation": "Leveringsbevestiging",
        "product_delivered": "Geleverd product",
        "tank_equipment": "Tank / Installatie",
        "delivered_quantity": "Geleverde hoeveelheid",
        "driver_ref": "Chauffeursref",
        "customer_signature": "Handtekening klant",
        "delivery_note_footer": "Gegenereerd voor het testen van leverbonnen voor stationaire verbranding in Scope 1.",
        "fuel_card_title": "Tankkaartoverzicht",
        "account_name": "Accountnaam",
        "provider": "Aanbieder",
        "statement_period": "Overzichtsperiode",
        "card_no": "Kaartnr.",
        "date": "Datum",
        "merchant": "Leverancier",
        "reference": "Referentie",
        "qty": "Aantal",
        "statement_total": "Totaal overzicht",
        "fuel_card_footer": "Overzicht gegenereerd voor QA van tankkaarttransacties voor stationaire apparatuur.",
        "account_details": "Accountgegevens",
        "statement_details": "Overzichtsgegevens",
        "company": "Bedrijf",
        "site": "Locatie",
        "equipment": "Installatie",
        "emission_source": "Emissiebron",
        "generator_log_sheet_title": "Generatorlog",
        "period": "Periode",
        "start_time": "Starttijd",
        "end_time": "Eindtijd",
        "run_hours": "Draaiuren",
        "start_fuel": "Brandstof start",
        "end_fuel": "Brandstof eind",
        "fuel_used": "Verbruikte brandstof",
        "notes": "Notities",
        "bems_equipment_title": "BEMS-brandstofverbruiksoverzicht",
        "assets": "Assets",
        "operating_hours": "Bedrijfsuren",
        "top_asset": "Belangrijkste asset",
        "equipment_trend_snapshot": "Momentopname apparatuurtendens",
        "equipment_tag": "Apparaatcode",
        "equipment_name": "Apparaatnaam",
        "fuel_type": "Brandstoftype",
        "consumption": "Verbruik",
        "dashboard_summary_footer": "Dashboardsamenvatting gegenereerd uit BEMS-trendgegevens.",
        "bems_summary_sheet_title": "BEMS-overzicht",
        "reporting_period": "Rapportageperiode",
        "bems_time_series_title": "BEMS-tijdreeks export",
        "interval": "Interval",
        "rows": "Rijen",
        "timestamp": "Tijdstempel",
        "sensor_name": "Sensornaam",
        "value": "Waarde",
        "time_series_footer": "Tijdreeks-export weergegeven als PDF-voorbeeld.",
        "time_series_word_footer": "Tijdreeks-export weergegeven als Word-voorbeeld.",
        "site_fallback": "Locatie",
    },
}


def _q2(value) -> Decimal:
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    return value.quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def _parse_decimal(value, fallback: str = "0") -> Decimal:
    if value in (None, ""):
        return Decimal(fallback)
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def _parse_date(value: str) -> date:
    return datetime.strptime(value, "%Y-%m-%d").date()


def _language(raw_config: dict) -> str:
    language = str(raw_config.get("document", {}).get("language", "en")).lower()
    return language if language in STATIONARY_TRANSLATIONS else "en"


def _tr(raw_config: dict, key: str, **kwargs) -> str:
    template = STATIONARY_TRANSLATIONS[_language(raw_config)][key]
    return template.format(**kwargs) if kwargs else template


def _fmt_date(value: date) -> str:
    return value.strftime("%d %b %Y")


def _fmt_money(value) -> str:
    return f"{_q2(value):,.2f}"


def _fmt_optional_number(value, suffix: str = "") -> str:
    if value in (None, ""):
        return ""
    return f"{_fmt_money(value)}{suffix}"


def _currency_symbol(currency_raw: str) -> str:
    mapping = {
        "(£)": "£",
        "(€)": "€",
        "($)": "$",
        "(¥)": "¥",
        "(kr)": "kr",
        "(Ft)": "Ft",
    }
    for token, symbol in mapping.items():
        if token in currency_raw:
            return symbol
    return ""


def _with_special_chars(config: dict, value: str) -> str:
    if not value:
        return value
    if not config.get("document", {}).get("inject_special_chars", False):
        return value
    return value + _SPECIAL_CHARS_SUFFIX


def _financial_period(raw_config: dict) -> dict:
    fp = raw_config.get("financial_period", {})
    return {
        "label": fp.get("label", ""),
        "start_date": _parse_date(fp.get("start_date", "2026-01-01")),
        "end_date": _parse_date(fp.get("end_date", "2026-01-31")),
    }


def _months_in_range(start_date: date, end_date: date) -> list[tuple[int, int]]:
    months: list[tuple[int, int]] = []
    current = date(start_date.year, start_date.month, 1)
    while current <= end_date:
        months.append((current.year, current.month))
        if current.month == 12:
            current = date(current.year + 1, 1, 1)
        else:
            current = date(current.year, current.month + 1, 1)
    return months


def _bems_interval_minutes(raw_config: dict) -> int:
    minutes = int(raw_config.get("document", {}).get("bems_interval_minutes", 60))
    return minutes if minutes in {15, 30, 60} else 60


def _bems_report_type(raw_config: dict) -> str:
    report_type = str(raw_config.get("document", {}).get("bems_report_type", "equipment_trend_report"))
    return report_type if report_type in {"equipment_trend_report", "time_series_trend_export"} else "equipment_trend_report"


def _timestamp_range(start_date: date, end_date: date, interval_minutes: int) -> list[datetime]:
    timestamps: list[datetime] = []
    current = datetime.combine(start_date, datetime.min.time())
    end_dt = datetime.combine(end_date + timedelta(days=1), datetime.min.time())
    while current < end_dt:
        timestamps.append(current)
        current += timedelta(minutes=interval_minutes)
    return timestamps


def _iter_company_sites(raw_config: dict):
    for company_index, company in enumerate(raw_config.get("companies", []), start=1):
        for site_index, site in enumerate(company.get("sites", []), start=1):
            yield company_index, site_index, company, site


def _site_equipment_items(raw_config: dict, site: dict, *, include_emission_source: bool) -> list[dict]:
    site_omit = site.get("_omit", {})
    equipment_omitted = bool(site_omit.get("equipment", False))
    site_emission_source_omitted = bool(site_omit.get("emission_source", False))
    raw_items = site.get("equipment_items")

    if equipment_omitted:
        raw_items = raw_items if isinstance(raw_items, list) and raw_items else []
        first_item = raw_items[0] if raw_items else {}
        if isinstance(first_item, dict):
            first_emission_source = first_item.get("emission_source", site.get("emission_source", ""))
            first_omit = first_item.get("_omit", {})
        else:
            first_emission_source = site.get("emission_source", "")
            first_omit = {}
        raw_item = dict(first_item) if isinstance(first_item, dict) else {}
        raw_item.update({
            "equipment": "",
            "emission_source": first_emission_source,
            "_omit": first_omit,
        })
        raw_items = [raw_item]

    if not isinstance(raw_items, list) or not raw_items:
        raw_items = [
            {
                "equipment": site.get("equipment", ""),
                "emission_source": site.get("emission_source", ""),
                "_omit": {"emission_source": site_emission_source_omitted},
            }
        ]

    equipment_items: list[dict] = []
    for raw_item in raw_items:
        if isinstance(raw_item, str):
            equipment = raw_item
            emission_source = site.get("emission_source", "")
            item_data: dict = {}
            item_omit: dict = {}
        else:
            equipment = raw_item.get("equipment", raw_item.get("name", ""))
            emission_source = raw_item.get("emission_source", site.get("emission_source", ""))
            item_data = raw_item
            item_omit = raw_item.get("_omit", {})

        emission_source_omitted = site_emission_source_omitted or bool(item_omit.get("emission_source", False))
        equipment_item_omitted = equipment_omitted or bool(item_omit.get("equipment", False))
        normalized_item = {
            "equipment": "" if equipment_item_omitted else _with_special_chars(raw_config, "" if equipment is None else str(equipment)),
            "emission_source": ""
            if not include_emission_source or emission_source_omitted
            else _with_special_chars(raw_config, "" if emission_source is None else str(emission_source)),
        }
        for field in [
            "fuel",
            "unit",
            "quantity",
            "unit_price",
            "delivery_charge",
            "vat_rate",
            "runs_per_month",
            "fuel_used_per_hour",
            "quantity_mode",
            "tank_capacity",
            "run_hours_min",
            "run_hours_max",
        ]:
            if field in item_data:
                normalized_item[field] = item_data.get(field)
            elif field in site:
                normalized_item[field] = site.get(field)
        normalized_item.setdefault("unit", site.get("unit", default_fuel_volume_unit(raw_config.get("document_type"))))
        normalized_item.setdefault("fuel", site.get("fuel", ""))
        equipment_items.append(normalized_item)

    return equipment_items or [{"equipment": "", "emission_source": ""}]


def _build_fuel_invoice_records(raw_config: dict) -> list[dict]:
    fp = _financial_period(raw_config)
    records: list[dict] = []
    seed = int(raw_config.get("random_seed", 42))

    for company_index, site_index, company, site in _iter_company_sites(raw_config):
        site_omit = site.get("_omit", {})
        equipment_items = _site_equipment_items(raw_config, site, include_emission_source=True)

        for equipment_index, equipment_item in enumerate(equipment_items, start=1):
            quantity = _q2(_parse_decimal(equipment_item.get("quantity"), "0"))
            unit_price = _q2(_parse_decimal(equipment_item.get("unit_price"), "0"))
            delivery_charge = _q2(_parse_decimal(equipment_item.get("delivery_charge"), "0"))
            vat_rate = _parse_decimal(equipment_item.get("vat_rate"), "20")
            fuel_cost = _q2(quantity * unit_price)
            subtotal = _q2(fuel_cost + delivery_charge)
            vat = _q2(subtotal * vat_rate / Decimal("100"))
            total = _q2(subtotal + vat)
            rng = random.Random(f"{seed}:fuel_invoice:{company_index}:{site_index}:{equipment_index}")
            invoice_date = fp["end_date"] + timedelta(days=rng.randint(2, 8))
            invoice_suffix = f"{company_index:02d}{site_index:02d}"
            if len(equipment_items) > 1:
                invoice_suffix = f"{invoice_suffix}-{equipment_index:02d}"
            invoice_no = (
                f"{company.get('supplier_code', 'INV')}-{fp['start_date'].strftime('%Y%m')}"
                f"-{invoice_suffix}"
            )

            record = {
                "company": _with_special_chars(raw_config, company.get("label", "")),
                "supplier": _with_special_chars(raw_config, company.get("supplier", "")),
                "supplier_address": [
                    _with_special_chars(raw_config, line) for line in company.get("supplier_address", [])
                ],
                "customer": _with_special_chars(raw_config, company.get("customer", "")),
                "customer_code": company.get("customer_code", ""),
                "site": _with_special_chars(raw_config, site.get("label", "")),
                "site_address": [_with_special_chars(raw_config, line) for line in site.get("customer_address", [])],
                "country": "" if site_omit.get("country", False) else _with_special_chars(raw_config, site.get("country", "")),
                "equipment": equipment_item["equipment"],
                "emission_source": equipment_item["emission_source"],
                "fuel": _with_special_chars(raw_config, equipment_item.get("fuel", "")),
                "unit": equipment_item.get("unit", site.get("unit", _tr(raw_config, "litres"))),
                "quantity": quantity,
                "unit_price": unit_price,
                "fuel_cost": fuel_cost,
                "delivery_charge": delivery_charge,
                "subtotal": subtotal,
                "vat_rate": vat_rate,
                "vat": vat,
                "total": total,
                "currency": company.get("currency", "GBP (£)"),
                "invoice_no": invoice_no,
                "invoice_date": invoice_date,
                "period_label": fp["label"],
                "period_start": fp["start_date"],
                "period_end": fp["end_date"],
            }
            records.append(record)
    return records


def _build_delivery_note_records(raw_config: dict) -> list[dict]:
    fp = _financial_period(raw_config)
    records: list[dict] = []
    seed = int(raw_config.get("random_seed", 42))
    days_in_period = max((fp["end_date"] - fp["start_date"]).days, 0)

    for company_index, site_index, company, site in _iter_company_sites(raw_config):
        site_omit = site.get("_omit", {})
        equipment_items = _site_equipment_items(raw_config, site, include_emission_source=False)

        for equipment_index, equipment_item in enumerate(equipment_items, start=1):
            rng = random.Random(f"{seed}:delivery_note:{company_index}:{site_index}:{equipment_index}")
            delivery_date = fp["start_date"] + timedelta(days=rng.randint(0, days_in_period))
            note_no = (
                f"DN-{delivery_date.strftime('%Y')}-"
                f"{rng.randint(10000, 99999)}"
            )

            records.append({
                "company": _with_special_chars(raw_config, company.get("label", "")),
                "supplier": _with_special_chars(raw_config, company.get("supplier", "")),
                "supplier_address": [
                    _with_special_chars(raw_config, line) for line in company.get("supplier_address", [])
                ],
                "customer": _with_special_chars(raw_config, company.get("customer", "")),
                "site": _with_special_chars(raw_config, site.get("label", "")),
                "site_address": [_with_special_chars(raw_config, line) for line in site.get("customer_address", [])],
                "country": "" if site_omit.get("country", False) else _with_special_chars(raw_config, site.get("country", "")),
                "equipment": equipment_item["equipment"],
                "fuel": _with_special_chars(raw_config, equipment_item.get("fuel", "")),
                "unit": equipment_item.get("unit", site.get("unit", _tr(raw_config, "litres"))),
                "quantity": _q2(_parse_decimal(equipment_item.get("quantity"), "0")),
                "delivery_note_no": note_no,
                "delivery_date": delivery_date,
                "driver_ref": f"TRK-{rng.randint(1, 24):02d}",
                "customer_signature": _tr(raw_config, "received"),
                "period_label": fp["label"],
                "period_start": fp["start_date"],
                "period_end": fp["end_date"],
            })

    return records


def _build_fuel_card_statements(raw_config: dict) -> list[dict]:
    fp = _financial_period(raw_config)
    seed = int(raw_config.get("random_seed", 42))
    days_in_period = max((fp["end_date"] - fp["start_date"]).days, 0)
    statements: list[dict] = []

    for company_index, company in enumerate(raw_config.get("companies", []), start=1):
        transactions: list[dict] = []
        for transaction_index, site in enumerate(company.get("sites", []), start=1):
            site_omit = site.get("_omit", {})
            site_value = "" if site_omit.get("label", False) else _with_special_chars(raw_config, site.get("label", ""))
            country_value = "" if site_omit.get("country", False) else _with_special_chars(raw_config, site.get("country", ""))
            equipment_items = _site_equipment_items(raw_config, site, include_emission_source=True)

            for equipment_index, equipment_item in enumerate(equipment_items, start=1):
                quantity = _q2(_parse_decimal(equipment_item.get("quantity"), "0"))
                unit_price = _q2(_parse_decimal(equipment_item.get("unit_price"), "0"))
                total = _q2(quantity * unit_price)
                rng = random.Random(f"{seed}:fuel_card:{company_index}:{transaction_index}:{equipment_index}")
                transaction_date = fp["start_date"] + timedelta(days=rng.randint(0, days_in_period))
                reference_value = site_value or equipment_item["equipment"] or _tr(raw_config, "stationary_equipment")

                transactions.append({
                    "card_number": company.get("card_number") or site.get("card_number", ""),
                    "date": transaction_date,
                    "merchant": _with_special_chars(raw_config, company.get("merchant") or site.get("merchant", "")),
                    "site": site_value,
                    "country": country_value,
                    "equipment": equipment_item["equipment"],
                    "emission_source": equipment_item["emission_source"],
                    "reference": reference_value,
                    "fuel": _with_special_chars(raw_config, equipment_item.get("fuel", "")),
                    "quantity": quantity,
                    "unit": equipment_item.get("unit", site.get("unit", default_fuel_volume_unit("fuel_card"))),
                    "unit_price": unit_price,
                    "total": total,
                })

        transactions.sort(key=lambda row: (row["date"], row["merchant"], row["card_number"]))
        statements.append({
            "company": _with_special_chars(raw_config, company.get("label", "")),
            "account_name": _with_special_chars(raw_config, company.get("customer") or company.get("label", "")),
            "provider": _with_special_chars(raw_config, company.get("supplier", "")),
            "currency": company.get("currency", "GBP (£)"),
            "period_label": fp["label"],
            "period_start": fp["start_date"],
            "period_end": fp["end_date"],
            "transactions": transactions,
            "statement_total": _q2(sum((row["total"] for row in transactions), Decimal("0"))),
        })

    return statements


def _date_within_month(year: int, month: int, day: int) -> date:
    return date(year, month, min(day, monthrange(year, month)[1]))


def _fmt_percent(value: float) -> str:
    return f"{round(value):.0f}%"


def _format_time(minutes_total: int) -> str:
    hours = (minutes_total // 60) % 24
    minutes = minutes_total % 60
    return f"{hours:02d}:{minutes:02d}"


def _build_generator_log_rows(raw_config: dict) -> list[dict]:
    fp = _financial_period(raw_config)
    months = _months_in_range(fp["start_date"], fp["end_date"])
    rows: list[dict] = []
    seed = int(raw_config.get("random_seed", 42))

    for company_index, site_index, company, site in _iter_company_sites(raw_config):
        site_omit = site.get("_omit", {})
        equipment_items = _site_equipment_items(raw_config, site, include_emission_source=True)

        for equipment_index, equipment_item in enumerate(equipment_items, start=1):
            runs_per_month = max(int(equipment_item.get("runs_per_month", 3)), 1)
            tank_capacity = float(equipment_item.get("tank_capacity", 800))
            burn_rate = float(equipment_item.get("fuel_used_per_hour", 15))
            min_hours = float(equipment_item.get("run_hours_min", 0.5))
            max_hours = max(float(equipment_item.get("run_hours_max", 4.0)), min_hours)
            quantity_mode = equipment_item.get("quantity_mode", "tank_level_change")
            rng = random.Random(f"{seed}:generator_log:{company_index}:{site_index}:{equipment_index}")
            for year, month in months:
                days_in_month = monthrange(year, month)[1]
                chosen_days = sorted(rng.sample(range(1, days_in_month + 1), k=min(runs_per_month, days_in_month)))
                for day in chosen_days:
                    run_date = _date_within_month(year, month, day)
                    start_minutes = rng.choice([7 * 60, 8 * 60, 9 * 60, 13 * 60, 18 * 60])
                    run_hours = round(rng.uniform(min_hours, max_hours), 2)
                    end_minutes = start_minutes + int(run_hours * 60)

                    if quantity_mode == "explicit_fuel_used":
                        fuel_used = round(run_hours * burn_rate * rng.uniform(0.92, 1.12), 2)
                        start_pct = float(rng.randint(52, 92))
                        end_pct = max(0.0, start_pct - ((fuel_used / max(tank_capacity, 1.0)) * 100.0))
                    else:
                        start_pct = float(rng.randint(55, 95))
                        estimated_delta = max((run_hours * burn_rate / max(tank_capacity, 1.0)) * 100.0, 1.0)
                        delta_pct = min(start_pct, estimated_delta * rng.uniform(0.9, 1.15))
                        end_pct = max(0.0, start_pct - delta_pct)
                        fuel_used = round((start_pct - end_pct) / 100.0 * tank_capacity, 2)

                    rows.append({
                        "company": _with_special_chars(raw_config, company.get("label", "")),
                        "site": _with_special_chars(raw_config, site.get("label", "")),
                        "country": "" if site_omit.get("country", False) else _with_special_chars(raw_config, site.get("country", "")),
                        "equipment": equipment_item["equipment"],
                        "emission_source": equipment_item["emission_source"],
                        "fuel": _with_special_chars(raw_config, equipment_item.get("fuel", "")),
                        "period": run_date.isoformat(),
                        "date": run_date,
                        "start_time": _format_time(start_minutes),
                        "end_time": _format_time(end_minutes),
                        "run_hours": run_hours,
                        "start_fuel": _fmt_percent(start_pct),
                        "end_fuel": _fmt_percent(end_pct),
                        "fuel_used": round(fuel_used, 2),
                        "unit": equipment_item.get("unit", site.get("unit", default_fuel_volume_unit("generator_log"))),
                        "notes": _with_special_chars(
                            raw_config,
                            rng.choice(
                                [
                                    _tr(raw_config, "test_run"),
                                    _tr(raw_config, "power_outage"),
                                    _tr(raw_config, "maintenance_test"),
                                    _tr(raw_config, "load_bank_test"),
                                ]
                            ),
                        ),
                    })
    return sorted(rows, key=lambda row: (row["site"], row["equipment"], row["date"], row["start_time"]))


def _build_bems_site_blocks(raw_config: dict) -> list[dict]:
    fp = _financial_period(raw_config)
    blocks: list[dict] = []

    for _, _, company, site in _iter_company_sites(raw_config):
        site_omit = site.get("_omit", {})
        assets: list[dict] = []
        for asset in site.get("assets", []):
            asset_omit = asset.get("_omit", {})
            assets.append({
                "asset_tag": _with_special_chars(raw_config, asset.get("asset_tag", "")),
                "equipment_name": "" if asset_omit.get("equipment_name", False) else _with_special_chars(raw_config, asset.get("equipment_name", "")),
                "emission_source": "" if asset_omit.get("emission_source", False) else _with_special_chars(raw_config, asset.get("emission_source", "")),
                "fuel": "" if asset_omit.get("fuel", False) else _with_special_chars(raw_config, asset.get("fuel", "")),
                "unit": asset.get("unit", "kWh"),
                "sensor_name": "" if asset_omit.get("sensor_name", False) else _with_special_chars(raw_config, asset.get("sensor_name", "")),
                "quantity": _q2(_parse_decimal(asset.get("quantity"), "0")),
                "operating_hours": None if asset_omit.get("operating_hours", False) else _q2(_parse_decimal(asset.get("operating_hours"), "0")),
                "_omit": asset_omit,
            })
        blocks.append({
            "company": _with_special_chars(raw_config, company.get("label", "")),
            "site": _with_special_chars(raw_config, site.get("label", "")),
            "country": "" if site_omit.get("country", False) else _with_special_chars(raw_config, site.get("country", "")),
            "period_label": fp["label"],
            "period_start": fp["start_date"],
            "period_end": fp["end_date"],
            "assets": assets,
            "_omit": site_omit,
        })
    return blocks


def _bems_asset_weights(asset: dict, timestamps: list[datetime], rng: random.Random) -> list[float]:
    emission_source = asset.get("emission_source", "").lower()
    equipment = asset.get("equipment_name", "").lower()

    if "generator" in emission_source or "generator" in equipment:
        active_count = max(1, min(len(timestamps), max(4, len(timestamps) // 40)))
        active_indices = set(rng.sample(range(len(timestamps)), active_count))
        weights: list[float] = []
        for idx, ts in enumerate(timestamps):
            if idx not in active_indices:
                weights.append(0.0)
                continue
            base = 1.0 if 8 <= ts.hour <= 18 else 0.6
            weights.append(base * rng.uniform(0.9, 1.15))
        return weights

    weights = []
    for ts in timestamps:
        hour = ts.hour + ts.minute / 60
        base = 0.55
        if 5 <= hour < 8:
            base = 0.95
        elif 8 <= hour < 18:
            base = 1.25
        elif 18 <= hour < 22:
            base = 0.88
        if ts.weekday() >= 5:
            base *= 0.82
        weights.append(base * rng.uniform(0.92, 1.08))
    return weights


def _distribute_bems_series(total_quantity: Decimal, weights: list[float]) -> list[Decimal]:
    if not weights:
        return []
    weight_total = sum(weights)
    if weight_total <= 0:
        even_value = _q2(total_quantity / Decimal(str(len(weights))))
        series = [even_value for _ in weights]
        if series:
            series[-1] = _q2(total_quantity - sum(series[:-1]))
        return series

    series = [
        _q2(total_quantity * Decimal(str(weight / weight_total)))
        for weight in weights
    ]
    series[-1] = _q2(total_quantity - sum(series[:-1]))
    return series


def _build_bems_trend_exports(raw_config: dict) -> list[dict]:
    fp = _financial_period(raw_config)
    interval_minutes = _bems_interval_minutes(raw_config)
    timestamps = _timestamp_range(fp["start_date"], fp["end_date"], interval_minutes)
    seed = int(raw_config.get("random_seed", 42))
    exports: list[dict] = []

    for block_index, block in enumerate(_build_bems_site_blocks(raw_config), start=1):
        rows: list[dict] = []
        for asset_index, asset in enumerate(block["assets"], start=1):
            rng = random.Random(f"{seed}:bems:{block_index}:{asset_index}:{asset['asset_tag']}")
            weights = _bems_asset_weights(asset, timestamps, rng)
            values = _distribute_bems_series(asset["quantity"], weights)
            for timestamp, value in zip(timestamps, values):
                rows.append({
                    "timestamp": timestamp,
                    "site": block["site"],
                    "asset_tag": asset["asset_tag"],
                    "sensor_name": asset["sensor_name"],
                    "value": float(value),
                    "unit": asset["unit"],
                })
        exports.append({
            "company": block["company"],
            "site": block["site"],
            "country": block["country"],
            "period_label": block["period_label"],
            "period_start": block["period_start"],
            "period_end": block["period_end"],
            "rows": rows,
            "assets": block["assets"],
        })
    return exports


def _ground_truth_entries(raw_config: dict) -> list[dict]:
    document_type = raw_config.get("document_type", "fuel_invoice")
    if document_type == "fuel_invoice":
        return [
            {
                "Company": record["company"],
                "Site": record["site"],
                "Country": record["country"],
                "Period": f"{record['period_start'].isoformat()} to {record['period_end'].isoformat()}",
                "Equipment": record["equipment"],
                "Emission source": record["emission_source"],
                "Fuel": record["fuel"],
                "Quantity": float(record["quantity"]),
                "Unit": record["unit"],
                "Cost": float(record["subtotal"]),
                "Currency": record["currency"].split()[0],
            }
            for record in _build_fuel_invoice_records(raw_config)
        ]
    if document_type == "delivery_note":
        return [
            {
                "Company": record["company"],
                "Site": record["site"],
                "Country": record["country"],
                "Period": record["delivery_date"].isoformat(),
                "Equipment": record["equipment"],
                "Fuel": record["fuel"],
                "Quantity": float(record["quantity"]),
                "Unit": record["unit"],
            }
            for record in _build_delivery_note_records(raw_config)
        ]
    if document_type == "fuel_card":
        return [
            {
                "Company": statement["company"],
                "Site": transaction["site"],
                "Country": transaction["country"],
                "Period": f"{statement['period_start'].isoformat()} to {statement['period_end'].isoformat()}",
                "Equipment": transaction["equipment"],
                "Emission source": transaction["emission_source"],
                "Fuel": transaction["fuel"],
                "Quantity": float(transaction["quantity"]),
                "Unit": transaction["unit"],
                "Cost": float(transaction["total"]),
                "Currency": statement["currency"].split()[0],
            }
            for statement in _build_fuel_card_statements(raw_config)
            for transaction in statement["transactions"]
        ]
    if document_type == "bems":
        return [
            {
                "Company": block["company"],
                "Site": block["site"],
                "Country": block["country"],
                "Period": f"{block['period_start'].isoformat()} to {block['period_end'].isoformat()}",
                "Equipment": asset["equipment_name"],
                "Emission source": asset["emission_source"],
                "Fuel": asset["fuel"],
                "Quantity": float(asset["quantity"]),
                "Unit": asset["unit"],
            }
            for block in _build_bems_site_blocks(raw_config)
            for asset in block["assets"]
        ]
    return [
        {
            "Site": row["site"],
            "Country": row["country"],
            "Period": row["period"],
            "Equipment": row["equipment"],
            "Emission source": row["emission_source"],
            "Fuel": row["fuel"],
            "Quantity": row["fuel_used"],
            "Unit": row["unit"],
        }
        for row in _build_generator_log_rows(raw_config)
    ]


def generate_ground_truth_json(raw_config: dict) -> bytes:
    import json

    return json.dumps(_ground_truth_entries(raw_config), indent=2).encode("utf-8")


def _draw_multiline(c: canvas.Canvas, x: float, y: float, lines: list[str], leading: int = 12) -> float:
    for line in lines:
        c.drawString(x, y, line)
        y -= leading
    return y


def _shade_docx_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def _style_docx_document(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(9.5)


def _set_docx_cell_text(cell, text: str, *, bold: bool = False, color: str | None = None, size: float | None = None) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(text)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if size is not None:
        run.font.size = Pt(size)


def generate_fuel_invoice_pdf(raw_config: dict) -> bytes:
    records = _build_fuel_invoice_records(raw_config)
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setTitle(raw_config.get("document", {}).get("title", _tr(raw_config, "fuel_invoice_title")))
    c.setSubject(raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion"))

    for index, record in enumerate(records):
        if index > 0:
            c.showPage()

        accent = colors.HexColor("#1E5B88")
        c.setFillColor(accent)
        c.rect(36, PAGE_H - 72, PAGE_W - 72, 28, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(48, PAGE_H - 62, record["supplier"])
        c.setFont("Helvetica", 8)
        c.drawRightString(PAGE_W - 48, PAGE_H - 61, _tr(raw_config, "fuel_invoice_title"))

        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        y = PAGE_H - 108
        y = _draw_multiline(c, 48, y, record["supplier_address"])

        c.setFont("Helvetica-Bold", 11)
        c.drawString(48, y - 10, _tr(raw_config, "bill_to"))
        c.setFont("Helvetica", 10)
        bill_to_lines = [record["customer"], *record["site_address"]]
        _draw_multiline(c, 48, y - 26, bill_to_lines)

        c.setFont("Helvetica-Bold", 11)
        c.drawString(320, PAGE_H - 108, _tr(raw_config, "invoice_details"))
        c.setFont("Helvetica", 10)
        meta_lines = [
            f"{_tr(raw_config, 'invoice_no')}: {record['invoice_no']}",
            f"{_tr(raw_config, 'invoice_date')}: {_fmt_date(record['invoice_date'])}",
            f"{_tr(raw_config, 'billing_period')}: {_fmt_date(record['period_start'])} - {_fmt_date(record['period_end'])}",
            f"{_tr(raw_config, 'currency')}: {record['currency']}",
            f"{_tr(raw_config, 'country')}: {record['country']}",
        ]
        _draw_multiline(c, 320, PAGE_H - 126, meta_lines)

        c.setFont("Helvetica-Bold", 11)
        c.drawString(320, PAGE_H - 236, _tr(raw_config, "delivery_site"))
        c.setFont("Helvetica", 10)
        delivery_lines = [record["site"], record["equipment"], record["emission_source"]]
        _draw_multiline(c, 320, PAGE_H - 254, [line for line in delivery_lines if line])

        table_top = PAGE_H - 330
        table_x = 48
        table_widths = [210, 68, 58, 84, 84]
        headers = [
            _tr(raw_config, "product"),
            _tr(raw_config, "quantity"),
            _tr(raw_config, "unit"),
            _tr(raw_config, "unit_price"),
            _tr(raw_config, "amount"),
        ]
        c.setFillColor(accent)
        c.rect(table_x, table_top, sum(table_widths), 22, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 9)

        x_cursor = table_x + 6
        for header, width in zip(headers, table_widths):
            c.drawString(x_cursor, table_top + 7, header)
            x_cursor += width

        rows = [
            (
                record["fuel"],
                f"{record['quantity']:,.2f}",
                record["unit"],
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['unit_price'])}",
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['fuel_cost'])}",
            ),
            (
                _tr(raw_config, "delivery_charge"),
                "1",
                _tr(raw_config, "each"),
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['delivery_charge'])}",
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['delivery_charge'])}",
            ),
        ]

        y_row = table_top - 24
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.black)
        for row in rows:
            c.rect(table_x, y_row, sum(table_widths), 20, fill=0, stroke=1)
            x_cursor = table_x + 6
            for value, width in zip(row, table_widths):
                c.drawString(x_cursor, y_row + 6, str(value))
                x_cursor += width
            y_row -= 20

        summary_y = y_row - 20
        summary_lines = [
            (_tr(raw_config, "subtotal"), record["subtotal"], False),
            (f"VAT ({record['vat_rate']}%)", record["vat"], False),
            (_tr(raw_config, "total"), record["total"], True),
        ]
        for label, value, is_total in summary_lines:
            c.setFont("Helvetica-Bold" if is_total else "Helvetica", 10)
            c.drawRightString(PAGE_W - 180, summary_y, label)
            c.drawRightString(
                PAGE_W - 48,
                summary_y,
                f"{_currency_symbol(record['currency'])}{_fmt_money(value)}",
            )
            summary_y -= 18

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.grey)
        c.drawString(
            48,
            42,
            _tr(raw_config, "fuel_invoice_footer"),
        )

    c.save()
    return buffer.getvalue()


def generate_fuel_invoice_docx(raw_config: dict) -> bytes:
    records = _build_fuel_invoice_records(raw_config)
    document = Document()
    document.core_properties.title = raw_config.get("document", {}).get("title", _tr(raw_config, "fuel_invoice_title"))
    document.core_properties.subject = raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion")

    for index, record in enumerate(records):
        document.add_heading(record["supplier"], level=1)
        document.add_paragraph(_tr(raw_config, "fuel_invoice_title"))

        top_table = document.add_table(rows=2, cols=2)
        top_table.style = "Table Grid"
        top_table.cell(0, 0).text = _tr(raw_config, "invoice_details")
        top_table.cell(0, 1).text = _tr(raw_config, "delivery_site")
        top_table.cell(1, 0).text = (
            f"{_tr(raw_config, 'invoice_no')}: {record['invoice_no']}\n"
            f"{_tr(raw_config, 'invoice_date')}: {_fmt_date(record['invoice_date'])}\n"
            f"{_tr(raw_config, 'billing_period')}: {_fmt_date(record['period_start'])} - {_fmt_date(record['period_end'])}\n"
            f"{_tr(raw_config, 'currency')}: {record['currency']}\n"
            f"{_tr(raw_config, 'country')}: {record['country']}"
        )
        top_table.cell(1, 1).text = "\n".join(
            [line for line in [record["site"], record["equipment"], record["emission_source"]] if line]
        )

        bill_to_heading = document.add_paragraph()
        bill_to_heading.add_run(_tr(raw_config, "bill_to")).bold = True
        document.add_paragraph("\n".join([record["customer"], *record["site_address"]]))

        line_table = document.add_table(rows=1, cols=5)
        line_table.style = "Table Grid"
        for cell, header in zip(
            line_table.rows[0].cells,
            [
                _tr(raw_config, "product"),
                _tr(raw_config, "quantity"),
                _tr(raw_config, "unit"),
                _tr(raw_config, "unit_price"),
                _tr(raw_config, "amount"),
            ],
        ):
            cell.text = header

        line_rows = [
            (
                record["fuel"],
                f"{record['quantity']:,.2f}",
                record["unit"],
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['unit_price'])}",
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['fuel_cost'])}",
            ),
            (
                _tr(raw_config, "delivery_charge"),
                "1",
                _tr(raw_config, "each"),
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['delivery_charge'])}",
                f"{_currency_symbol(record['currency'])}{_fmt_money(record['delivery_charge'])}",
            ),
        ]
        for values in line_rows:
            row = line_table.add_row().cells
            for cell, value in zip(row, values):
                cell.text = str(value)

        totals = document.add_table(rows=3, cols=2)
        totals.style = "Table Grid"
        totals.cell(0, 0).text = _tr(raw_config, "subtotal")
        totals.cell(0, 1).text = f"{_currency_symbol(record['currency'])}{_fmt_money(record['subtotal'])}"
        totals.cell(1, 0).text = f"VAT ({record['vat_rate']}%)"
        totals.cell(1, 1).text = f"{_currency_symbol(record['currency'])}{_fmt_money(record['vat'])}"
        totals.cell(2, 0).text = _tr(raw_config, "total")
        totals.cell(2, 1).text = f"{_currency_symbol(record['currency'])}{_fmt_money(record['total'])}"

        if index < len(records) - 1:
            document.add_page_break()

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def generate_delivery_note_pdf(raw_config: dict) -> bytes:
    records = _build_delivery_note_records(raw_config)
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setTitle(raw_config.get("document", {}).get("title", _tr(raw_config, "delivery_note_title")))
    c.setSubject(raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion fuel delivery"))

    for index, record in enumerate(records):
        if index > 0:
            c.showPage()

        accent = colors.HexColor("#1E5B88")
        c.setFillColor(accent)
        c.rect(36, PAGE_H - 72, PAGE_W - 72, 30, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 15)
        c.drawString(48, PAGE_H - 62, _tr(raw_config, "delivery_note_title"))

        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        y = PAGE_H - 114
        meta_lines = [
            f"{_tr(raw_config, 'supplier')}: {record['supplier']}",
            f"{_tr(raw_config, 'delivery_note_no')}: {record['delivery_note_no']}",
            f"{_tr(raw_config, 'delivery_date')}: {_fmt_date(record['delivery_date'])}",
        ]
        y = _draw_multiline(c, 48, y, meta_lines, leading=16)

        c.setFont("Helvetica-Bold", 11)
        c.drawString(48, y - 8, _tr(raw_config, "customer"))
        c.setFont("Helvetica", 10)
        c.drawString(48, y - 24, record["customer"])

        c.setFont("Helvetica-Bold", 11)
        c.drawString(48, y - 62, _tr(raw_config, "delivery_address"))
        c.setFont("Helvetica", 10)
        _draw_multiline(c, 48, y - 78, [record["site"], *record["site_address"]], leading=14)

        panel_top = PAGE_H - 336
        panel_height = 166
        c.setFillColor(colors.HexColor("#F5F8FB"))
        c.roundRect(36, panel_top - panel_height, PAGE_W - 72, panel_height, 8, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 11)
        detail_lines = [
            (_tr(raw_config, "product_delivered"), record["fuel"]),
            (_tr(raw_config, "tank_equipment"), record["equipment"]),
            (_tr(raw_config, "delivered_quantity"), f"{_fmt_money(record['quantity'])} {record['unit']}"),
            (_tr(raw_config, "driver_ref"), record["driver_ref"]),
            (_tr(raw_config, "customer_signature"), record["customer_signature"]),
        ]
        line_y = panel_top - 28
        for label, value in detail_lines:
            if not value:
                continue
            c.drawString(52, line_y, f"{label}:")
            c.setFont("Helvetica", 10)
            c.drawString(184, line_y, value)
            c.setFont("Helvetica-Bold", 11)
            line_y -= 28

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.grey)
        c.drawString(48, 42, _tr(raw_config, "delivery_note_footer"))

    c.save()
    return buffer.getvalue()


def generate_delivery_note_docx(raw_config: dict) -> bytes:
    records = _build_delivery_note_records(raw_config)
    document = Document()
    _style_docx_document(document)
    document.core_properties.title = raw_config.get("document", {}).get("title", _tr(raw_config, "delivery_note_title"))
    document.core_properties.subject = raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion fuel delivery")

    for index, record in enumerate(records):
        banner = document.add_table(rows=1, cols=2)
        banner.style = "Table Grid"
        banner.autofit = False
        banner.columns[0].width = Inches(4.8)
        banner.columns[1].width = Inches(2.0)
        _shade_docx_cell(banner.cell(0, 0), "1E5B88")
        _shade_docx_cell(banner.cell(0, 1), "1E5B88")
        _set_docx_cell_text(banner.cell(0, 0), record["supplier"], bold=True, color="FFFFFF", size=13)
        _set_docx_cell_text(banner.cell(0, 1), _tr(raw_config, "delivery_note_title"), bold=True, color="FFFFFF", size=12)
        banner.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        meta = document.add_table(rows=2, cols=2)
        meta.style = "Table Grid"
        meta.autofit = False
        meta.columns[0].width = Inches(4.2)
        meta.columns[1].width = Inches(2.6)
        for cell, heading in zip(meta.rows[0].cells, [_tr(raw_config, "delivery_address"), _tr(raw_config, "delivery_details")]):
            _shade_docx_cell(cell, "DCEBF5")
            _set_docx_cell_text(cell, heading, bold=True)
        _set_docx_cell_text(
            meta.cell(1, 0),
            "\n".join([record["customer"], "", record["site"], *record["site_address"]]),
        )
        _set_docx_cell_text(
            meta.cell(1, 1),
            (
                f"{_tr(raw_config, 'delivery_note_no')}: {record['delivery_note_no']}\n"
                f"{_tr(raw_config, 'delivery_date')}: {_fmt_date(record['delivery_date'])}\n"
                f"{_tr(raw_config, 'country')}: {record['country']}"
            ),
        )

        spacer = document.add_paragraph()
        spacer.paragraph_format.space_after = Pt(2)

        section_heading = document.add_paragraph()
        section_run = section_heading.add_run(_tr(raw_config, "delivery_confirmation"))
        section_run.bold = True
        section_run.font.size = Pt(11)
        section_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT

        details = document.add_table(rows=0, cols=2)
        details.style = "Table Grid"
        details.autofit = False
        details.columns[0].width = Inches(2.1)
        details.columns[1].width = Inches(4.7)
        for label, value in [
            (_tr(raw_config, "product_delivered"), record["fuel"]),
            (_tr(raw_config, "tank_equipment"), record["equipment"]),
            (_tr(raw_config, "delivered_quantity"), f"{_fmt_money(record['quantity'])} {record['unit']}"),
            (_tr(raw_config, "driver_ref"), record["driver_ref"]),
            (_tr(raw_config, "customer_signature"), record["customer_signature"]),
        ]:
            if not value:
                continue
            row = details.add_row().cells
            _shade_docx_cell(row[0], "F5F8FB")
            _set_docx_cell_text(row[0], label, bold=True)
            _set_docx_cell_text(row[1], value)

        footer = document.add_paragraph()
        footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        footer_run = footer.add_run(_tr(raw_config, "delivery_note_footer"))
        footer_run.font.size = Pt(8)
        footer_run.font.color.rgb = RGBColor.from_string("6E7A86")

        if index < len(records) - 1:
            document.add_page_break()

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def generate_fuel_card_pdf(raw_config: dict) -> bytes:
    statements = _build_fuel_card_statements(raw_config)
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setTitle(raw_config.get("document", {}).get("title", _tr(raw_config, "fuel_card_title")))
    c.setSubject(raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion fuel card transactions"))

    page_size = 18
    for statement_index, statement in enumerate(statements):
        transactions = statement["transactions"]
        for page_start in range(0, max(len(transactions), 1), page_size):
            if statement_index > 0 or page_start > 0:
                c.showPage()

            accent = colors.HexColor("#1E5B88")
            c.setFillColor(accent)
            c.rect(36, PAGE_H - 74, PAGE_W - 72, 30, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 15)
            c.drawString(48, PAGE_H - 63, _tr(raw_config, "fuel_card_title"))

            c.setFillColor(colors.black)
            c.setFont("Helvetica", 10)
            meta_lines = [
                f"{_tr(raw_config, 'account_name')}: {statement['account_name']}",
                f"{_tr(raw_config, 'provider')}: {statement['provider']}",
                f"{_tr(raw_config, 'statement_period')}: {_fmt_date(statement['period_start'])} - {_fmt_date(statement['period_end'])}",
                f"{_tr(raw_config, 'currency')}: {statement['currency']}",
            ]
            _draw_multiline(c, 48, PAGE_H - 108, meta_lines, leading=14)

            table_x = 36
            table_top = PAGE_H - 188
            column_widths = [62, 54, 116, 102, 66, 40, 28, 54, 54]
            headers = [
                _tr(raw_config, "card_no"),
                _tr(raw_config, "date"),
                _tr(raw_config, "merchant"),
                _tr(raw_config, "reference"),
                _tr(raw_config, "product"),
                _tr(raw_config, "qty"),
                _tr(raw_config, "unit"),
                _tr(raw_config, "unit_price"),
                _tr(raw_config, "total"),
            ]
            c.setFillColor(accent)
            c.rect(table_x, table_top, sum(column_widths), 22, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 7.5)
            cursor = table_x + 4
            for header, width in zip(headers, column_widths):
                c.drawString(cursor, table_top + 7, header)
                cursor += width

            row_y = table_top - 18
            c.setFillColor(colors.black)
            c.setFont("Helvetica", 7.5)
            currency_symbol = _currency_symbol(statement["currency"])
            for transaction in transactions[page_start:page_start + page_size]:
                c.rect(table_x, row_y, sum(column_widths), 18, fill=0, stroke=1)
                cursor = table_x + 4
                row_values = [
                    transaction["card_number"],
                    transaction["date"].strftime("%d-%m-%y"),
                    transaction["merchant"],
                    transaction["reference"],
                    transaction["fuel"],
                    _fmt_money(transaction["quantity"]),
                    transaction["unit"],
                    f"{currency_symbol}{_fmt_money(transaction['unit_price'])}",
                    f"{currency_symbol}{_fmt_money(transaction['total'])}",
                ]
                for value, width in zip(row_values, column_widths):
                    c.drawString(cursor, row_y + 5, str(value))
                    cursor += width
                row_y -= 18

            if page_start + page_size >= len(transactions):
                c.setFont("Helvetica-Bold", 10)
                c.drawRightString(PAGE_W - 180, row_y - 18, _tr(raw_config, "statement_total"))
                c.drawRightString(PAGE_W - 48, row_y - 18, f"{currency_symbol}{_fmt_money(statement['statement_total'])}")

            c.setFont("Helvetica", 8)
            c.setFillColor(colors.grey)
            c.drawString(48, 42, _tr(raw_config, "fuel_card_footer"))

    c.save()
    return buffer.getvalue()


def generate_fuel_card_docx(raw_config: dict) -> bytes:
    statements = _build_fuel_card_statements(raw_config)
    document = Document()
    _style_docx_document(document)
    document.core_properties.title = raw_config.get("document", {}).get("title", _tr(raw_config, "fuel_card_title"))
    document.core_properties.subject = raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion fuel card transactions")

    for statement_index, statement in enumerate(statements):
        banner = document.add_table(rows=1, cols=2)
        banner.style = "Table Grid"
        banner.autofit = False
        banner.columns[0].width = Inches(4.8)
        banner.columns[1].width = Inches(2.0)
        _shade_docx_cell(banner.cell(0, 0), "1E5B88")
        _shade_docx_cell(banner.cell(0, 1), "1E5B88")
        _set_docx_cell_text(banner.cell(0, 0), statement["account_name"], bold=True, color="FFFFFF", size=13)
        _set_docx_cell_text(banner.cell(0, 1), _tr(raw_config, "fuel_card_title"), bold=True, color="FFFFFF", size=12)
        banner.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        meta = document.add_table(rows=2, cols=2)
        meta.style = "Table Grid"
        meta.autofit = False
        meta.columns[0].width = Inches(4.2)
        meta.columns[1].width = Inches(2.6)
        for cell, heading in zip(meta.rows[0].cells, [_tr(raw_config, "account_details"), _tr(raw_config, "statement_details")]):
            _shade_docx_cell(cell, "DCEBF5")
            _set_docx_cell_text(cell, heading, bold=True)
        _set_docx_cell_text(
            meta.cell(1, 0),
            f"{_tr(raw_config, 'account_name')}: {statement['account_name']}\n{_tr(raw_config, 'provider')}: {statement['provider']}",
        )
        _set_docx_cell_text(
            meta.cell(1, 1),
            (
                f"{_tr(raw_config, 'statement_period')}: {_fmt_date(statement['period_start'])} - {_fmt_date(statement['period_end'])}\n"
                f"{_tr(raw_config, 'currency')}: {statement['currency']}"
            ),
        )

        document.add_paragraph()
        table = document.add_table(rows=1, cols=9)
        table.style = "Table Grid"
        headers = [
            _tr(raw_config, "card_no"),
            _tr(raw_config, "date"),
            _tr(raw_config, "merchant"),
            _tr(raw_config, "reference"),
            _tr(raw_config, "product"),
            _tr(raw_config, "qty"),
            _tr(raw_config, "unit"),
            _tr(raw_config, "unit_price"),
            _tr(raw_config, "total"),
        ]
        for cell, header in zip(table.rows[0].cells, headers):
            _shade_docx_cell(cell, "F5F8FB")
            _set_docx_cell_text(cell, header, bold=True)

        currency_symbol = _currency_symbol(statement["currency"])
        for transaction in statement["transactions"]:
            row = table.add_row().cells
            values = [
                transaction["card_number"],
                transaction["date"].strftime("%d-%m-%y"),
                transaction["merchant"],
                transaction["reference"],
                transaction["fuel"],
                _fmt_money(transaction["quantity"]),
                transaction["unit"],
                f"{currency_symbol}{_fmt_money(transaction['unit_price'])}",
                f"{currency_symbol}{_fmt_money(transaction['total'])}",
            ]
            for cell, value in zip(row, values):
                _set_docx_cell_text(cell, str(value))

        totals = document.add_table(rows=1, cols=2)
        totals.style = "Table Grid"
        _shade_docx_cell(totals.cell(0, 0), "F5F8FB")
        _set_docx_cell_text(totals.cell(0, 0), _tr(raw_config, "statement_total"), bold=True)
        _set_docx_cell_text(totals.cell(0, 1), f"{currency_symbol}{_fmt_money(statement['statement_total'])}", bold=True)

        if statement_index < len(statements) - 1:
            document.add_page_break()

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def generate_fuel_card_xlsx(raw_config: dict) -> bytes:
    statements = _build_fuel_card_statements(raw_config)
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    for index, statement in enumerate(statements, start=1):
        sheet = workbook.create_sheet(title=(statement["company"] or f"Account {index}")[:31])
        sheet["A1"] = _tr(raw_config, "fuel_card_title")
        sheet["A1"].font = Font(size=14, bold=True)
        sheet["A2"] = _tr(raw_config, "account_name")
        sheet["B2"] = statement["account_name"]
        sheet["A3"] = _tr(raw_config, "statement_period")
        sheet["B3"] = f"{statement['period_start'].isoformat()} to {statement['period_end'].isoformat()}"
        sheet["A4"] = _tr(raw_config, "currency")
        sheet["B4"] = statement["currency"]

        headers = [
            _tr(raw_config, "card_no"),
            _tr(raw_config, "date"),
            _tr(raw_config, "merchant"),
            _tr(raw_config, "site"),
            _tr(raw_config, "country"),
            _tr(raw_config, "equipment"),
            _tr(raw_config, "emission_source"),
            _tr(raw_config, "product"),
            _tr(raw_config, "qty"),
            _tr(raw_config, "unit"),
            _tr(raw_config, "unit_price"),
            _tr(raw_config, "total"),
        ]
        header_fill = PatternFill(fill_type="solid", fgColor="1E5B88")
        for column_index, header in enumerate(headers, start=1):
            cell = sheet.cell(row=6, column=column_index, value=header)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        row_index = 7
        for transaction in statement["transactions"]:
            values = [
                transaction["card_number"],
                transaction["date"].strftime("%d-%m-%y"),
                transaction["merchant"],
                transaction["site"],
                transaction["country"],
                transaction["equipment"],
                transaction["emission_source"],
                transaction["fuel"],
                float(transaction["quantity"]),
                transaction["unit"],
                float(transaction["unit_price"]),
                float(transaction["total"]),
            ]
            for column_index, value in enumerate(values, start=1):
                sheet.cell(row=row_index, column=column_index, value=value)
            row_index += 1

        widths = [12, 12, 22, 18, 14, 18, 18, 14, 10, 8, 12, 12]
        for column_index, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(column_index)].width = width

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def generate_fuel_card_csv(raw_config: dict) -> bytes:
    statements = _build_fuel_card_statements(raw_config)
    buffer = StringIO()
    writer = csv.writer(buffer)

    for statement_index, statement in enumerate(statements):
        if statement_index > 0:
            writer.writerow([])
        writer.writerow([_tr(raw_config, "fuel_card_title")])
        writer.writerow([_tr(raw_config, "account_name"), statement["account_name"]])
        writer.writerow([_tr(raw_config, "statement_period"), f"{statement['period_start'].isoformat()} to {statement['period_end'].isoformat()}"])
        writer.writerow([_tr(raw_config, "currency"), statement["currency"]])
        writer.writerow([])
        writer.writerow([
            _tr(raw_config, "card_no"),
            _tr(raw_config, "date"),
            _tr(raw_config, "merchant"),
            _tr(raw_config, "site"),
            _tr(raw_config, "country"),
            _tr(raw_config, "equipment"),
            _tr(raw_config, "emission_source"),
            _tr(raw_config, "product"),
            _tr(raw_config, "qty"),
            _tr(raw_config, "unit"),
            _tr(raw_config, "unit_price"),
            _tr(raw_config, "total"),
        ])
        for transaction in statement["transactions"]:
            writer.writerow([
                transaction["card_number"],
                transaction["date"].strftime("%d-%m-%y"),
                transaction["merchant"],
                transaction["site"],
                transaction["country"],
                transaction["equipment"],
                transaction["emission_source"],
                transaction["fuel"],
                f"{float(transaction['quantity']):.2f}",
                transaction["unit"],
                f"{float(transaction['unit_price']):.2f}",
                f"{float(transaction['total']):.2f}",
            ])

    return buffer.getvalue().encode("utf-8-sig")


def _log_headers(raw_config: dict) -> list[str]:
    return [
        _tr(raw_config, "company"),
        _tr(raw_config, "site"),
        _tr(raw_config, "country"),
        _tr(raw_config, "date"),
        _tr(raw_config, "start_time"),
        _tr(raw_config, "end_time"),
        _tr(raw_config, "run_hours"),
        _tr(raw_config, "start_fuel"),
        _tr(raw_config, "end_fuel"),
        _tr(raw_config, "fuel_used"),
        _tr(raw_config, "unit"),
        _tr(raw_config, "equipment"),
        _tr(raw_config, "emission_source"),
        _tr(raw_config, "fuel_type"),
        _tr(raw_config, "notes"),
    ]


def generate_generator_log_xlsx(raw_config: dict) -> bytes:
    rows = _build_generator_log_rows(raw_config)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = _tr(raw_config, "generator_log_sheet_title")

    title = raw_config.get("document", {}).get("title", _tr(raw_config, "generator_operation_log_title"))
    sheet["A1"] = title
    sheet["A1"].font = Font(size=14, bold=True)
    sheet["A2"] = raw_config.get("financial_period", {}).get("label", "")
    sheet["A2"].font = Font(italic=True)

    headers = _log_headers(raw_config)
    header_fill = PatternFill(fill_type="solid", fgColor="1E5B88")
    for column_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=4, column=column_index, value=header)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row_index, row in enumerate(rows, start=5):
        values = [
            row["company"],
            row["site"],
            row["country"],
            row["date"].strftime("%d-%m-%y"),
            row["start_time"],
            row["end_time"],
            row["run_hours"],
            row["start_fuel"],
            row["end_fuel"],
            row["fuel_used"],
            row["unit"],
            row["equipment"],
            row["emission_source"],
            row["fuel"],
            row["notes"],
        ]
        for column_index, value in enumerate(values, start=1):
            sheet.cell(row=row_index, column=column_index, value=value)

    widths = [18, 22, 18, 12, 12, 12, 10, 12, 12, 10, 8, 20, 22, 18, 18]
    for column_index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def generate_generator_log_csv(raw_config: dict) -> bytes:
    rows = _build_generator_log_rows(raw_config)
    buffer = StringIO()
    writer = csv.writer(buffer)
    writer.writerow(_log_headers(raw_config))
    for row in rows:
        writer.writerow([
            row["company"],
            row["site"],
            row["country"],
            row["date"].strftime("%d-%m-%y"),
            row["start_time"],
            row["end_time"],
            f"{row['run_hours']:.2f}",
            row["start_fuel"],
            row["end_fuel"],
            f"{row['fuel_used']:.2f}",
            row["unit"],
            row["equipment"],
            row["emission_source"],
            row["fuel"],
            row["notes"],
        ])
    return buffer.getvalue().encode("utf-8-sig")


def generate_bems_equipment_report_pdf(raw_config: dict) -> bytes:
    blocks = _build_bems_site_blocks(raw_config)
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setTitle(raw_config.get("document", {}).get("title", _tr(raw_config, "bems_equipment_title")))
    c.setSubject(raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion BEMS export"))

    for block_index, block in enumerate(blocks):
        if block_index > 0:
            c.showPage()

        accent = colors.HexColor("#1E5B88")
        c.setFillColor(accent)
        c.rect(36, PAGE_H - 78, PAGE_W - 72, 34, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 15)
        c.drawString(48, PAGE_H - 64, _tr(raw_config, "bems_equipment_title"))

        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        meta_lines = [
            f"{_tr(raw_config, 'company')}: {block['company']}",
            f"{_tr(raw_config, 'site')}: {block['site']}",
            f"{_tr(raw_config, 'country')}: {block['country']}",
            f"{_tr(raw_config, 'reporting_period')}: {block['period_label']}",
        ]
        meta_lines = [line for line in meta_lines if not line.endswith(": ")]
        _draw_multiline(c, 48, PAGE_H - 108, meta_lines, leading=14)

        total_assets = len(block["assets"])
        total_hours = sum(asset["operating_hours"] or Decimal("0") for asset in block["assets"])
        dominant_asset = max(block["assets"], key=lambda asset: asset["quantity"], default=None)
        cards = [
            (_tr(raw_config, "assets"), str(total_assets)),
            (_tr(raw_config, "operating_hours"), _fmt_optional_number(total_hours, " h") or "n/a"),
            (_tr(raw_config, "top_asset"), dominant_asset["asset_tag"] if dominant_asset else "n/a"),
        ]

        x = 48
        card_y = PAGE_H - 205
        for title, value in cards:
            c.setFillColor(colors.HexColor("#F2F6FA"))
            c.roundRect(x, card_y, 150, 46, 6, stroke=0, fill=1)
            c.setFillColor(colors.HexColor("#567389"))
            c.setFont("Helvetica", 8)
            c.drawString(x + 10, card_y + 30, title)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 12)
            c.drawString(x + 10, card_y + 13, value)
            x += 166

        table_x = 48
        table_top = PAGE_H - 290
        column_widths = [64, 136, 88, 74, 64, 48, 74]
        headers = [
            _tr(raw_config, "equipment_tag"),
            _tr(raw_config, "equipment_name"),
            _tr(raw_config, "emission_source"),
            _tr(raw_config, "fuel_type"),
            _tr(raw_config, "consumption"),
            _tr(raw_config, "unit"),
            _tr(raw_config, "operating_hours"),
        ]

        c.setFillColor(accent)
        c.rect(table_x, table_top, sum(column_widths), 24, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8)
        cursor = table_x + 6
        for header, width in zip(headers, column_widths):
            c.drawString(cursor, table_top + 8, header)
            cursor += width

        row_y = table_top - 22
        max_quantity = max((float(asset["quantity"]) for asset in block["assets"]), default=1.0)
        for asset in block["assets"]:
            c.setFillColor(colors.black)
            c.rect(table_x, row_y, sum(column_widths), 20, fill=0, stroke=1)
            cursor = table_x + 6
            row_values = [
                asset["asset_tag"],
                asset["equipment_name"],
                asset["emission_source"],
                asset["fuel"],
                _fmt_money(asset["quantity"]),
                asset["unit"],
                _fmt_optional_number(asset["operating_hours"]),
            ]
            c.setFont("Helvetica", 8)
            for value, width in zip(row_values, column_widths):
                c.drawString(cursor, row_y + 6, str(value))
                cursor += width
            row_y -= 20

        chart_y = row_y - 110
        c.setFont("Helvetica-Bold", 10)
        c.drawString(48, chart_y + 96, _tr(raw_config, "equipment_trend_snapshot"))
        for idx, asset in enumerate(block["assets"][:5]):
            bar_y = chart_y + 72 - (idx * 18)
            bar_width = 220 * (float(asset["quantity"]) / max_quantity if max_quantity else 0)
            c.setFont("Helvetica", 8)
            c.drawString(48, bar_y + 4, asset["asset_tag"])
            c.setFillColor(colors.HexColor("#DCEBF5"))
            c.rect(110, bar_y, 220, 10, fill=1, stroke=0)
            c.setFillColor(accent)
            c.rect(110, bar_y, bar_width, 10, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.drawString(340, bar_y + 2, f"{_fmt_money(asset['quantity'])} {asset['unit']}")

        c.setFont("Helvetica", 8)
        c.setFillColor(colors.grey)
        c.drawString(48, 42, _tr(raw_config, "dashboard_summary_footer"))

    c.save()
    return buffer.getvalue()


def generate_bems_time_series_pdf(raw_config: dict) -> bytes:
    blocks = _build_bems_trend_exports(raw_config)
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.setTitle(raw_config.get("document", {}).get("title", _tr(raw_config, "bems_time_series_title")))
    c.setSubject(raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion BEMS export"))

    for block_index, block in enumerate(blocks):
        rows = block["rows"]
        page_size = 28
        for page_start in range(0, max(len(rows), 1), page_size):
            if block_index > 0 or page_start > 0:
                c.showPage()

            accent = colors.HexColor("#1E5B88")
            c.setFillColor(accent)
            c.rect(36, PAGE_H - 78, PAGE_W - 72, 34, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 15)
            c.drawString(48, PAGE_H - 64, _tr(raw_config, "bems_time_series_title"))

            c.setFillColor(colors.black)
            c.setFont("Helvetica", 10)
            meta_lines = [
                f"{_tr(raw_config, 'company')}: {block['company']}",
                f"{_tr(raw_config, 'site')}: {block['site']}",
                f"{_tr(raw_config, 'country')}: {block['country']}",
                f"{_tr(raw_config, 'reporting_period')}: {block['period_label']}",
            ]
            meta_lines = [line for line in meta_lines if not line.endswith(": ")]
            _draw_multiline(c, 48, PAGE_H - 108, meta_lines, leading=14)

            asset_count = len(block["assets"])
            interval_minutes = _bems_interval_minutes(raw_config)
            cards = [
                (_tr(raw_config, "assets"), str(asset_count)),
                (_tr(raw_config, "interval"), f"{interval_minutes} min"),
                (_tr(raw_config, "rows"), str(len(rows))),
            ]
            x = 48
            card_y = PAGE_H - 205
            for title, value in cards:
                c.setFillColor(colors.HexColor("#F2F6FA"))
                c.roundRect(x, card_y, 150, 46, 6, stroke=0, fill=1)
                c.setFillColor(colors.HexColor("#567389"))
                c.setFont("Helvetica", 8)
                c.drawString(x + 10, card_y + 30, title)
                c.setFillColor(colors.black)
                c.setFont("Helvetica-Bold", 12)
                c.drawString(x + 10, card_y + 13, value)
                x += 166

            table_x = 48
            table_top = PAGE_H - 292
            column_widths = [98, 66, 88, 180, 56, 40]
            headers = [
                _tr(raw_config, "timestamp"),
                _tr(raw_config, "site"),
                _tr(raw_config, "equipment_tag"),
                _tr(raw_config, "sensor_name"),
                _tr(raw_config, "value"),
                _tr(raw_config, "unit"),
            ]

            c.setFillColor(accent)
            c.rect(table_x, table_top, sum(column_widths), 24, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 8)
            cursor = table_x + 4
            for header, width in zip(headers, column_widths):
                c.drawString(cursor, table_top + 8, header)
                cursor += width

            row_y = table_top - 20
            c.setFillColor(colors.black)
            c.setFont("Helvetica", 7)
            for row in rows[page_start:page_start + page_size]:
                c.rect(table_x, row_y, sum(column_widths), 18, fill=0, stroke=1)
                cursor = table_x + 4
                row_values = [
                    row["timestamp"].strftime("%Y-%m-%d %H:%M"),
                    row["site"],
                    row["asset_tag"],
                    row["sensor_name"],
                    f"{row['value']:.2f}",
                    row["unit"],
                ]
                for value, width in zip(row_values, column_widths):
                    c.drawString(cursor, row_y + 5, str(value))
                    cursor += width
                row_y -= 18

            c.setFont("Helvetica", 8)
            c.setFillColor(colors.grey)
            c.drawString(48, 42, _tr(raw_config, "time_series_footer"))

    c.save()
    return buffer.getvalue()


def generate_bems_equipment_report_docx(raw_config: dict) -> bytes:
    blocks = _build_bems_site_blocks(raw_config)
    document = Document()
    core_props = document.core_properties
    core_props.title = raw_config.get("document", {}).get("title", _tr(raw_config, "bems_equipment_title"))
    core_props.subject = raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion BEMS export")

    for block_index, block in enumerate(blocks):
        document.add_heading(_tr(raw_config, "bems_equipment_title"), level=0)

        for line in [
            f"{_tr(raw_config, 'company')}: {block['company']}",
            f"{_tr(raw_config, 'site')}: {block['site']}",
            f"{_tr(raw_config, 'country')}: {block['country']}",
            f"{_tr(raw_config, 'reporting_period')}: {block['period_label']}",
        ]:
            if not line.endswith(": "):
                document.add_paragraph(line)

        total_assets = len(block["assets"])
        total_hours = sum(asset["operating_hours"] or Decimal("0") for asset in block["assets"])
        dominant_asset = max(block["assets"], key=lambda asset: asset["quantity"], default=None)
        summary_table = document.add_table(rows=2, cols=3)
        summary_table.style = "Table Grid"
        summary_headers = [_tr(raw_config, "assets"), _tr(raw_config, "operating_hours"), _tr(raw_config, "top_asset")]
        summary_values = [
            str(total_assets),
            _fmt_optional_number(total_hours, " h") or "n/a",
            dominant_asset["asset_tag"] if dominant_asset else "n/a",
        ]
        for cell, value in zip(summary_table.rows[0].cells, summary_headers):
            cell.text = value
        for cell, value in zip(summary_table.rows[1].cells, summary_values):
            cell.text = value

        document.add_paragraph(_tr(raw_config, "equipment_trend_snapshot")).runs[0].bold = True
        rank_table = document.add_table(rows=1, cols=3)
        rank_table.style = "Table Grid"
        for cell, header in zip(rank_table.rows[0].cells, [_tr(raw_config, "equipment_tag"), _tr(raw_config, "consumption"), _tr(raw_config, "unit")]):
            cell.text = header
        for asset in sorted(block["assets"], key=lambda item: item["quantity"], reverse=True)[:5]:
            row = rank_table.add_row().cells
            row[0].text = asset["asset_tag"]
            row[1].text = _fmt_money(asset["quantity"])
            row[2].text = asset["unit"]

        detail_table = document.add_table(rows=1, cols=7)
        detail_table.style = "Table Grid"
        for cell, header in zip(
            detail_table.rows[0].cells,
            [
                _tr(raw_config, "equipment_tag"),
                _tr(raw_config, "equipment_name"),
                _tr(raw_config, "emission_source"),
                _tr(raw_config, "fuel_type"),
                _tr(raw_config, "consumption"),
                _tr(raw_config, "unit"),
                _tr(raw_config, "operating_hours"),
            ],
        ):
            cell.text = header

        for asset in block["assets"]:
            row = detail_table.add_row().cells
            row[0].text = asset["asset_tag"]
            row[1].text = asset["equipment_name"]
            row[2].text = asset["emission_source"]
            row[3].text = asset["fuel"]
            row[4].text = _fmt_money(asset["quantity"])
            row[5].text = asset["unit"]
            row[6].text = _fmt_optional_number(asset["operating_hours"])

        document.add_paragraph(_tr(raw_config, "dashboard_summary_footer"))
        if block_index < len(blocks) - 1:
            document.add_page_break()

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def generate_bems_time_series_docx(raw_config: dict) -> bytes:
    blocks = _build_bems_trend_exports(raw_config)
    document = Document()
    core_props = document.core_properties
    core_props.title = raw_config.get("document", {}).get("title", _tr(raw_config, "bems_time_series_title"))
    core_props.subject = raw_config.get("document", {}).get("subject", "Scope 1 stationary combustion BEMS export")

    interval_minutes = _bems_interval_minutes(raw_config)
    for block_index, block in enumerate(blocks):
        document.add_heading(_tr(raw_config, "bems_time_series_title"), level=0)

        for line in [
            f"{_tr(raw_config, 'company')}: {block['company']}",
            f"{_tr(raw_config, 'site')}: {block['site']}",
            f"{_tr(raw_config, 'country')}: {block['country']}",
            f"{_tr(raw_config, 'reporting_period')}: {block['period_label']}",
        ]:
            if not line.endswith(": "):
                document.add_paragraph(line)

        summary_table = document.add_table(rows=2, cols=3)
        summary_table.style = "Table Grid"
        summary_headers = [_tr(raw_config, "assets"), _tr(raw_config, "interval"), _tr(raw_config, "rows")]
        summary_values = [
            str(len(block["assets"])),
            f"{interval_minutes} min",
            str(len(block["rows"])),
        ]
        for cell, value in zip(summary_table.rows[0].cells, summary_headers):
            cell.text = value
        for cell, value in zip(summary_table.rows[1].cells, summary_values):
            cell.text = value

        detail_table = document.add_table(rows=1, cols=6)
        detail_table.style = "Table Grid"
        for cell, header in zip(
            detail_table.rows[0].cells,
            [
                _tr(raw_config, "timestamp"),
                _tr(raw_config, "site"),
                _tr(raw_config, "equipment_tag"),
                _tr(raw_config, "sensor_name"),
                _tr(raw_config, "value"),
                _tr(raw_config, "unit"),
            ],
        ):
            cell.text = header

        for row_data in block["rows"]:
            row = detail_table.add_row().cells
            row[0].text = row_data["timestamp"].strftime("%Y-%m-%d %H:%M")
            row[1].text = row_data["site"]
            row[2].text = row_data["asset_tag"]
            row[3].text = row_data["sensor_name"]
            row[4].text = f"{row_data['value']:.2f}"
            row[5].text = row_data["unit"]

        document.add_paragraph(_tr(raw_config, "time_series_word_footer"))
        if block_index < len(blocks) - 1:
            document.add_page_break()

    output = BytesIO()
    document.save(output)
    return output.getvalue()


def generate_bems_time_series_xlsx(raw_config: dict) -> bytes:
    blocks = _build_bems_trend_exports(raw_config)
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    for sheet_index, block in enumerate(blocks, start=1):
        sheet_name = block["site"][:31] or f"{_tr(raw_config, 'site_fallback')} {sheet_index}"
        sheet = workbook.create_sheet(title=sheet_name)
        sheet["A1"] = _tr(raw_config, "bems_time_series_title")
        sheet["A1"].font = Font(size=14, bold=True)
        sheet["A2"] = _tr(raw_config, "company")
        sheet["B2"] = block["company"]
        sheet["A3"] = _tr(raw_config, "site")
        sheet["B3"] = block["site"]
        sheet["A4"] = _tr(raw_config, "country")
        sheet["B4"] = block["country"]
        sheet["A5"] = _tr(raw_config, "reporting_period")
        sheet["B5"] = block["period_label"]

        headers = [
            _tr(raw_config, "timestamp"),
            _tr(raw_config, "site"),
            _tr(raw_config, "equipment_tag"),
            _tr(raw_config, "sensor_name"),
            _tr(raw_config, "value"),
            _tr(raw_config, "unit"),
        ]
        header_fill = PatternFill(fill_type="solid", fgColor="1E5B88")
        for column_index, header in enumerate(headers, start=1):
            cell = sheet.cell(row=7, column=column_index, value=header)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        row_index = 8
        for row in block["rows"]:
            sheet.cell(row=row_index, column=1, value=row["timestamp"].strftime("%Y-%m-%d %H:%M"))
            sheet.cell(row=row_index, column=2, value=row["site"])
            sheet.cell(row=row_index, column=3, value=row["asset_tag"])
            sheet.cell(row=row_index, column=4, value=row["sensor_name"])
            sheet.cell(row=row_index, column=5, value=row["value"])
            sheet.cell(row=row_index, column=6, value=row["unit"])
            row_index += 1

        widths = [19, 20, 14, 24, 12, 8]
        for column_index, width in enumerate(widths, start=1):
            sheet.column_dimensions[get_column_letter(column_index)].width = width

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def generate_bems_time_series_csv(raw_config: dict) -> bytes:
    blocks = _build_bems_trend_exports(raw_config)
    buffer = StringIO()
    writer = csv.writer(buffer)

    for block_index, block in enumerate(blocks):
        if block_index > 0:
            writer.writerow([])
        writer.writerow([_tr(raw_config, "bems_time_series_title")])
        writer.writerow([_tr(raw_config, "company"), block["company"]])
        writer.writerow([_tr(raw_config, "site"), block["site"]])
        writer.writerow([_tr(raw_config, "country"), block["country"]])
        writer.writerow([_tr(raw_config, "reporting_period"), block["period_label"]])
        writer.writerow([])
        writer.writerow([
            _tr(raw_config, "timestamp"),
            _tr(raw_config, "site"),
            _tr(raw_config, "equipment_tag"),
            _tr(raw_config, "sensor_name"),
            _tr(raw_config, "value"),
            _tr(raw_config, "unit"),
        ])
        for row in block["rows"]:
            writer.writerow([
                row["timestamp"].strftime("%Y-%m-%d %H:%M"),
                row["site"],
                row["asset_tag"],
                row["sensor_name"],
                f"{row['value']:.2f}",
                row["unit"],
            ])

    return buffer.getvalue().encode("utf-8-sig")


def generate_bems_equipment_report_xlsx(raw_config: dict) -> bytes:
    blocks = _build_bems_site_blocks(raw_config)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = _tr(raw_config, "bems_summary_sheet_title")
    sheet["A1"] = _tr(raw_config, "bems_equipment_title")
    sheet["A1"].font = Font(size=14, bold=True)

    headers = [
        _tr(raw_config, "company"),
        _tr(raw_config, "site"),
        _tr(raw_config, "country"),
        _tr(raw_config, "reporting_period"),
        _tr(raw_config, "equipment_tag"),
        _tr(raw_config, "equipment_name"),
        _tr(raw_config, "emission_source"),
        _tr(raw_config, "fuel_type"),
        _tr(raw_config, "consumption"),
        _tr(raw_config, "unit"),
        _tr(raw_config, "operating_hours"),
    ]
    header_fill = PatternFill(fill_type="solid", fgColor="1E5B88")
    for column_index, header in enumerate(headers, start=1):
        cell = sheet.cell(row=3, column=column_index, value=header)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    row_index = 4
    for block in blocks:
        for asset in block["assets"]:
            values = [
                block["company"],
                block["site"],
                block["country"],
                block["period_label"],
                asset["asset_tag"],
                asset["equipment_name"],
                asset["emission_source"],
                asset["fuel"],
                float(asset["quantity"]),
                asset["unit"],
                None if asset["operating_hours"] is None else float(asset["operating_hours"]),
            ]
            for column_index, value in enumerate(values, start=1):
                sheet.cell(row=row_index, column=column_index, value=value)
            row_index += 1

    widths = [24, 20, 16, 18, 14, 22, 20, 16, 14, 8, 16]
    for column_index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column_index)].width = width

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def generate_bems_equipment_report_csv(raw_config: dict) -> bytes:
    blocks = _build_bems_site_blocks(raw_config)
    buffer = StringIO()
    writer = csv.writer(buffer)
    writer.writerow([
        _tr(raw_config, "company"),
        _tr(raw_config, "site"),
        _tr(raw_config, "country"),
        _tr(raw_config, "reporting_period"),
        _tr(raw_config, "equipment_tag"),
        _tr(raw_config, "equipment_name"),
        _tr(raw_config, "emission_source"),
        _tr(raw_config, "fuel_type"),
        _tr(raw_config, "consumption"),
        _tr(raw_config, "unit"),
        _tr(raw_config, "operating_hours"),
    ])
    for block in blocks:
        for asset in block["assets"]:
            writer.writerow([
                block["company"],
                block["site"],
                block["country"],
                block["period_label"],
                asset["asset_tag"],
                asset["equipment_name"],
                asset["emission_source"],
                asset["fuel"],
                f"{float(asset['quantity']):.2f}",
                asset["unit"],
                "" if asset["operating_hours"] is None else f"{float(asset['operating_hours']):.2f}",
            ])

    return buffer.getvalue().encode("utf-8-sig")
