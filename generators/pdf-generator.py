import argparse
import json
import math
import os
import random
from calendar import monthrange
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

from PIL import Image, ImageDraw, ImageFilter, ImageChops
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "pdf-generator.config.json")

DEFAULT_OUTPUT_DIR = "./outputs/utility-billing-2026"
DEFAULT_PDF_FILENAME = "district_heating_billing_2026_scanned_style_readable.pdf"
DEFAULT_BG_DIRNAME = "_page_backgrounds_v2"
DEFAULT_RANDOM_SEED = 20260325
DEFAULT_DOCUMENT_TITLE = "District Heating Billing Statement"
DEFAULT_DOCUMENT_AUTHOR = "Genspark Document Generator"
DEFAULT_DOCUMENT_SUBJECT = "Purchased Heat billing statements"
DEFAULT_COMPANY_STYLES = [
    {"accent": "#1E5B88", "accent_soft": "#DCEBF5", "skew": -0.22},
    {"accent": "#3F6F47", "accent_soft": "#E2EFE5", "skew": 0.18},
    {"accent": "#7B3247", "accent_soft": "#F1E2E8", "skew": -0.10},
]

TWOPLACES = Decimal("0.01")
PAGE_W, PAGE_H = A4

MONTH_FACTORS = {
    1: Decimal("1.30"),
    2: Decimal("1.20"),
    3: Decimal("1.05"),
    4: Decimal("0.90"),
    5: Decimal("0.78"),
    6: Decimal("0.62"),
    7: Decimal("0.56"),
    8: Decimal("0.60"),
    9: Decimal("0.76"),
    10: Decimal("0.93"),
    11: Decimal("1.12"),
    12: Decimal("1.27"),
}


def q2(value):
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    return value.quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def fmt_money(value):
    return f"£{q2(value):,.2f}"


def fmt_rate(value, places=3):
    if not isinstance(value, Decimal):
        value = Decimal(str(value))
    fmt = "1." + ("0" * places)
    return f"{value.quantize(Decimal(fmt), rounding=ROUND_HALF_UP):f}"


def parse_decimal(value):
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value))


def parse_date(value):
    return datetime.strptime(value, "%Y-%m-%d").date()


def register_fonts():
    regular_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    ]
    bold_candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
    ]
    reg = next((p for p in regular_candidates if os.path.exists(p)), None)
    bold = next((p for p in bold_candidates if os.path.exists(p)), None)
    if reg and bold:
        pdfmetrics.registerFont(TTFont("DocSans", reg))
        pdfmetrics.registerFont(TTFont("DocSansBold", bold))
        return "DocSans", "DocSansBold"
    return "Helvetica", "Helvetica-Bold"


FONT_REG, FONT_BOLD = register_fonts()
FONT_MONO = "Courier"


# ── translations ──────────────────────────────────────────────────────────────

TRANSLATIONS: dict[str, dict[str, str]] = {
    "en": {
        "logo_subtitle":       "Metered District Heating Services",
        "doc_title_heading":   "Heating Billing Document",
        "doc_subtitle":        "District Heating",
        "box_supplier":        "Supplier Details",
        "box_customer":        "Customer / Service Address",
        "box_invoice":         "Invoice Metadata",
        "meta_invoice_no":     "Invoice Number",
        "meta_issue_date":     "Issue Date",
        "meta_due_date":       "Payment Due Date",
        "meta_currency":       "Currency",
        "tbl_header_field":    "Billing Field",
        "tbl_header_value":    "Recorded Value",
        "row_supplier":        "Supplier",
        "row_customer":        "Customer Name",
        "row_site":            "Site",
        "row_city":            "City",
        "row_postcode":        "Postcode",
        "row_period_start":    "Billing Period Start",
        "row_period_end":      "Billing Period End",
        "row_meter_id":        "Heat Meter ID",
        "row_prev_read":       "Previous Meter Reading (kWh)",
        "row_curr_read":       "Current Meter Reading (kWh)",
        "row_consumption":     "Heat Consumption (kWh)",
        "row_unit_price":      "Heat Unit Price (£/kWh)",
        "row_capacity":        "Contracted Capacity (kW)",
        "row_capacity_rate":   "Capacity Charge (£/kW/month)",
        "row_supplier_ef":     "Supplier Emission Factor (kg CO\u2082e/kWh)",
        "box_charges":         "Charges & VAT Summary",
        "charge_heat":         "Heat Consumption Cost",
        "charge_capacity":     "Capacity Charge",
        "charge_subtotal":     "Subtotal",
        "charge_vat":          "VAT (5%)",
        "charge_total":        "Total Amount Due",
        "footer_vat":          "VAT applied at 5%. Payment terms: 14 days from issue date unless otherwise specified in the supply agreement.",
        "footer_page":         "Page {page} / {total}",
    },
    "fr": {
        "logo_subtitle":       "Services de chauffage urbain mesurés",
        "doc_title_heading":   "Document de facturation thermique",
        "doc_subtitle":        "Chauffage urbain",
        "box_supplier":        "Détails du fournisseur",
        "box_customer":        "Client / Adresse du service",
        "box_invoice":         "Métadonnées de la facture",
        "meta_invoice_no":     "Numéro de facture",
        "meta_issue_date":     "Date d'émission",
        "meta_due_date":       "Date d'échéance",
        "meta_currency":       "Devise",
        "tbl_header_field":    "Champ de facturation",
        "tbl_header_value":    "Valeur enregistrée",
        "row_supplier":        "Fournisseur",
        "row_customer":        "Nom du client",
        "row_site":            "Site",
        "row_city":            "Ville",
        "row_postcode":        "Code postal",
        "row_period_start":    "Début de la période",
        "row_period_end":      "Fin de la période",
        "row_meter_id":        "Identifiant du compteur",
        "row_prev_read":       "Relevé précédent (kWh)",
        "row_curr_read":       "Relevé actuel (kWh)",
        "row_consumption":     "Consommation thermique (kWh)",
        "row_unit_price":      "Prix unitaire (£/kWh)",
        "row_capacity":        "Capacité contractée (kW)",
        "row_capacity_rate":   "Frais de capacité (£/kW/mois)",
        "row_supplier_ef":     "Facteur d'émission fournisseur (kg CO\u2082e/kWh)",
        "box_charges":         "Résumé des charges et TVA",
        "charge_heat":         "Coût de consommation thermique",
        "charge_capacity":     "Frais de capacité",
        "charge_subtotal":     "Sous-total",
        "charge_vat":          "TVA (5%)",
        "charge_total":        "Montant total dû",
        "footer_vat":          "TVA appliquée à 5 %. Conditions de paiement : 14 jours à compter de la date d'émission, sauf accord contraire.",
        "footer_page":         "Page {page} / {total}",
    },
    "de": {
        "logo_subtitle":       "Gemessene Fernwärmedienste",
        "doc_title_heading":   "Fernwärme-Abrechnungsdokument",
        "doc_subtitle":        "Fernwärme",
        "box_supplier":        "Lieferantendetails",
        "box_customer":        "Kunde / Serviceadresse",
        "box_invoice":         "Rechnungsmetadaten",
        "meta_invoice_no":     "Rechnungsnummer",
        "meta_issue_date":     "Ausstellungsdatum",
        "meta_due_date":       "Zahlungsfälligkeitsdatum",
        "meta_currency":       "Währung",
        "tbl_header_field":    "Abrechnungsfeld",
        "tbl_header_value":    "Erfasster Wert",
        "row_supplier":        "Lieferant",
        "row_customer":        "Kundenname",
        "row_site":            "Standort",
        "row_city":            "Stadt",
        "row_postcode":        "Postleitzahl",
        "row_period_start":    "Abrechnungszeitraum Beginn",
        "row_period_end":      "Abrechnungszeitraum Ende",
        "row_meter_id":        "Wärmezähler-ID",
        "row_prev_read":       "Vorheriger Zählerstand (kWh)",
        "row_curr_read":       "Aktueller Zählerstand (kWh)",
        "row_consumption":     "Wärmeverbrauch (kWh)",
        "row_unit_price":      "Einheitspreis (£/kWh)",
        "row_capacity":        "Vertragsleistung (kW)",
        "row_capacity_rate":   "Leistungsgebühr (£/kW/Monat)",
        "row_supplier_ef":     "Emissionsfaktor Lieferant (kg CO\u2082e/kWh)",
        "box_charges":         "Kosten- und MwSt.-Übersicht",
        "charge_heat":         "Wärmeverbrauchskosten",
        "charge_capacity":     "Leistungsgebühr",
        "charge_subtotal":     "Zwischensumme",
        "charge_vat":          "MwSt. (5%)",
        "charge_total":        "Gesamtbetrag fällig",
        "footer_vat":          "MwSt. zu 5 % angewendet. Zahlungsbedingungen: 14 Tage ab Ausstellungsdatum, sofern im Liefervertrag nichts anderes angegeben.",
        "footer_page":         "Seite {page} / {total}",
    },
    "nl": {
        "logo_subtitle":       "Gemeten stadsverwarmingsdiensten",
        "doc_title_heading":   "Factuur stadsverwarming",
        "doc_subtitle":        "Stadsverwarming",
        "box_supplier":        "Leveranciersgegevens",
        "box_customer":        "Klant / Serviceadres",
        "box_invoice":         "Factuurmetadata",
        "meta_invoice_no":     "Factuurnummer",
        "meta_issue_date":     "Uitgiftedatum",
        "meta_due_date":       "Vervaldatum",
        "meta_currency":       "Valuta",
        "tbl_header_field":    "Factuurveld",
        "tbl_header_value":    "Geregistreerde waarde",
        "row_supplier":        "Leverancier",
        "row_customer":        "Klantnaam",
        "row_site":            "Locatie",
        "row_city":            "Stad",
        "row_postcode":        "Postcode",
        "row_period_start":    "Begin facturatieperiode",
        "row_period_end":      "Einde facturatieperiode",
        "row_meter_id":        "Warmtemeter-ID",
        "row_prev_read":       "Vorige meterstand (kWh)",
        "row_curr_read":       "Huidige meterstand (kWh)",
        "row_consumption":     "Warmteverbruik (kWh)",
        "row_unit_price":      "Eenheidsprijs (£/kWh)",
        "row_capacity":        "Gecontracteerd vermogen (kW)",
        "row_capacity_rate":   "Vermogenstoeslag (£/kW/maand)",
        "row_supplier_ef":     "Emissiefactor leverancier (kg CO\u2082e/kWh)",
        "box_charges":         "Kosten- en BTW-overzicht",
        "charge_heat":         "Warmteverbruikskosten",
        "charge_capacity":     "Vermogenstoeslag",
        "charge_subtotal":     "Subtotaal",
        "charge_vat":          "BTW (5%)",
        "charge_total":        "Totaal verschuldigd bedrag",
        "footer_vat":          "BTW 5 % toegepast. Betalingsvoorwaarden: 14 dagen na factuurdatum, tenzij anders vermeld in de leveringsovereenkomst.",
        "footer_page":         "Pagina {page} / {total}",
    },
}

LANGUAGE_LABELS: dict[str, str] = {
    "English":             "en",
    "French (Français)":   "fr",
    "German (Deutsch)":    "de",
    "Dutch (Nederlands)":  "nl",
}


def parse_args():
    parser = argparse.ArgumentParser(description="Generate district heating billing PDFs from config.")
    parser.add_argument(
        "--per-company",
        action="store_true",
        help="Generate one PDF per company instead of a single combined PDF.",
    )
    parser.add_argument(
        "--company",
        action="append",
        default=[],
        help="Generate only the named company label. Repeat to include multiple companies.",
    )
    return parser.parse_args()


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
        config = json.load(fh)

    document = config.get("document", {})
    output_dir = document.get("output_dir", DEFAULT_OUTPUT_DIR)
    pdf_filename = document.get("pdf_filename", DEFAULT_PDF_FILENAME)
    bg_dirname = document.get("background_dirname", DEFAULT_BG_DIRNAME)
    pdf_path = os.path.join(output_dir, pdf_filename)
    bg_dir = os.path.join(output_dir, bg_dirname)

    financial_period = config.get("financial_period", {})
    if "label" not in financial_period or "start_date" not in financial_period or "end_date" not in financial_period:
        raise ValueError("financial_period must define label, start_date, and end_date")

    normalized_config = {
        "random_seed": int(config.get("random_seed", DEFAULT_RANDOM_SEED)),
        "document": {
            "output_dir": output_dir,
            "pdf_filename": pdf_filename,
            "pdf_path": pdf_path,
            "background_dir": bg_dir,
            "title": document.get("title", DEFAULT_DOCUMENT_TITLE),
            "subject": document.get("subject", DEFAULT_DOCUMENT_SUBJECT),
        },
        "financial_period": {
            "label": financial_period["label"],
            "start_date": parse_date(financial_period["start_date"]),
            "end_date": parse_date(financial_period["end_date"]),
        },
    }

    companies = config.get("companies", [])
    if not companies:
        raise ValueError("Configuration must include at least one company")

    normalized_config["companies"] = [
        normalize_company(company, normalized_config["financial_period"], idx) for idx, company in enumerate(companies)
    ]
    return normalized_config


def normalize_company(company, financial_period, company_index):
    sites = company.get("sites", [])
    if not sites:
        raise ValueError(f"Company {company.get('label', '<unknown>')} must include at least one site")

    style = DEFAULT_COMPANY_STYLES[company_index % len(DEFAULT_COMPANY_STYLES)]

    return {
        "label": company["label"],
        "supplier": company["supplier"],
        "supplier_code": company["supplier_code"],
        "supplier_address": company["supplier_address"],
        "customer": company["customer"],
        "customer_code": company["customer_code"],
        "accent": company.get("accent", style["accent"]),
        "accent_soft": company.get("accent_soft", style["accent_soft"]),
        "skew": float(company.get("skew", style["skew"])),
        "currency": company.get("currency", "GBP (£)"),
        "sites": [normalize_site(company, site, financial_period) for site in sites],
    }


def normalize_site(company, site, financial_period):
    billing_periods = site.get("billing_periods", company.get("billing_periods"))
    if billing_periods is None:
        billing_periods = derive_month_periods(financial_period["start_date"], financial_period["end_date"])

    normalized_periods = normalize_billing_periods(billing_periods, financial_period["start_date"].year)
    billing_period_count = site.get("billing_period_count", company.get("billing_period_count"))
    if billing_period_count is not None:
        normalized_periods = normalized_periods[: int(billing_period_count)]

    if not normalized_periods:
        raise ValueError(f"Site {site.get('label', site.get('meter_id', '<unknown>'))} has no billing periods")

    return {
        "label": site.get("label", site["meter_id"]),
        "customer": site.get("customer", company["customer"]),
        "customer_code": site.get("customer_code", company["customer_code"]),
        "customer_address": site["customer_address"],
        "city": site["city"],
        "postcode": site["postcode"],
        "meter_id": site["meter_id"],
        "capacity_kw": int(site["capacity_kw"]),
        "capacity_rate": parse_decimal(site["capacity_rate"]),
        "supplier_ef":   parse_decimal(site.get("supplier_ef", "0")),
        "base_consumption": int(site["base_consumption"]),
        "unit_price_base": parse_decimal(site["unit_price_base"]),
        "start_reading": int(site["start_reading"]),
        "billing_periods": normalized_periods,
    }


def derive_month_periods(start_date, end_date):
    periods = []
    current = date(start_date.year, start_date.month, 1)
    while current <= end_date:
        periods.append({"year": current.year, "month": current.month})
        if current.month == 12:
            current = date(current.year + 1, 1, 1)
        else:
            current = date(current.year, current.month + 1, 1)
    return periods


def normalize_billing_periods(periods, default_year):
    normalized = []
    for period in periods:
        if isinstance(period, int):
            normalized.append({"year": default_year, "month": period})
            continue
        if "month" in period:
            normalized.append({
                "year": int(period.get("year", default_year)),
                "month": int(period["month"]),
                "label": period.get("label"),
                "invoice_suffix": period.get("invoice_suffix"),
            })
            continue
        if "start_date" in period and "end_date" in period:
            normalized.append({
                "start_date": parse_date(period["start_date"]),
                "end_date": parse_date(period["end_date"]),
                "label": period.get("label"),
                "invoice_suffix": period.get("invoice_suffix"),
            })
            continue
        raise ValueError(f"Unsupported billing period definition: {period}")
    return normalized


def billing_period_dates(period):
    if "month" in period:
        year = period["year"]
        month = period["month"]
        start = date(year, month, 1)
        end = date(year, month, monthrange(year, month)[1])
        return start, end
    return period["start_date"], period["end_date"]


def billing_period_label(period):
    if period.get("label"):
        return period["label"]
    if "month" in period:
        return date(period["year"], period["month"], 1).strftime("%B %Y")
    start, end = billing_period_dates(period)
    return f"{start.strftime('%d %b %Y')} - {end.strftime('%d %b %Y')}"


def billing_period_factor(period):
    if "month" in period:
        return MONTH_FACTORS[period["month"]]

    start, end = billing_period_dates(period)
    days = (end - start).days + 1
    midpoint = start + timedelta(days=days // 2)
    base_factor = MONTH_FACTORS[midpoint.month]
    month_days = monthrange(midpoint.year, midpoint.month)[1]
    duration_factor = Decimal(str(days / month_days))
    return max(Decimal("0.35"), base_factor * duration_factor)


def invoice_suffix(period, index):
    if period.get("invoice_suffix"):
        return str(period["invoice_suffix"])
    if "month" in period:
        return f"{period['year']}-{period['month']:02d}"
    return f"P{index:02d}"


def generate_billing_records(company, site):
    prev = site["start_reading"]
    records = []
    for index, period in enumerate(site["billing_periods"], start=1):
        first, last = billing_period_dates(period)
        factor = billing_period_factor(period)
        variation = Decimal(str(1 + random.uniform(-0.05, 0.05)))
        consumption = int((Decimal(site["base_consumption"]) * factor * variation).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
        consumption = max(5000, min(50000, consumption))
        curr = prev + consumption

        midpoint = first + timedelta(days=((last - first).days // 2))
        seasonal = Decimal("0.004") if midpoint.month in (1, 2, 11, 12) else Decimal("0.000")
        summer = Decimal("-0.002") if midpoint.month in (6, 7, 8) else Decimal("0.000")
        random_adjust = Decimal(str(round(random.uniform(-0.003, 0.003), 4)))
        unit_price = site["unit_price_base"] + seasonal + summer + random_adjust
        unit_price = min(Decimal("0.120"), max(Decimal("0.040"), unit_price)).quantize(Decimal("1.000"), rounding=ROUND_HALF_UP)

        heat_cost = q2(Decimal(consumption) * unit_price)
        capacity_charge = q2(Decimal(site["capacity_kw"]) * site["capacity_rate"])
        subtotal = q2(heat_cost + capacity_charge)
        vat = q2(subtotal * Decimal("0.05"))
        total = q2(subtotal + vat)

        issue_date = last + timedelta(days=4)
        due_date = issue_date + timedelta(days=14)
        invoice_no = f"{company['supplier_code']}-{site['customer_code']}-{invoice_suffix(period, index)}"

        records.append({
            "billing_period_label": billing_period_label(period),
            "supplier": company["supplier"],
            "customer": site["customer"],
            "site_label": site["label"],
            "city": site["city"],
            "postcode": site["postcode"],
            "period_start": first,
            "period_end": last,
            "meter_id": site["meter_id"],
            "prev_read": prev,
            "curr_read": curr,
            "consumption": consumption,
            "unit_price": unit_price,
            "heat_cost": heat_cost,
            "capacity_kw": site["capacity_kw"],
            "capacity_rate": site["capacity_rate"],
            "supplier_ef":   site["supplier_ef"],
            "capacity_charge": capacity_charge,
            "subtotal": subtotal,
            "vat": vat,
            "total": total,
            "invoice_no": invoice_no,
            "issue_date": issue_date,
            "due_date": due_date,
        })
        prev = curr
    return records


def build_background(path, accent="#355C7D", seed=1, width=1240, height=1754, noise_level=1.0):
    rng = random.Random(seed)
    img = Image.new("RGB", (width, height), (247, 246, 242))
    if noise_level > 0:
        noise1 = Image.effect_noise((width, height), 8 * noise_level).convert("L")
        noise2 = Image.effect_noise((width, height), 14 * noise_level).convert("L")
        noise1 = ImageChops.add_modulo(noise1, Image.new("L", (width, height), 126))
        noise2 = ImageChops.add_modulo(noise2, Image.new("L", (width, height), 120))
        tint = Image.merge("RGB", (noise1, noise1, noise2)).filter(ImageFilter.GaussianBlur(0.55))
        img = Image.blend(img, tint, 0.08)

    draw = ImageDraw.Draw(img, "RGBA")
    if noise_level > 0:
        for y in range(0, height, 5):
            alpha = int((4 + int(3 * math.sin(y / 39.0))) * noise_level)
            draw.line([(0, y), (width, y)], fill=(110, 110, 110, alpha), width=1)

        for x in range(0, 45):
            alpha = int(20 * (1 - x / 45) * noise_level)
            draw.line([(x, 0), (x, height)], fill=(125, 125, 125, alpha), width=1)
        for x in range(width - 25, width):
            alpha = int(12 * ((x - (width - 25)) / 25) * noise_level)
            draw.line([(x, 0), (x, height)], fill=(125, 125, 125, alpha), width=1)

    accent_rgb = tuple(int(accent[i:i + 2], 16) for i in (1, 3, 5))
    overlay = Image.new("RGBA", (width, height), (0, 0, 0, 0))
    overlay_draw = ImageDraw.Draw(overlay, "RGBA")
    overlay_draw.rectangle((0, 0, width, 140), fill=accent_rgb + (int(16 * noise_level),))
    overlay_draw.rectangle((0, height - 70, width, height), fill=(140, 140, 140, int(8 * noise_level)))
    overlay = overlay.filter(ImageFilter.GaussianBlur(10))
    img = Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")

    if noise_level > 0:
        n_spots = int(35 * noise_level)
        for _ in range(n_spots):
            x = rng.randint(0, width - 1)
            y = rng.randint(0, height - 1)
            radius = rng.randint(1, 4)
            draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=(90, 90, 90, rng.randint(8, 18)))

    if noise_level > 0:
        img = img.filter(ImageFilter.GaussianBlur(0.28 * noise_level))
    img.save(path, quality=72)


def build_foreground_noise(path, seed, width=1240, height=1754, noise_level=1.0):
    """Build a transparent RGBA foreground noise overlay for scan simulation.

    At noise_level=0 the image is fully transparent (no effect on drawings).
    """
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))

    if noise_level > 0:
        rng = random.Random(seed + 500)
        draw = ImageDraw.Draw(img)

        # Random dark speckles
        n_speckles = int(600 * noise_level)
        for _ in range(n_speckles):
            x = rng.randint(0, width - 1)
            y = rng.randint(0, height - 1)
            dark = rng.randint(40, 110)
            alpha = int(rng.randint(20, 65) * noise_level)
            radius = rng.randint(0, 1)
            fill = (dark, dark, dark, alpha)
            if radius == 0:
                draw.point((x, y), fill=fill)
            else:
                draw.ellipse(
                    (x - radius, y - radius, x + radius, y + radius), fill=fill
                )

        # Occasional faint horizontal scan-line artefacts
        y = rng.randint(5, 30)
        while y < height:
            alpha = int(rng.randint(8, 28) * noise_level)
            draw.line([(0, y), (width, y)], fill=(60, 60, 60, alpha), width=1)
            y += rng.randint(40, 180)

    img.save(path, format="PNG")


def draw_logo(c, x, y, accent, supplier_name, strings):
    acc = HexColor(accent)
    c.saveState()
    c.setFillColor(acc)
    c.circle(x + 14, y + 12, 11, stroke=0, fill=1)
    c.setFillColor(white)
    c.rect(x + 8.7, y + 9.5, 10.8, 2.0, stroke=0, fill=1)
    c.rect(x + 10.5, y + 13.5, 7.3, 2.0, stroke=0, fill=1)
    c.rect(x + 12.1, y + 17.5, 4.1, 2.0, stroke=0, fill=1)
    c.setFillColor(acc)
    c.setFont(FONT_BOLD, 14)
    c.drawString(x + 32, y + 10, supplier_name)
    c.setFont(FONT_REG, 7.4)
    c.drawString(x + 32, y + 1, strings["logo_subtitle"])
    c.restoreState()


def round_box(c, x, y, w, h, stroke="#B8BEC5", fill="#FFFFFF", radius=4):
    c.setStrokeColor(HexColor(stroke))
    c.setFillColor(HexColor(fill))
    c.roundRect(x, y, w, h, radius, stroke=1, fill=1)


def draw_info_box(c, x, y, w, h, title, lines, accent, accent_soft):
    c.saveState()
    round_box(c, x, y, w, h)
    c.setFillColor(HexColor(accent_soft))
    c.roundRect(x, y + h - 16, w, 16, 4, stroke=0, fill=1)
    c.rect(x, y + h - 16, w, 10, stroke=0, fill=1)
    c.setFillColor(HexColor(accent))
    c.setFont(FONT_BOLD, 8.2)
    c.drawString(x + 8, y + h - 11, title)
    c.setFillColor(HexColor("#202428"))
    c.setFont(FONT_REG, 7.2)
    ty = y + h - 28
    for line in lines:
        c.drawString(x + 8, ty, str(line))
        ty -= 10
    c.restoreState()


def draw_table(c, x, y, w, rows, accent, accent_soft, strings, row_h=18):
    """rows: list of (label, value) or (label, value, use_monospace) tuples."""
    c.saveState()
    total_h = row_h * (len(rows) + 1)
    round_box(c, x, y, w, total_h)
    c.setFillColor(HexColor(accent_soft))
    c.rect(x, y + total_h - row_h, w, row_h, stroke=0, fill=1)
    split = x + w * 0.56
    c.setStrokeColor(HexColor("#D5DADF"))
    for i in range(len(rows) + 1):
        yy = y + i * row_h
        c.line(x, yy, x + w, yy)
    c.line(split, y, split, y + total_h)

    c.setFillColor(HexColor(accent))
    c.setFont(FONT_BOLD, 8.1)
    c.drawString(x + 8, y + total_h - 12.2, strings["tbl_header_field"])
    c.drawString(split + 8, y + total_h - 12.2, strings["tbl_header_value"])

    cy = y + total_h - row_h - 12.5
    for row in rows:
        field, value = row[0], row[1]
        use_mono = row[2] if len(row) > 2 else False
        c.setFillColor(HexColor("#5A6066"))
        c.setFont(FONT_REG, 7.5)
        c.drawString(x + 8, cy, field)
        c.setFillColor(HexColor("#1F2328"))
        c.setFont(FONT_MONO if use_mono else FONT_REG, 8.0)
        val = str(value)
        if len(val) > 58:
            val = val[:57] + "…"
        c.drawString(split + 8, cy, val)
        cy -= row_h
    c.restoreState()


def draw_amounts_box(c, x, y, w, h, rec, accent, accent_soft, strings):
    c.saveState()
    draw_info_box(c, x, y, w, h, strings["box_charges"], [], accent, accent_soft)
    lines = [
        (strings["charge_heat"],     fmt_money(rec["heat_cost"])),
        (strings["charge_capacity"], fmt_money(rec["capacity_charge"])),
        (strings["charge_subtotal"], fmt_money(rec["subtotal"])),
        (strings["charge_vat"],      fmt_money(rec["vat"])),
        (strings["charge_total"],    fmt_money(rec["total"])),
    ]
    sy = y + h - 30
    for idx, (label, value) in enumerate(lines):
        if idx == 4:
            c.setFillColor(HexColor(accent_soft))
            c.roundRect(x + 8, sy - 8, w - 16, 18, 3, stroke=0, fill=1)
            c.setFillColor(HexColor(accent))
            c.setFont(FONT_BOLD, 8.5)
            c.drawString(x + 14, sy + 3, label)
            c.drawRightString(x + w - 14, sy + 3, value)
        else:
            c.setFillColor(HexColor("#5A6066"))
            c.setFont(FONT_REG, 7.5)
            c.drawString(x + 12, sy, label)
            c.setFillColor(HexColor("#1F2328"))
            c.drawRightString(x + w - 12, sy, value)
        sy -= 18
    c.restoreState()


def draw_billing_invoice(c, company, site, rec, page_no, total_pages, bg_path, fg_overlay_path, financial_period_label, strings, noise_level=1.0):
    accent = company["accent"]
    accent_soft = company["accent_soft"]
    margin = 32

    c.drawImage(ImageReader(bg_path), 0, 0, width=PAGE_W, height=PAGE_H, mask="auto")
    c.saveState()
    c.translate(PAGE_W / 2, PAGE_H / 2)
    effective_skew = company["skew"] * noise_level
    jitter = random.choice([-0.04, 0.03, 0.05]) * noise_level
    c.rotate(effective_skew + jitter)
    c.translate(-PAGE_W / 2, -PAGE_H / 2)

    draw_logo(c, margin, PAGE_H - 72, accent, company["supplier"], strings)
    c.setFillColor(HexColor("#1E2328"))
    c.setFont(FONT_BOLD, 15)
    c.drawRightString(PAGE_W - margin, PAGE_H - 50, strings["doc_title_heading"])
    c.setFont(FONT_REG, 8.2)
    c.drawRightString(PAGE_W - margin, PAGE_H - 64, strings["doc_subtitle"])
    c.setFont(FONT_REG, 7.2)
    c.drawRightString(PAGE_W - margin, PAGE_H - 77, f"{financial_period_label} • {company['label']} • {rec['billing_period_label']}")

    top_y = PAGE_H - 170
    draw_info_box(c, margin, top_y, 240, 92, strings["box_supplier"], company["supplier_address"], accent, accent_soft)
    draw_info_box(c, margin + 252, top_y, 240, 92, strings["box_customer"], site["customer_address"], accent, accent_soft)

    meta_lines = [
        f"{strings['meta_invoice_no']}: {rec['invoice_no']}",
        f"{strings['meta_issue_date']}: {rec['issue_date'].strftime('%d %b %Y')}",
        f"{strings['meta_due_date']}: {rec['due_date'].strftime('%d %b %Y')}",
        f"{strings['meta_currency']}: {company['currency']}",
    ]
    draw_info_box(c, margin, top_y - 108, PAGE_W - margin * 2, 84, strings["box_invoice"], meta_lines, accent, accent_soft)

    fields = [
        (strings["row_supplier"],      rec["supplier"]),
        (strings["row_customer"],      rec["customer"]),
        (strings["row_site"],          rec["site_label"]),
        (strings["row_city"],          rec["city"]),
        (strings["row_postcode"],      rec["postcode"]),
        (strings["row_period_start"],  rec["period_start"].strftime("%d %b %Y")),
        (strings["row_period_end"],    rec["period_end"].strftime("%d %b %Y")),
        (strings["row_meter_id"],      rec["meter_id"], True),  # monospace
        (strings["row_prev_read"],     f"{rec['prev_read']:,}"),
        (strings["row_curr_read"],     f"{rec['curr_read']:,}"),
        (strings["row_consumption"],   f"{rec['consumption']:,}"),
        (strings["row_unit_price"],    fmt_rate(rec["unit_price"], 3)),
        (strings["row_capacity"],      f"{rec['capacity_kw']}"),
        (strings["row_capacity_rate"], fmt_rate(rec["capacity_rate"], 2)),
        (strings["row_supplier_ef"],   fmt_rate(rec["supplier_ef"], 4)),
    ]
    table_y = 171
    draw_table(c, margin, table_y, PAGE_W - margin * 2, fields, accent, accent_soft, strings, row_h=17)

    draw_amounts_box(c, margin, 62, PAGE_W - margin * 2, 95, rec, accent, accent_soft, strings)

    c.setStrokeColor(HexColor("#C9CDD2"))
    c.line(margin, 42, PAGE_W - margin, 42)
    c.setFillColor(HexColor("#5A6066"))
    c.setFont(FONT_REG, 6.4)
    c.drawString(margin, 25, strings["footer_vat"])
    c.drawRightString(PAGE_W - margin, 25, strings["footer_page"].format(page=page_no, total=total_pages))

    c.restoreState()
    if noise_level > 0:
        c.drawImage(ImageReader(fg_overlay_path), 0, 0, width=PAGE_W, height=PAGE_H, mask="auto")
    c.showPage()


def validate_records(records):
    for record in records:
        assert record["curr_read"] - record["prev_read"] == record["consumption"]
        assert 5000 <= record["consumption"] <= 50000
        assert Decimal("0.04") <= record["unit_price"] <= Decimal("0.12")
        assert 50 <= record["capacity_kw"] <= 500
        assert record["heat_cost"] == q2(Decimal(record["consumption"]) * record["unit_price"])
        assert record["capacity_charge"] == q2(Decimal(record["capacity_kw"]) * record["capacity_rate"])
        assert record["subtotal"] == q2(record["heat_cost"] + record["capacity_charge"])
        assert record["vat"] == q2(record["subtotal"] * Decimal("0.05"))
        assert record["total"] == q2(record["subtotal"] + record["vat"])


def build_sections(config):
    random.seed(config["random_seed"])
    sections = []
    for company in config["companies"]:
        for site in company["sites"]:
            records = generate_billing_records(company, site)
            validate_records(records)
            sections.append({
                "company": company,
                "site": site,
                "records": records,
            })
    return sections


def slugify(value):
    normalized = "".join(ch.lower() if ch.isalnum() else "-" for ch in value)
    collapsed = "-".join(part for part in normalized.split("-") if part)
    return collapsed or "company"


def filtered_config(config, company_labels):
    if not company_labels:
        return config

    available_labels = {company["label"] for company in config["companies"]}
    missing = [label for label in company_labels if label not in available_labels]
    if missing:
        raise ValueError(f"Unknown company labels: {', '.join(missing)}")

    filtered = dict(config)
    filtered["companies"] = [company for company in config["companies"] if company["label"] in company_labels]
    return filtered


def output_path_for_company(config, company):
    pdf_filename = config["document"]["pdf_filename"]
    base, ext = os.path.splitext(pdf_filename)
    filename = f"{base}-{slugify(company['label'])}{ext or '.pdf'}"
    return os.path.join(config["document"]["output_dir"], filename)


def render_pdf(config, sections, output_path, noise_level=1.0):
    lang = config["document"].get("language", "en")
    strings = TRANSLATIONS.get(lang, TRANSLATIONS["en"])

    total_pages = sum(len(section["records"]) for section in sections)

    backgrounds = []
    fg_overlays = []
    page_index = 1
    for section in sections:
        for _ in section["records"]:
            bg_dir = config["document"]["background_dir"]
            bg_path = os.path.join(bg_dir, f"bg_{page_index:02d}.jpg")
            fg_path = os.path.join(bg_dir, f"fg_{page_index:02d}.png")
            build_background(bg_path, accent=section["company"]["accent"], seed=900 + page_index, noise_level=noise_level)
            build_foreground_noise(fg_path, seed=900 + page_index, noise_level=noise_level)
            backgrounds.append(bg_path)
            fg_overlays.append(fg_path)
            page_index += 1

    c = canvas.Canvas(output_path, pagesize=A4, pageCompression=1)
    c.setTitle(config["document"]["title"])
    c.setSubject(config["document"]["subject"])

    page_no = 1
    bg_idx = 0
    for section in sections:
        for record in section["records"]:
            draw_billing_invoice(
                c,
                section["company"],
                section["site"],
                record,
                page_no,
                total_pages,
                backgrounds[bg_idx],
                fg_overlays[bg_idx],
                config["financial_period"]["label"],
                strings,
                noise_level=noise_level,
            )
            page_no += 1
            bg_idx += 1

    c.save()
    return output_path


def generate_pdf(per_company=False, company_labels=None):
    config = load_config()
    config = filtered_config(config, company_labels or [])
    os.makedirs(config["document"]["output_dir"], exist_ok=True)
    os.makedirs(config["document"]["background_dir"], exist_ok=True)

    if per_company:
        output_paths = []
        for company in config["companies"]:
            company_config = dict(config)
            company_config["companies"] = [company]
            sections = build_sections(company_config)
            output_paths.append(render_pdf(company_config, sections, output_path_for_company(company_config, company)))
        return output_paths

    sections = build_sections(config)
    return [render_pdf(config, sections, config["document"]["pdf_path"])]


if __name__ == "__main__":
    args = parse_args()
    for output_path in generate_pdf(per_company=args.per_company, company_labels=args.company):
        print(output_path)
