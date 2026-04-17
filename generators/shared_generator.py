from __future__ import annotations

import math
import os
import random
from calendar import monthrange
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP

from PIL import Image, ImageChops, ImageDraw, ImageFilter
from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

TWOPLACES = Decimal("0.01")
PAGE_W, PAGE_H = A4

DEFAULT_COMPANY_STYLES = [
    {"accent": "#1E5B88", "accent_soft": "#DCEBF5", "skew": -0.22},
    {"accent": "#3F6F47", "accent_soft": "#E2EFE5", "skew": 0.18},
    {"accent": "#7B3247", "accent_soft": "#F1E2E8", "skew": -0.10},
]

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
        for _ in range(int(35 * noise_level)):
            x = rng.randint(0, width - 1)
            y = rng.randint(0, height - 1)
            radius = rng.randint(1, 4)
            draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=(90, 90, 90, rng.randint(8, 18)))

    if noise_level > 0:
        img = img.filter(ImageFilter.GaussianBlur(0.28 * noise_level))
    img.save(path, quality=72)


def build_foreground_noise(path, seed, width=1240, height=1754, noise_level=1.0):
    img = Image.new("RGBA", (width, height), (0, 0, 0, 0))

    if noise_level > 0:
        rng = random.Random(seed + 500)
        draw = ImageDraw.Draw(img)

        for _ in range(int(600 * noise_level)):
            x = rng.randint(0, width - 1)
            y = rng.randint(0, height - 1)
            dark = rng.randint(40, 110)
            alpha = int(rng.randint(20, 65) * noise_level)
            radius = rng.randint(0, 1)
            fill = (dark, dark, dark, alpha)
            if radius == 0:
                draw.point((x, y), fill=fill)
            else:
                draw.ellipse((x - radius, y - radius, x + radius, y + radius), fill=fill)

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
