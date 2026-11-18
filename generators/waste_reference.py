"""Static reference data for Scope 3 Category 5 – Waste generated in operations.

Enum values mirror ``scope3Waste.types.ts`` (waste management spike) so generated ground
truth can be scored directly against the extraction payload.

Sources
-------
* EWC → DESNZ material mapping: ``ewc_codes.txt`` proposal (UK-retained List of Wastes).
* Densities (t/m³): Environment Agency "Density conversion factors for waste" – one factor per
  EWC code (17 09 04 revised from 0.42 to 0.32 in January 2014).
* R/D → route: ``rd_codes.txt`` proposal, Waste Framework Directive Annex I / II.
* Landfill tax: HMRC / Revenue Scotland / WRA published rates.

All organisations, permits and registration numbers below are fictional.
"""
from __future__ import annotations

from datetime import date

SCOPE_LABEL = "Scope 3"
CATEGORY_WASTE = "waste_generated_in_operations"

CUBIC_YARD_TO_CUBIC_METRE = 0.764555
WEIGHT_MISMATCH_TOLERANCE_KG = 20.0

# ── enums (string values match scope3Waste.types.ts) ─────────────────────────

DOC_WTN = "waste_transfer_note"
DOC_INVOICE = "waste_contractor_invoice"
DOC_WEIGHBRIDGE = "weighbridge_ticket"
DOC_REPORT = "waste_report"
DOC_HWCN = "hazardous_waste_consignment_note"
DOC_SPREADSHEET = "waste_spreadsheet"

ROLE_ORIGINAL = "original"
ROLE_CREDIT_NOTE = "credit_note"
ROLE_SEASON_TICKET = "season_ticket"

ROUTE_REUSE = "re_use"
ROUTE_OPEN_LOOP = "open_loop"
ROUTE_CLOSED_LOOP = "closed_loop"
ROUTE_COMBUSTION = "combustion"
ROUTE_COMPOSTING = "composting"
ROUTE_AD = "anaerobic_digestion"
ROUTE_LANDFILL = "landfill"

SRC_STATED = "stated"
SRC_RD_CODE = "rd_code"
SRC_LANDFILL_TAX = "landfill_tax_line"
SRC_KEYWORD = "keyword"
SRC_DEFAULT = "default"

UNIT_KG = "kg"
UNIT_T = "t"
UNIT_M3 = "m3"
UNIT_YD3 = "yd3"
UNIT_L = "l"
UNIT_LIFT = "lift"

BASIS_WEIGHED = "weighed"
BASIS_ESTIMATED = "estimated"
BASIS_VOLUME = "volume_converted"
BASIS_LIFT = "lift_estimated"
BASIS_TAX = "back_calculated_from_tax"

CONTAINER_BIN = "bin"
CONTAINER_SKIP = "skip"
CONTAINER_RORO = "roll_on_roll_off"
CONTAINER_COMPACTOR = "compactor"
CONTAINER_DRUM = "drum"
CONTAINER_IBC = "ibc"
CONTAINER_BAG = "bag"
CONTAINER_LOOSE = "loose"
CONTAINER_OTHER = "other"

LINE_WASTE = "waste"
LINE_REBATE = "rebate"
LINE_LANDFILL_TAX = "landfill_tax"
LINE_NON_WASTE = "non_waste_charge"
LINE_TOTAL = "total"

FLAG_LOW_CONFIDENCE = "low_confidence"
FLAG_ESTIMATED = "estimated_quantity"
FLAG_VOLUME = "volume_converted"
FLAG_LIFT = "lift_estimated"
FLAG_AMBIGUOUS_UNIT = "ambiguous_unit"
FLAG_NET_MISMATCH = "net_weight_mismatch"
FLAG_GROSS_ONLY = "gross_weight_only"
FLAG_NEGATIVE = "negative_quantity"
FLAG_HAZ_NON_HWCN = "hazardous_on_non_hwcn"
FLAG_ROUTE_DEFAULTED = "route_defaulted"
FLAG_INTERIM = "interim_treatment"
FLAG_OUTLIER = "quantity_outlier"
FLAG_CONSIGNEE_DIFFERS = "consignee_quantity_differs"
FLAG_MULTI_SITE = "multi_site"
FLAG_POSSIBLE_DUPLICATE = "possible_duplicate"
FLAG_NO_WASTE_LINES = "no_waste_lines"

# ── materials ────────────────────────────────────────────────────────────────

MAT_CI = "commercial_and_industrial_waste"
MAT_FOOD = "organic_food_and_drink_waste"
MAT_GARDEN = "organic_garden_waste"
MAT_PAPER_MIXED = "paper_and_board_mixed"
MAT_BOARD = "paper_and_board_board"
MAT_PLASTICS = "plastics_average_plastics"
MAT_PLASTIC_FILM = "plastics_average_plastic_film"
MAT_GLASS = "glass"
MAT_METAL_CANS = "metal_mixed_cans"
MAT_METAL_SCRAP = "metal_scrap_metal"
MAT_WOOD = "wood"
MAT_CLOTHING = "clothing"
MAT_WEEE = "weee_mixed"
MAT_BATTERIES = "batteries"
MAT_MINERAL_OIL = "mineral_oil"
MAT_TYRES = "tyres"
MAT_CONSTRUCTION = "average_construction"
MAT_AGGREGATES = "aggregates"
MAT_CONCRETE = "concrete"
MAT_SOILS = "soils"
MAT_PLASTERBOARD = "plasterboard"
MAT_INSULATION = "insulation"
MAT_ASBESTOS = "asbestos"

# Fallback route when a document carries no route evidence – always flagged RouteDefaulted.
MATERIAL_DEFAULT_ROUTE: dict[str, str] = {
    MAT_CI: ROUTE_COMBUSTION,
    MAT_FOOD: ROUTE_AD,
    MAT_GARDEN: ROUTE_COMPOSTING,
    MAT_PAPER_MIXED: ROUTE_OPEN_LOOP,
    MAT_BOARD: ROUTE_OPEN_LOOP,
    MAT_PLASTICS: ROUTE_OPEN_LOOP,
    MAT_PLASTIC_FILM: ROUTE_OPEN_LOOP,
    MAT_GLASS: ROUTE_OPEN_LOOP,
    MAT_METAL_CANS: ROUTE_OPEN_LOOP,
    MAT_METAL_SCRAP: ROUTE_OPEN_LOOP,
    MAT_WOOD: ROUTE_OPEN_LOOP,
    MAT_CLOTHING: ROUTE_REUSE,
    MAT_WEEE: ROUTE_OPEN_LOOP,
    MAT_BATTERIES: ROUTE_OPEN_LOOP,
    MAT_MINERAL_OIL: ROUTE_CLOSED_LOOP,
    MAT_TYRES: ROUTE_OPEN_LOOP,
    MAT_CONSTRUCTION: ROUTE_OPEN_LOOP,
    MAT_AGGREGATES: ROUTE_OPEN_LOOP,
    MAT_CONCRETE: ROUTE_OPEN_LOOP,
    MAT_SOILS: ROUTE_OPEN_LOOP,
    MAT_PLASTERBOARD: ROUTE_OPEN_LOOP,
    MAT_INSULATION: ROUTE_LANDFILL,
    MAT_ASBESTOS: ROUTE_LANDFILL,
}

# ── EWC table (subset used by stream presets) ────────────────────────────────
# code → (description, hazardous, DESNZ material or None = user selection, EA density t/m³)

EWC_CODES: dict[str, tuple[str, bool, str | None, float]] = {
    "08 01 11*": ("waste paint and varnish containing organic solvents or other dangerous substances", True, None, 0.567),
    "12 01 01": ("ferrous metal filings and turnings", False, MAT_METAL_SCRAP, 0.3037),
    "13 01 10*": ("mineral based non-chlorinated hydraulic oils", True, MAT_MINERAL_OIL, 0.9),
    "13 02 05*": ("mineral-based non-chlorinated engine, gear and lubricating oils", True, MAT_MINERAL_OIL, 0.9),
    "15 01 01": ("paper and cardboard packaging", False, MAT_BOARD, 0.2),
    "15 01 02": ("plastic packaging", False, MAT_PLASTICS, 0.22),
    "15 01 03": ("wooden packaging", False, MAT_WOOD, 0.11),
    "15 01 04": ("metallic packaging", False, MAT_METAL_CANS, 0.22),
    "15 01 06": ("mixed packaging", False, None, 0.2),
    "15 01 07": ("glass packaging", False, MAT_GLASS, 0.3332),
    "15 02 02*": ("absorbents, filter materials, wiping cloths, protective clothing contaminated by dangerous substances", True, None, 0.418),
    "15 02 03": ("absorbents, filter materials, wiping cloths and protective clothing other than those mentioned in 15 02 02", False, None, 0.07),
    "16 01 03": ("end-of-life tyres", False, MAT_TYRES, 0.4657),
    "16 06 01*": ("lead batteries", True, MAT_BATTERIES, 1.35),
    "17 01 07": ("mixtures of concrete, bricks, tiles and ceramics", False, MAT_AGGREGATES, 0.66),
    "17 02 01": ("wood", False, MAT_WOOD, 0.33),
    "17 04 05": ("iron and steel", False, MAT_METAL_SCRAP, 0.43),
    "17 05 04": ("soil and stones other than those mentioned in 17 05 03", False, MAT_SOILS, 1.06),
    "17 06 04": ("insulation materials other than those mentioned in 17 06 01 and 17 06 03", False, MAT_INSULATION, 0.25),
    "17 06 05*": ("construction materials containing asbestos", True, MAT_ASBESTOS, None),
    "17 08 02": ("gypsum-based construction materials other than those mentioned in 17 08 01", False, MAT_PLASTERBOARD, 0.6092),
    "17 09 04": ("mixed construction and demolition wastes", False, MAT_CONSTRUCTION, 0.32),
    "20 01 01": ("paper and cardboard", False, MAT_PAPER_MIXED, 0.2105),
    "20 01 08": ("biodegradable kitchen and canteen waste", False, MAT_FOOD, 0.2),
    "20 01 10": ("clothes", False, MAT_CLOTHING, 0.2),
    "20 01 21*": ("fluorescent tubes and other mercury-containing waste", True, MAT_WEEE, 0.1886),
    "20 01 36": ("discarded electrical and electronic equipment", False, MAT_WEEE, 0.2131),
    "20 02 01": ("biodegradable waste", False, MAT_GARDEN, 0.38),
    "20 03 01": ("mixed municipal waste", False, MAT_CI, 0.26),
    "20 03 07": ("bulky waste", False, MAT_CI, 0.18),
}

# Material-level fallback densities (t/m³) when the EWC code has no EA factor (EA values of the
# representative code). 17 06 05* has no EA factor – asbestos is only ever generated as weighed.
MATERIAL_DENSITY_FALLBACK: dict[str, float] = {
    MAT_CI: 0.26,
    MAT_PAPER_MIXED: 0.2,
}


def ewc_density(ewc_code: str | None, material: str | None) -> float | None:
    entry = EWC_CODES.get(ewc_code or "")
    if entry and entry[3]:
        return entry[3]
    return MATERIAL_DENSITY_FALLBACK.get(material or "")


def ewc_is_hazardous(ewc_code: str | None) -> bool | None:
    entry = EWC_CODES.get(ewc_code or "")
    return entry[1] if entry else None


# ── R/D codes ────────────────────────────────────────────────────────────────

RECYCLED = "recycled"  # resolved by the loop rule
USER_SELECTION = "user_selection"
INTERIM = "interim"

RD_CODE_ROUTE: dict[str, str] = {
    "R1": ROUTE_COMBUSTION,
    "R2": RECYCLED,
    "R3": "by_material",
    "R4": RECYCLED,
    "R5": RECYCLED,
    "R6": RECYCLED,
    "R7": RECYCLED,
    "R8": RECYCLED,
    "R9": ROUTE_CLOSED_LOOP,
    "R10": USER_SELECTION,
    "R11": USER_SELECTION,
    "R12": INTERIM,
    "R13": INTERIM,
    "D1": ROUTE_LANDFILL,
    "D5": ROUTE_LANDFILL,
    "D8": USER_SELECTION,
    "D9": USER_SELECTION,
    "D10": ROUTE_COMBUSTION,
    "D13": INTERIM,
    "D14": INTERIM,
    "D15": INTERIM,
}

RD_CODE_TEXT: dict[str, str] = {
    "R1": "use principally as a fuel to generate energy",
    "R3": "recycling / reclamation of organic substances",
    "R4": "recycling / reclamation of metals",
    "R5": "recycling / reclamation of other inorganic materials",
    "R9": "oil re-refining or other reuses of oil",
    "R13": "storage pending any of the operations R1-R12",
    "D1": "deposit into or onto land (landfill)",
    "D5": "specially engineered landfill",
    "D10": "incineration on land",
    "D15": "storage pending any of the operations D1-D14",
}

_ORGANIC_MATERIALS = {MAT_FOOD, MAT_GARDEN}


def rd_code_route(rd_code: str | None, material: str | None, loop_route: str) -> str | None:
    """Route implied by an R/D code; None = interim / user selection (flag)."""
    mapped = RD_CODE_ROUTE.get((rd_code or "").upper())
    if mapped in (None, USER_SELECTION, INTERIM):
        return None
    if mapped == RECYCLED:
        return loop_route
    if mapped == "by_material":
        # R3 organics = composting OR anaerobic digestion → needs other evidence.
        return None if material in _ORGANIC_MATERIALS else loop_route
    return mapped


def rd_code_is_interim(rd_code: str | None) -> bool:
    return RD_CODE_ROUTE.get((rd_code or "").upper()) == INTERIM


# ── route words (stated) and stream keywords ─────────────────────────────────

ROUTE_LABEL: dict[str, str] = {
    ROUTE_COMBUSTION: "Energy from Waste",
    ROUTE_LANDFILL: "Landfill",
    ROUTE_AD: "Anaerobic Digestion",
    ROUTE_COMPOSTING: "Composting",
    ROUTE_OPEN_LOOP: "Recycling",
    ROUTE_CLOSED_LOOP: "Closed-loop recycling",
    ROUTE_REUSE: "Re-use",
}

# ── landfill tax (GBP per tonne) ─────────────────────────────────────────────
# (effective_from, standard_rate, lower_rate). England & NI (HMRC); Scotland (SLfT) and
# Wales (LDT) set identical rates for 2025/26 and 2026/27.
LANDFILL_TAX_RATES: list[tuple[date, float, float]] = [
    (date(2023, 4, 1), 102.10, 3.25),
    (date(2024, 4, 1), 103.70, 3.30),
    (date(2025, 4, 1), 126.15, 4.05),
    (date(2026, 4, 1), 130.75, 8.65),
]


def landfill_tax_rate(on: date, *, lower: bool = False) -> float:
    rate = LANDFILL_TAX_RATES[0]
    for entry in LANDFILL_TAX_RATES:
        if on >= entry[0]:
            rate = entry
    return rate[2] if lower else rate[1]


# ── non-waste charges (rendered as distractor lines) ─────────────────────────

NON_WASTE_CHARGES: list[tuple[str, float]] = [
    ("Bin rental 1100L x2", 12.00),
    ("Duty of Care documentation fee (annual)", 45.00),
    ("Skip permit (highway) – local council", 45.00),
    ("Wait & load (30 min)", 35.00),
    ("Container exchange / haulage", 165.00),
    ("Bin cleaning", 15.00),
    ("Contamination charge", 60.00),
    ("Compactor service visit", 140.00),
    ("Replacement bin lids 1100L", 36.00),
    ("Environmental surcharge", 18.50),
]

FUEL_SURCHARGE_PCT = 4

# ── stream presets ───────────────────────────────────────────────────────────
# Each preset is a recurring waste stream at a site.
#   container: (type, size, unit)
#   route: the true treatment route at the receiving facility
#   rd_code: code the receiving facility records
#   keyword_route: route implied by the stream name alone (findings: "stream keyword"), or None
#   kg_per_lift: typical weighed range per lift/collection

STREAM_PRESETS: dict[str, dict] = {
    "general_waste_1100": {
        "label": "General waste – 1100L bin",
        "description": "General Waste 1100L",
        "ewc": "20 03 01",
        "container": (CONTAINER_BIN, 1100, UNIT_L),
        "route": ROUTE_COMBUSTION,
        "rd_code": "R1",
        "keyword_route": None,
        "kg_per_lift": (45, 75),
    },
    "dmr_1100": {
        "label": "Dry mixed recycling – 1100L bin",
        "description": "Dry Mixed Recycling 1100L",
        "ewc": "15 01 06",
        "material": MAT_PAPER_MIXED,
        "container": (CONTAINER_BIN, 1100, UNIT_L),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R3",
        "keyword_route": ROUTE_OPEN_LOOP,
        "kg_per_lift": (22, 40),
    },
    "food_240": {
        "label": "Food waste – 240L bin",
        "description": "Food Waste 240L",
        "ewc": "20 01 08",
        "container": (CONTAINER_BIN, 240, UNIT_L),
        "route": ROUTE_AD,
        "rd_code": "R3",
        "keyword_route": ROUTE_AD,
        "kg_per_lift": (35, 70),
    },
    "general_waste_360": {
        "label": "General waste – 360L wheelie bin",
        "description": "General waste – 360L wheelie bin",
        "ewc": "20 03 01",
        "container": (CONTAINER_BIN, 360, UNIT_L),
        "route": ROUTE_COMBUSTION,
        "rd_code": "R1",
        "keyword_route": None,
        "kg_per_lift": (15, 30),
    },
    "cardboard_baled": {
        "label": "Cardboard (baled) – rebate",
        "description": "Baled cardboard (OCC)",
        "ewc": "15 01 01",
        "container": (CONTAINER_LOOSE, None, None),
        "route": ROUTE_CLOSED_LOOP,
        "rd_code": "R3",
        "keyword_route": ROUTE_OPEN_LOOP,
        "kg_per_lift": (1800, 3600),
        "rebate": True,
    },
    "plastic_film": {
        "label": "Plastic film (bales)",
        "description": "Mixed plastic film (bales)",
        "ewc": "15 01 02",
        "material": MAT_PLASTIC_FILM,
        "container": (CONTAINER_LOOSE, None, None),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R3",
        "keyword_route": None,
        "kg_per_lift": (300, 700),
    },
    "compactor_general": {
        "label": "General waste – 35yd RORO compactor",
        "description": "35yd RORO compactor – general waste",
        "ewc": "20 03 01",
        "container": (CONTAINER_COMPACTOR, 35, UNIT_YD3),
        "route": ROUTE_LANDFILL,
        "rd_code": "D1",
        "keyword_route": None,
        "kg_per_lift": (3800, 6200),
    },
    "cd_skip_8": {
        "label": "Mixed C&D – 8yd skip",
        "description": "8yd skip – mixed C&D waste",
        "ewc": "17 09 04",
        "container": (CONTAINER_SKIP, 8, UNIT_YD3),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R5",
        "keyword_route": None,
        "kg_per_lift": (1400, 2600),
    },
    "plasterboard_skip_4": {
        "label": "Plasterboard – 4yd skip",
        "description": "4yd skip – plasterboard (segregated)",
        "ewc": "17 08 02",
        "container": (CONTAINER_SKIP, 4, UNIT_YD3),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R5",
        "keyword_route": None,
        "kg_per_lift": (900, 1700),
    },
    "wood_skip_12": {
        "label": "Wood – 12yd skip",
        "description": "12yd skip – wood (grade A/B)",
        "ewc": "17 02 01",
        "container": (CONTAINER_SKIP, 12, UNIT_YD3),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R3",
        "keyword_route": None,
        "kg_per_lift": (1200, 2400),
    },
    "metal_skip_8": {
        "label": "Scrap metal – 8yd skip",
        "description": "8yd skip – mixed steel scrap",
        "ewc": "17 04 05",
        "container": (CONTAINER_SKIP, 8, UNIT_YD3),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R4",
        "keyword_route": None,
        "kg_per_lift": (1500, 3400),
    },
    "swarf_skip": {
        "label": "Steel swarf – 6yd skip",
        "description": "Ferrous metal filings and turnings (steel swarf)",
        "ewc": "12 01 01",
        "container": (CONTAINER_SKIP, 6, UNIT_YD3),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R4",
        "keyword_route": None,
        "kg_per_lift": (1200, 2600),
    },
    "weee_cage": {
        "label": "WEEE – cage",
        "description": "Mixed WEEE (IT equipment) – cage",
        "ewc": "20 01 36",
        "container": (CONTAINER_OTHER, None, None),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R4",
        "keyword_route": None,
        "kg_per_lift": (80, 260),
    },
    "waste_oil_ibc": {
        "label": "Waste hydraulic oil – IBC (hazardous)",
        "description": "Waste mineral hydraulic oil",
        "ewc": "13 01 10*",
        "container": (CONTAINER_IBC, 1000, UNIT_L),
        "route": ROUTE_CLOSED_LOOP,
        "rd_code": "R9",
        "keyword_route": None,
        "kg_per_lift": (150, 450),
        "physical_form": "Liquid",
        "hazard_codes": "HP14",
        "components": "Mineral oil 95%, water 5%",
        "un_number": "UN3082",
    },
    "fluorescent_tubes": {
        "label": "Fluorescent tubes (hazardous)",
        "description": "Fluorescent tubes",
        "ewc": "20 01 21*",
        "container": (CONTAINER_OTHER, None, None),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R13",
        "keyword_route": None,
        "kg_per_lift": (20, 60),
        "physical_form": "Solid",
        "hazard_codes": "HP6, HP14",
        "container_text": "Coffin box",
        "components": "Glass 88%, metals 10%, phosphor powder 2%, mercury <0.01%",
        "un_number": "UN3506",
    },
    "oily_rags_drum": {
        "label": "Oily rags & absorbents – drum (hazardous)",
        "description": "Oily rags and absorbents",
        "ewc": "15 02 02*",
        "material": MAT_CI,
        "container": (CONTAINER_DRUM, 205, UNIT_L),
        "route": ROUTE_COMBUSTION,
        "rd_code": "R1",
        "keyword_route": None,
        "kg_per_lift": (60, 140),
        "physical_form": "Solid",
        "hazard_codes": "HP14",
        "components": "Textiles 70%, mineral oil 25%, water 5%",
        "un_number": "UN3175",
    },
    "paint_solvent_drum": {
        "label": "Waste paint & solvent – drum (hazardous)",
        "description": "Waste paint and solvent",
        "ewc": "08 01 11*",
        "material": MAT_CI,
        "container": (CONTAINER_DRUM, 25, UNIT_L),
        "route": ROUTE_COMBUSTION,
        "rd_code": "D10",
        "keyword_route": None,
        "kg_per_lift": (40, 90),
        "physical_form": "Liquid",
        "hazard_codes": "HP3, HP14",
        "components": "Paint solids 55%, xylene 30%, other solvents 15%",
        "un_number": "UN1263",
    },
    "lead_batteries": {
        "label": "Lead-acid batteries (hazardous)",
        "description": "Lead-acid batteries (forklift / UPS)",
        "ewc": "16 06 01*",
        "container": (CONTAINER_OTHER, None, None),
        "route": ROUTE_OPEN_LOOP,
        "rd_code": "R4",
        "keyword_route": None,
        "kg_per_lift": (120, 480),
        "physical_form": "Solid",
        "hazard_codes": "HP8, HP10",
        "container_text": "Battery box",
        "components": "Lead 65%, sulphuric acid 20%, polypropylene 15%",
        "un_number": "UN2794",
    },
    "asbestos": {
        "label": "Asbestos cement sheets (hazardous)",
        "description": "Asbestos cement sheets (roof repair)",
        "ewc": "17 06 05*",
        "container": (CONTAINER_BAG, None, None),
        "route": ROUTE_LANDFILL,
        "rd_code": "D5",
        "keyword_route": None,
        "kg_per_lift": (200, 600),
        "physical_form": "Solid",
        "hazard_codes": "HP7",
        "container_text": "Sealed bags",
        "components": "Cement 85-90%, chrysotile asbestos 10-15%",
        "un_number": "UN2590",
    },
}

DEFAULT_SITE_STREAMS: list[list[str]] = [
    ["general_waste_1100", "dmr_1100", "food_240", "cardboard_baled", "compactor_general", "waste_oil_ibc", "fluorescent_tubes"],
    ["cd_skip_8", "plasterboard_skip_4", "wood_skip_12", "metal_skip_8", "oily_rags_drum", "asbestos"],
    ["general_waste_360", "dmr_1100", "food_240", "weee_cage", "paint_solvent_drum", "lead_batteries"],
]

HAZARDOUS_PRESET_KEYS = [key for key, preset in STREAM_PRESETS.items() if preset["ewc"].endswith("*")]


def preset_material(preset: dict) -> str:
    if preset.get("material"):
        return preset["material"]
    entry = EWC_CODES.get(preset["ewc"])
    return (entry[2] if entry else None) or MAT_CI


def preset_loop_route(preset: dict) -> str:
    return ROUTE_CLOSED_LOOP if preset["route"] == ROUTE_CLOSED_LOOP else ROUTE_OPEN_LOOP


# ── fictional contractors and receiving facilities ───────────────────────────

WASTE_CONTRACTORS: list[dict] = [
    {
        "name": "Greenway Waste Services Ltd",
        "code": "GW",
        "carrier_reg": "CBDU204417",
        "address": ["Greenway House", "2 Canal Wharf", "Leeds LS11 5QR", "United Kingdom"],
    },
    {
        "name": "Northfield Skip Hire Ltd",
        "code": "NSH",
        "carrier_reg": "CBDU318820",
        "address": ["Yard 3, Calder Road", "Dewsbury WF13 2HT", "United Kingdom"],
    },
    {
        "name": "Meridian Environmental Ltd",
        "code": "MEL",
        "carrier_reg": "CBDU771034",
        "address": ["Treatment Centre, Dock Lane", "Hull HU9 1XX", "United Kingdom"],
    },
]

RECEIVING_FACILITIES: dict[str, dict] = {
    ROUTE_COMBUSTION: {"name": "Aire Valley Energy Recovery Facility", "permit": "EPR/AV3301ER", "address": "Skelton Grange Road, Leeds LS10 1RS", "type_text": "EfW"},
    ROUTE_LANDFILL: {"name": "Pennine Landfill & Recovery Ltd", "permit": "EPR/PB3381KL", "address": "Holmfield Quarry, Halifax HX3 6TT", "type_text": "LANDFILL CELL 4"},
    ROUTE_AD: {"name": "Calder Biogas AD Plant", "permit": "EPR/CB5521AD", "address": "Wharf Road, Mirfield WF14 8LP", "type_text": "AD plant"},
    ROUTE_COMPOSTING: {"name": "Moorside Composting Ltd", "permit": "EPR/MC7740CP", "address": "Moorside Farm, Otley LS21 2PL", "type_text": "IVC composting"},
    ROUTE_OPEN_LOOP: {"name": "Northfield Materials Recovery Facility", "permit": "EPR/NB7718MR", "address": "Calder Road, Dewsbury WF13 2HT", "type_text": "MRF – recycling"},
    ROUTE_CLOSED_LOOP: {"name": "Apex Fibre Recycling Ltd", "permit": "EPR/AF5521PM", "address": "Mill Street, Bradford BD3 9LP", "type_text": "reprocessor – closed-loop recycling"},
    ROUTE_REUSE: {"name": "Second Life Re-use Centre", "permit": "EPR/SL1101RU", "address": "Unit 2, Hunslet Trading Estate, Leeds LS10 1SB", "type_text": "re-use"},
}

INTERIM_FACILITY = {"name": "Meridian Hazardous Waste Transfer Station", "permit": "EPR/MB4417PX", "address": "Dock Lane, Hull HU9 1XX", "type_text": "transfer station"}

SIC_CODES = ["25620", "28290", "41201", "46900", "52103", "64191"]
COUNCILS = ["Leeds City Council", "Wakefield Council", "City of York Council", "Kirklees Council", "Bradford Council"]
