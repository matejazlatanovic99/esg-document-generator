from __future__ import annotations

import random
from dataclasses import dataclass


@dataclass(frozen=True)
class DocumentDistractorField:
    field_id: str
    label: str
    value: str
    placement: str


@dataclass(frozen=True)
class TabularDistractorField:
    field_id: str
    label: str
    anchor: str
    position: str
    row_scope: str
    value_options: tuple[str, ...]


@dataclass(frozen=True)
class DistractorPlan:
    enabled: bool
    seed: str
    document_fields: tuple[DocumentDistractorField, ...] = ()
    tabular_fields: tuple[TabularDistractorField, ...] = ()


@dataclass(frozen=True)
class _Template:
    field_id: str
    label_key: str
    placement_or_anchor: str
    values: tuple[str, ...]
    position: str = "after"
    row_scope: str = "file"


_LABELS: dict[str, dict[str, str]] = {
    "en": {
        "connection_id": "Connection ID",
        "service_route": "Service Route",
        "network_zone": "Network Zone",
        "settlement_run": "Settlement Run",
        "feeder_code": "Feeder Code",
        "profile_class": "Profile Class",
        "meter_point_group": "Meter Point Group",
        "dispatch_ref": "Dispatch Ref",
        "haulier_code": "Haulier Code",
        "fleet_group": "Fleet Group",
        "cost_centre": "Cost Centre",
        "shift_code": "Shift Code",
        "plant_group": "Plant Group",
        "source_system": "Source System",
        "report_batch": "Report Batch",
    },
    "fr": {
        "connection_id": "ID de raccordement",
        "service_route": "Tournée de service",
        "network_zone": "Zone réseau",
        "settlement_run": "Cycle de règlement",
        "feeder_code": "Code départ",
        "profile_class": "Classe de profil",
        "meter_point_group": "Groupe point de mesure",
        "dispatch_ref": "Réf. expédition",
        "haulier_code": "Code transporteur",
        "fleet_group": "Groupe flotte",
        "cost_centre": "Centre de coût",
        "shift_code": "Code équipe",
        "plant_group": "Groupe usine",
        "source_system": "Système source",
        "report_batch": "Lot rapport",
    },
    "de": {
        "connection_id": "Anschluss-ID",
        "service_route": "Serviceroute",
        "network_zone": "Netzzone",
        "settlement_run": "Abrechnungslauf",
        "feeder_code": "Speisecode",
        "profile_class": "Profilklasse",
        "meter_point_group": "Messpunktgruppe",
        "dispatch_ref": "Dispo-Ref.",
        "haulier_code": "Spediteurcode",
        "fleet_group": "Flottengruppe",
        "cost_centre": "Kostenstelle",
        "shift_code": "Schichtcode",
        "plant_group": "Anlagengruppe",
        "source_system": "Quellsystem",
        "report_batch": "Berichtslauf",
    },
    "nl": {
        "connection_id": "Aansluit-ID",
        "service_route": "Serviceroute",
        "network_zone": "Netwerkzone",
        "settlement_run": "Verrekeningsrun",
        "feeder_code": "Voedercode",
        "profile_class": "Profielklasse",
        "meter_point_group": "Meetpuntgroep",
        "dispatch_ref": "Dispatch-ref.",
        "haulier_code": "Transportcode",
        "fleet_group": "Vlootgroep",
        "cost_centre": "Kostenplaats",
        "shift_code": "Ploegcode",
        "plant_group": "Installatiegroep",
        "source_system": "Bronsysteem",
        "report_batch": "Rapportbatch",
    },
}


_HEAT_DOCUMENT_FIELDS: tuple[_Template, ...] = (
    _Template("connection_id", "connection_id", "meta", ("HX-204", "HX-317", "HX-442")),
    _Template("service_route", "service_route", "meta", ("R-17B", "R-24A", "R-31C")),
    _Template("network_zone", "network_zone", "summary", ("NORTH-2", "CEN-4", "SOUTH-1")),
)

_HEAT_TABULAR_FIELDS: tuple[_Template, ...] = (
    _Template("connection_id", "connection_id", "meter_id", ("HX-204", "HX-317", "HX-442"), position="before", row_scope="statement"),
    _Template("service_route", "service_route", "meter_id", ("R-17B", "R-24A", "R-31C"), position="after", row_scope="statement"),
    _Template("network_zone", "network_zone", "period_start", ("NORTH-2", "CEN-4", "SOUTH-1"), position="before", row_scope="statement"),
)

_ELECTRICITY_DOCUMENT_FIELDS: tuple[_Template, ...] = (
    _Template("settlement_run", "settlement_run", "period_meta", ("SF-01", "SF-03", "RF-08")),
    _Template("feeder_code", "feeder_code", "meter_table", ("FD-118", "FD-224", "FD-405")),
    _Template("profile_class", "profile_class", "grid_table", ("PC-1", "PC-4", "PC-7")),
    _Template("meter_point_group", "meter_point_group", "period_meta", ("MPG-A", "MPG-C", "MPG-H")),
)

_ELECTRICITY_TABULAR_FIELDS: tuple[_Template, ...] = (
    _Template("settlement_run", "settlement_run", "reference", ("SF-01", "SF-03", "RF-08"), position="after", row_scope="statement"),
    _Template("feeder_code", "feeder_code", "meter_id", ("FD-118", "FD-224", "FD-405"), position="after", row_scope="statement"),
    _Template("profile_class", "profile_class", "total_qty", ("PC-1", "PC-4", "PC-7"), position="before", row_scope="row"),
    _Template("meter_point_group", "meter_point_group", "meter_id", ("MPG-A", "MPG-C", "MPG-H"), position="before", row_scope="statement"),
)

_SMART_METER_TABULAR_FIELDS: tuple[_Template, ...] = (
    _Template("settlement_run", "settlement_run", "meter_id", ("SF-01", "SF-03", "RF-08"), position="after", row_scope="statement"),
    _Template("feeder_code", "feeder_code", "meter_id", ("FD-118", "FD-224", "FD-405"), position="after", row_scope="statement"),
    _Template("profile_class", "profile_class", "meter_id", ("PC-1", "PC-4", "PC-7"), position="after", row_scope="row"),
)

_LOGISTICS_DOCUMENT_FIELDS: tuple[_Template, ...] = (
    _Template("dispatch_ref", "dispatch_ref", "meta", ("DSP-208", "DSP-331", "DSP-517")),
    _Template("haulier_code", "haulier_code", "summary", ("HL-17", "HL-24", "HL-66")),
    _Template("fleet_group", "fleet_group", "summary", ("FG-2", "FG-5", "FG-9")),
    _Template("cost_centre", "cost_centre", "meta", ("CC-104", "CC-228", "CC-642")),
)

_LOGISTICS_TABULAR_FIELDS: tuple[_Template, ...] = (
    _Template("dispatch_ref", "dispatch_ref", "site", ("DSP-208", "DSP-331", "DSP-517"), position="after", row_scope="statement"),
    _Template("haulier_code", "haulier_code", "site", ("HL-17", "HL-24", "HL-66"), position="before", row_scope="statement"),
    _Template("fleet_group", "fleet_group", "quantity", ("FG-2", "FG-5", "FG-9"), position="before", row_scope="statement"),
    _Template("cost_centre", "cost_centre", "quantity", ("CC-104", "CC-228", "CC-642"), position="before", row_scope="statement"),
)

_ADMIN_DOCUMENT_FIELDS: tuple[_Template, ...] = (
    _Template("shift_code", "shift_code", "summary", ("S-A", "S-B", "S-C")),
    _Template("plant_group", "plant_group", "summary", ("PL-01", "PL-03", "PL-07")),
    _Template("source_system", "source_system", "meta", ("BEMS-HUB", "OPS-RT", "GEN-LOG")),
    _Template("report_batch", "report_batch", "meta", ("RB-101", "RB-204", "RB-305")),
)

_ADMIN_TABULAR_FIELDS: tuple[_Template, ...] = (
    _Template("shift_code", "shift_code", "period", ("S-A", "S-B", "S-C"), position="before", row_scope="row"),
    _Template("plant_group", "plant_group", "site", ("PL-01", "PL-03", "PL-07"), position="after", row_scope="block"),
    _Template("source_system", "source_system", "site", ("BEMS-HUB", "OPS-RT", "GEN-LOG"), position="after", row_scope="block"),
    _Template("report_batch", "report_batch", "timestamp", ("RB-101", "RB-204", "RB-305"), position="before", row_scope="row"),
)


_CATALOG: dict[tuple[str, str, str], dict[str, tuple[_Template, ...]]] = {
    ("heat", "utility_bill", "PDF"): {"document": _HEAT_DOCUMENT_FIELDS},
    ("heat", "utility_bill", "DOCX"): {"document": _HEAT_DOCUMENT_FIELDS},
    ("heat", "supplier_portal_data", "CSV"): {"tabular": _HEAT_TABULAR_FIELDS},
    ("heat", "supplier_portal_data", "XLSX"): {"tabular": _HEAT_TABULAR_FIELDS},
    ("electricity", "electricity_bill", "PDF"): {"document": _ELECTRICITY_DOCUMENT_FIELDS},
    ("electricity", "electricity_bill", "DOCX"): {"document": _ELECTRICITY_DOCUMENT_FIELDS},
    ("electricity", "supplier_portal_data", "CSV"): {"tabular": _ELECTRICITY_TABULAR_FIELDS},
    ("electricity", "supplier_portal_data", "XLSX"): {"tabular": _ELECTRICITY_TABULAR_FIELDS},
    ("electricity", "smart_meter_data", "CSV"): {"tabular": _SMART_METER_TABULAR_FIELDS},
    ("electricity", "smart_meter_data", "XLSX"): {"tabular": _SMART_METER_TABULAR_FIELDS},
    ("stationary_combustion", "fuel_invoice", "PDF"): {"document": _LOGISTICS_DOCUMENT_FIELDS},
    ("stationary_combustion", "fuel_invoice", "DOCX"): {"document": _LOGISTICS_DOCUMENT_FIELDS},
    ("stationary_combustion", "delivery_note", "PDF"): {"document": _LOGISTICS_DOCUMENT_FIELDS},
    ("stationary_combustion", "delivery_note", "DOCX"): {"document": _LOGISTICS_DOCUMENT_FIELDS},
    ("stationary_combustion", "fuel_card", "PDF"): {"document": _LOGISTICS_DOCUMENT_FIELDS},
    ("stationary_combustion", "fuel_card", "DOCX"): {"document": _LOGISTICS_DOCUMENT_FIELDS},
    ("stationary_combustion", "fuel_card", "CSV"): {"tabular": _LOGISTICS_TABULAR_FIELDS},
    ("stationary_combustion", "fuel_card", "XLSX"): {"tabular": _LOGISTICS_TABULAR_FIELDS},
    ("stationary_combustion", "generator_log", "CSV"): {"tabular": _ADMIN_TABULAR_FIELDS},
    ("stationary_combustion", "generator_log", "XLSX"): {"tabular": _ADMIN_TABULAR_FIELDS},
    ("stationary_combustion", "bems", "PDF"): {"document": _ADMIN_DOCUMENT_FIELDS},
    ("stationary_combustion", "bems", "DOCX"): {"document": _ADMIN_DOCUMENT_FIELDS},
    ("stationary_combustion", "bems", "CSV"): {"tabular": _ADMIN_TABULAR_FIELDS},
    ("stationary_combustion", "bems", "XLSX"): {"tabular": _ADMIN_TABULAR_FIELDS},
}


def normalize_distractor_settings(document_cfg: dict) -> dict:
    distractor_cfg = document_cfg.get("distractor_fields", {})
    return {
        "enabled": bool(distractor_cfg.get("enabled", False)),
    }


def resolve_distractor_plan(
    document_cfg: dict,
    *,
    random_seed: int,
    category: str,
    document_type: str,
    output_format: str,
    artifact_key: str = "root",
    context: dict | None = None,
) -> DistractorPlan:
    enabled = normalize_distractor_settings(document_cfg).get("enabled", False)
    seed = f"{int(random_seed)}:{output_format}:{artifact_key}"
    if not enabled:
        return DistractorPlan(enabled=False, seed=seed)

    catalog = _CATALOG.get((category, document_type, output_format), {})
    language = str((context or {}).get("language") or document_cfg.get("language") or "en")
    document_fields = tuple(
        _resolve_document_field(language, seed, template, index)
        for index, template in enumerate(_pick_templates(catalog.get("document", ()), 2, seed, "document"))
    )
    tabular_count = _tabular_target_count(category, document_type, output_format, context or {})
    tabular_fields = tuple(
        _resolve_tabular_field(language, template)
        for template in _pick_templates(catalog.get("tabular", ()), tabular_count, seed, "tabular")
    )
    return DistractorPlan(
        enabled=True,
        seed=seed,
        document_fields=document_fields,
        tabular_fields=tabular_fields,
    )


def resolve_tabular_value(plan: DistractorPlan, field: TabularDistractorField, subkey: str) -> str:
    if not field.value_options:
        return ""
    rng = random.Random(f"{plan.seed}:{field.field_id}:{field.row_scope}:{subkey}")
    return field.value_options[rng.randrange(len(field.value_options))]


def _pick_templates(
    templates: tuple[_Template, ...],
    count: int,
    seed: str,
    scope: str,
) -> tuple[_Template, ...]:
    if not templates or count <= 0:
        return ()
    if len(templates) <= count:
        return templates
    rng = random.Random(f"{seed}:{scope}")
    return tuple(rng.sample(list(templates), count))


def _tabular_target_count(
    category: str,
    document_type: str,
    output_format: str,
    context: dict,
) -> int:
    del output_format
    if category == "electricity" and document_type == "smart_meter_data":
        return 2
    if category == "stationary_combustion" and document_type == "bems":
        if context.get("bems_report_type") == "time_series_trend_export":
            return 2
    return 3


def _resolve_document_field(
    language: str,
    seed: str,
    template: _Template,
    index: int,
) -> DocumentDistractorField:
    rng = random.Random(f"{seed}:document:{template.field_id}:{index}")
    return DocumentDistractorField(
        field_id=template.field_id,
        label=_label(language, template.label_key),
        value=template.values[rng.randrange(len(template.values))],
        placement=template.placement_or_anchor,
    )


def _resolve_tabular_field(language: str, template: _Template) -> TabularDistractorField:
    return TabularDistractorField(
        field_id=template.field_id,
        label=_label(language, template.label_key),
        anchor=template.placement_or_anchor,
        position=template.position,
        row_scope=template.row_scope,
        value_options=template.values,
    )


def _label(language: str, label_key: str) -> str:
    labels = _LABELS.get(language, _LABELS["en"])
    return labels.get(label_key, _LABELS["en"].get(label_key, label_key))