from __future__ import annotations

import hashlib
import json
import random
from typing import Any


_VALID_PRESETS = {"realistic", "balanced", "stress"}
_ALIAS_RATES = {
    "realistic": 0.25,
    "balanced": 0.45,
    "stress": 0.65,
}

_PREAMBLE_TEMPLATES: dict[str, list[list[str]]] = {
    "none": [],
    "title_only": [["{title}"]],
    "title_period": [["{title}"], ["Period", "{period_label}"]],
    "title_subject": [["{title}"], ["{subject}"]],
    "audit_banner": [["Document export"], ["Preset", "{preset}"], ["Period", "{period_label}"]],
}

_FIELD_ALIAS_POOLS: dict[str, dict[str, list[str]]] = {
    "invoice_no": {"en": ["Invoice No", "Invoice Number", "Billing Ref"]},
    "reference": {"en": ["Reference", "Statement Ref", "Record Ref"]},
    "company": {"en": ["Company", "Entity", "Customer Entity"]},
    "site": {"en": ["Site", "Facility", "Location"]},
    "city": {"en": ["City", "Town", "Municipality"]},
    "postcode": {"en": ["Postcode", "Postal Code", "ZIP / Postcode"]},
    "meter_id": {"en": ["Meter ID", "Meter Reference", "Meter No"]},
    "period": {"en": ["Billing Period", "Charge Period", "Statement Period"]},
    "period_start": {"en": ["Period Start", "Start Date", "From"]},
    "period_end": {"en": ["Period End", "End Date", "To"]},
    "issue_date": {"en": ["Issue Date", "Issued On", "Document Date"]},
    "due_date": {"en": ["Due Date", "Pay By", "Payment Due"]},
    "currency": {"en": ["Currency", "FX", "Billing Currency"]},
    "prev_read": {"en": ["Prev Reading", "Previous Reading", "Opening Reading"]},
    "curr_read": {"en": ["Curr Reading", "Current Reading", "Closing Reading"]},
    "consumption": {"en": ["Consumption", "Usage", "Measured Use"]},
    "unit_price": {"en": ["Unit Price", "Rate", "Unit Rate"]},
    "heat_cost": {"en": ["Heat Cost", "Consumption Cost", "Energy Charge"]},
    "capacity": {"en": ["Capacity", "Contracted Capacity", "Capacity kW"]},
    "cap_rate": {"en": ["Cap. Rate", "Capacity Rate", "Standing Rate"]},
    "supplier_ef": {"en": ["Supplier EF", "Emission Factor", "Grid EF"]},
    "cap_charge": {"en": ["Capacity Charge", "Standing Charge", "Capacity Cost"]},
    "subtotal": {"en": ["Subtotal", "Net Total", "Amount Before VAT"]},
    "vat": {"en": ["VAT", "Tax", "VAT Amount"]},
    "total": {"en": ["Total", "Amount Due", "Total Due"]},
    "unit": {"en": ["Unit", "Measurement Unit", "UoM"]},
    "start_read": {"en": ["Start Reading", "Opening Reading", "Read Start"]},
    "end_read": {"en": ["End Reading", "Closing Reading", "Read End"]},
    "total_qty": {"en": ["Total Quantity", "Total Consumption", "Usage Total"]},
    "total_cost": {"en": ["Total Cost", "Amount", "Charge Total"]},
    "emissions_kg": {"en": ["Emissions (kg CO2e)", "kg CO2e", "Emissions kg"]},
    "emissions_t": {"en": ["Emissions (tCO2e)", "tCO2e", "Emissions tonnes"]},
    "tariff_name": {"en": ["Tariff Name", "Tariff", "Rate Name"]},
    "tariff_qty": {"en": ["Quantity", "Tariff Qty", "Billed Qty"]},
    "tariff_rate": {"en": ["Unit Cost", "Tariff Rate", "Rate"]},
    "tariff_cost": {"en": ["Cost", "Tariff Cost", "Charge"]},
    "timestamp": {"en": ["Timestamp", "Date Time", "Reading Time"]},
    "import_kwh": {"en": ["Import kWh", "Imported kWh", "Grid Import"]},
    "export_kwh": {"en": ["Export kWh", "Exported kWh", "Grid Export"]},
    "tariff_type": {"en": ["Tariff Type", "Rate Type", "Band"]},
    "card_no": {"en": ["Card No", "Card Number", "Fuel Card"]},
    "date": {"en": ["Date", "Transaction Date", "Record Date"]},
    "merchant": {"en": ["Merchant", "Vendor", "Supplier"]},
    "country": {"en": ["Country", "Jurisdiction", "Nation"]},
    "equipment": {"en": ["Equipment", "Asset", "Equipment Name"]},
    "emission_source": {"en": ["Emission Source", "Source", "Emission Category"]},
    "product": {"en": ["Product", "Fuel Product", "Item"]},
    "qty": {"en": ["Qty", "Quantity", "Volume"]},
    "start_time": {"en": ["Start Time", "Run Start", "From Time"]},
    "end_time": {"en": ["End Time", "Run End", "To Time"]},
    "run_hours": {"en": ["Run Hours", "Operating Hours", "Hours"]},
    "start_fuel": {"en": ["Start Fuel", "Opening Fuel", "Fuel Start"]},
    "end_fuel": {"en": ["End Fuel", "Closing Fuel", "Fuel End"]},
    "fuel_used": {"en": ["Fuel Used", "Consumption", "Fuel Consumption"]},
    "fuel_type": {"en": ["Fuel Type", "Fuel", "Product Type"]},
    "notes": {"en": ["Notes", "Comments", "Remarks"]},
    "equipment_tag": {"en": ["Equipment Tag", "Asset Tag", "Tag"]},
    "equipment_name": {"en": ["Equipment Name", "Equipment", "Asset Name"]},
    "operating_hours": {"en": ["Operating Hours", "Run Hours", "Hours"]},
    "sensor_name": {"en": ["Sensor Name", "Sensor", "Channel"]},
    "value": {"en": ["Value", "Reading", "Measured Value"]},
    "time": {"en": ["Time", "Transaction Time", "Time of Day"]},
    "vehicle_reg": {"en": ["Vehicle Registration", "Registration", "Vehicle Reg", "Plate No"]},
    "vehicle_name": {"en": ["Vehicle", "Vehicle Description", "Make / Model"]},
    "vehicle_type": {"en": ["Vehicle Type", "Class", "Vehicle Class"]},
    "driver": {"en": ["Driver", "Driver Name", "Operator"]},
    "receipt_no": {"en": ["Receipt No", "Transaction No", "Receipt Ref"]},
    "odometer": {"en": ["Odometer", "Odometer Reading", "Mileage"]},
    "odometer_start": {"en": ["Odometer Start", "Opening Odometer", "Start Mileage"]},
    "odometer_end": {"en": ["Odometer End", "Closing Odometer", "End Mileage"]},
    "distance": {"en": ["Distance", "Distance Travelled", "Total Distance"]},
    "distance_unit": {"en": ["Distance Unit", "Unit", "UoM"]},
    "trip_start": {"en": ["Trip Start", "Start Time", "Departure"]},
    "trip_end": {"en": ["Trip End", "End Time", "Arrival"]},
    "start_location": {"en": ["Start Location", "Origin", "From"]},
    "end_location": {"en": ["End Location", "Destination", "To"]},
    "purpose": {"en": ["Purpose", "Trip Purpose", "Journey Type"]},
    "engine_hours": {"en": ["Engine Hours", "Run Hours", "Hours"]},
    "idle_fuel": {"en": ["Idle Fuel", "Idling Fuel", "Fuel at Idle"]},
    "avg_consumption": {"en": ["Avg Consumption", "Fuel Economy", "L/100km"]},
}

_HEAT_FIELDS = [
    "invoice_no",
    "company",
    "site",
    "city",
    "postcode",
    "meter_id",
    "period",
    "period_start",
    "period_end",
    "issue_date",
    "due_date",
    "prev_read",
    "curr_read",
    "consumption",
    "unit_price",
    "heat_cost",
    "capacity",
    "cap_rate",
    "supplier_ef",
    "cap_charge",
    "subtotal",
    "vat",
    "total",
    "currency",
]

_ELECTRICITY_CORE_FIELDS = [
    "reference",
    "company",
    "site",
    "period",
    "period_start",
    "period_end",
    "city",
    "postcode",
    "meter_id",
    "supplier_ef",
    "unit",
    "start_read",
    "end_read",
    "total_qty",
    "total_cost",
    "currency",
    "emissions_kg",
    "emissions_t",
]

_SMART_METER_MONTHLY_FIELDS = [
    "meter_id",
    "site",
    "period",
    "start_read",
    "end_read",
    "consumption",
    "unit",
    "tariff_type",
    "tariff_cost",
    "currency",
]

_SMART_METER_INTERVAL_FIELDS = {
    "consumption_diff": ["meter_id", "timestamp", "import_kwh", "export_kwh", "unit"],
    "cumulative_end_reading": ["meter_id", "timestamp", "end_read", "unit"],
}

_FUEL_CARD_FIELDS = [
    "card_no",
    "date",
    "merchant",
    "site",
    "country",
    "equipment",
    "emission_source",
    "product",
    "qty",
    "unit",
    "unit_price",
    "total",
    "currency",
]

_GENERATOR_LOG_FIELDS = [
    "company",
    "site",
    "country",
    "date",
    "start_time",
    "end_time",
    "run_hours",
    "start_fuel",
    "end_fuel",
    "fuel_used",
    "unit",
    "equipment",
    "emission_source",
    "fuel_type",
    "notes",
]

_BEMS_EQUIPMENT_FIELDS = [
    "equipment_tag",
    "equipment_name",
    "emission_source",
    "fuel_type",
    "consumption",
    "unit",
    "operating_hours",
]

_BEMS_TIME_SERIES_FIELDS = [
    "timestamp",
    "site",
    "equipment_tag",
    "sensor_name",
    "value",
    "unit",
]

_MOBILE_FUEL_CARD_FIELDS = [
    "date",
    "time",
    "card_no",
    "vehicle_reg",
    "driver",
    "merchant",
    "site",
    "receipt_no",
    "odometer",
    "product",
    "qty",
    "unit",
    "unit_price",
    "total",
    "currency",
]

_MOBILE_TELEMATICS_FUEL_FIELDS = [
    "period_start",
    "period_end",
    "vehicle_reg",
    "vehicle_name",
    "vehicle_type",
    "fuel_type",
    "distance",
    "distance_unit",
    "fuel_used",
    "unit",
    "idle_fuel",
    "engine_hours",
    "avg_consumption",
]

_MOBILE_TELEMATICS_TRIP_FIELDS = [
    "vehicle_reg",
    "driver",
    "trip_start",
    "trip_end",
    "purpose",
    "start_location",
    "end_location",
    "distance",
    "distance_unit",
    "fuel_type",
]

_MOBILE_TELEMATICS_ODOMETER_FIELDS = [
    "period_start",
    "period_end",
    "vehicle_reg",
    "vehicle_name",
    "fuel_type",
    "odometer_start",
    "odometer_end",
    "distance",
    "distance_unit",
]


def _simple_tabular_spec(field_ids: list[str], *stress_orders: list[str]) -> dict[str, Any]:
    return {
        "field_ids": field_ids,
        "column_orders": {
            "default": field_ids,
            "realistic": [field_ids],
            "balanced": [field_ids, *stress_orders[:1]] if stress_orders else [field_ids],
            "stress": list(stress_orders) or [field_ids],
        },
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    }


_TABULAR_SPECS: dict[tuple[str, str, str], dict[str, Any]] = {
    ("heat", "supplier_portal_data", "CSV"): {
        "field_ids": _HEAT_FIELDS,
        "column_orders": {
            "default": _HEAT_FIELDS,
            "realistic": [
                _HEAT_FIELDS,
                ["invoice_no", "company", "site", "period", "period_start", "period_end", "issue_date", "due_date", "meter_id", "city", "postcode", "prev_read", "curr_read", "consumption", "unit_price", "heat_cost", "capacity", "cap_rate", "cap_charge", "subtotal", "vat", "total", "supplier_ef", "currency"],
            ],
            "balanced": [
                _HEAT_FIELDS,
                ["company", "site", "city", "postcode", "meter_id", "invoice_no", "period", "period_start", "period_end", "issue_date", "due_date", "prev_read", "curr_read", "consumption", "capacity", "cap_rate", "unit_price", "heat_cost", "cap_charge", "subtotal", "vat", "total", "supplier_ef", "currency"],
                ["invoice_no", "period", "company", "site", "meter_id", "city", "postcode", "period_start", "period_end", "prev_read", "curr_read", "consumption", "unit_price", "heat_cost", "capacity", "cap_rate", "supplier_ef", "cap_charge", "subtotal", "vat", "total", "issue_date", "due_date", "currency"],
            ],
            "stress": [
                ["company", "site", "meter_id", "invoice_no", "period", "period_start", "period_end", "city", "postcode", "consumption", "prev_read", "curr_read", "capacity", "cap_rate", "unit_price", "heat_cost", "cap_charge", "subtotal", "vat", "total", "supplier_ef", "issue_date", "due_date", "currency"],
                ["period", "invoice_no", "company", "site", "city", "postcode", "meter_id", "issue_date", "due_date", "period_start", "period_end", "capacity", "cap_rate", "prev_read", "curr_read", "consumption", "unit_price", "heat_cost", "supplier_ef", "cap_charge", "subtotal", "vat", "total", "currency"],
            ],
        },
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none", "title_only"], "balanced": ["none", "title_only", "title_period"], "stress": ["none", "title_period", "title_subject", "audit_banner"]},
    },
    ("heat", "supplier_portal_data", "XLSX"): {
        "field_ids": _HEAT_FIELDS,
        "column_orders": {
            "default": _HEAT_FIELDS,
            "realistic": [_HEAT_FIELDS],
            "balanced": [
                _HEAT_FIELDS,
                ["invoice_no", "company", "site", "period", "meter_id", "city", "postcode", "period_start", "period_end", "issue_date", "due_date", "prev_read", "curr_read", "consumption", "unit_price", "heat_cost", "capacity", "cap_rate", "supplier_ef", "cap_charge", "subtotal", "vat", "total", "currency"],
            ],
            "stress": [
                ["company", "site", "meter_id", "invoice_no", "period", "period_start", "period_end", "city", "postcode", "consumption", "prev_read", "curr_read", "capacity", "cap_rate", "unit_price", "heat_cost", "cap_charge", "subtotal", "vat", "total", "supplier_ef", "issue_date", "due_date", "currency"],
            ],
        },
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
        "summary_sheet_positions": {"default": "first", "realistic": ["first"], "balanced": ["first", "last"], "stress": ["first", "last"]},
        "company_sheet_orders": {"default": "forward", "realistic": ["forward"], "balanced": ["forward", "reverse"], "stress": ["forward", "reverse"]},
    },
    ("electricity", "supplier_portal_data", "CSV"): {
        "field_ids": _ELECTRICITY_CORE_FIELDS,
        "column_orders": {
            "default": _ELECTRICITY_CORE_FIELDS,
            "realistic": [_ELECTRICITY_CORE_FIELDS],
            "balanced": [
                _ELECTRICITY_CORE_FIELDS,
                ["reference", "company", "site", "period", "meter_id", "city", "postcode", "period_start", "period_end", "supplier_ef", "unit", "start_read", "end_read", "total_qty", "emissions_kg", "emissions_t", "total_cost", "currency"],
            ],
            "stress": [
                ["company", "site", "meter_id", "reference", "period", "period_start", "period_end", "city", "postcode", "unit", "start_read", "end_read", "total_qty", "supplier_ef", "emissions_kg", "emissions_t", "total_cost", "currency"],
                ["period", "reference", "company", "site", "city", "postcode", "meter_id", "supplier_ef", "unit", "total_qty", "total_cost", "currency", "emissions_kg", "emissions_t", "start_read", "end_read", "period_start", "period_end"],
            ],
        },
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none", "title_only"], "balanced": ["none", "title_only", "title_period"], "stress": ["none", "title_period", "audit_banner"]},
        "tariff_block_positions": {"default": "tail", "realistic": ["tail"], "balanced": ["tail", "after_total_qty"], "stress": ["tail", "after_total_qty"]},
    },
    ("electricity", "supplier_portal_data", "XLSX"): {
        "field_ids": _ELECTRICITY_CORE_FIELDS,
        "column_orders": {
            "default": _ELECTRICITY_CORE_FIELDS,
            "realistic": [_ELECTRICITY_CORE_FIELDS],
            "balanced": [
                _ELECTRICITY_CORE_FIELDS,
                ["reference", "company", "site", "period", "meter_id", "city", "postcode", "period_start", "period_end", "supplier_ef", "unit", "start_read", "end_read", "total_qty", "emissions_kg", "emissions_t", "total_cost", "currency"],
            ],
            "stress": [
                ["company", "site", "meter_id", "reference", "period", "period_start", "period_end", "city", "postcode", "unit", "start_read", "end_read", "total_qty", "supplier_ef", "emissions_kg", "emissions_t", "total_cost", "currency"],
            ],
        },
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
        "summary_sheet_positions": {"default": "first", "realistic": ["first"], "balanced": ["first", "last"], "stress": ["first", "last"]},
        "tariff_block_positions": {"default": "tail", "realistic": ["tail"], "balanced": ["tail", "after_total_qty"], "stress": ["tail", "after_total_qty"]},
    },
    ("electricity", "smart_meter_data", "CSV"): {
        "field_ids": _SMART_METER_MONTHLY_FIELDS,
        "mode_field_ids": {
            "monthly": _SMART_METER_MONTHLY_FIELDS,
            "interval:consumption_diff": _SMART_METER_INTERVAL_FIELDS["consumption_diff"],
            "interval:cumulative_end_reading": _SMART_METER_INTERVAL_FIELDS["cumulative_end_reading"],
        },
        "column_orders": {"default": _SMART_METER_MONTHLY_FIELDS, "realistic": [_SMART_METER_MONTHLY_FIELDS], "balanced": [_SMART_METER_MONTHLY_FIELDS], "stress": [_SMART_METER_MONTHLY_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none", "title_only"], "balanced": ["none", "title_period"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("electricity", "smart_meter_data", "XLSX"): {
        "field_ids": _SMART_METER_MONTHLY_FIELDS,
        "mode_field_ids": {
            "monthly": _SMART_METER_MONTHLY_FIELDS,
            "interval:consumption_diff": _SMART_METER_INTERVAL_FIELDS["consumption_diff"],
            "interval:cumulative_end_reading": _SMART_METER_INTERVAL_FIELDS["cumulative_end_reading"],
        },
        "column_orders": {"default": _SMART_METER_MONTHLY_FIELDS, "realistic": [_SMART_METER_MONTHLY_FIELDS], "balanced": [_SMART_METER_MONTHLY_FIELDS], "stress": [_SMART_METER_MONTHLY_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "fuel_card", "CSV"): {
        "field_ids": _FUEL_CARD_FIELDS,
        "column_orders": {
            "default": _FUEL_CARD_FIELDS,
            "realistic": [_FUEL_CARD_FIELDS],
            "balanced": [_FUEL_CARD_FIELDS, ["date", "card_no", "merchant", "site", "country", "product", "qty", "unit", "unit_price", "total", "currency", "equipment", "emission_source"]],
            "stress": [["merchant", "date", "card_no", "site", "country", "equipment", "emission_source", "product", "qty", "unit", "unit_price", "total", "currency"]],
        },
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "fuel_card", "XLSX"): {
        "field_ids": _FUEL_CARD_FIELDS,
        "column_orders": {
            "default": _FUEL_CARD_FIELDS,
            "realistic": [_FUEL_CARD_FIELDS],
            "balanced": [_FUEL_CARD_FIELDS, ["date", "card_no", "merchant", "site", "country", "product", "qty", "unit", "unit_price", "total", "currency", "equipment", "emission_source"]],
            "stress": [["merchant", "date", "card_no", "site", "country", "equipment", "emission_source", "product", "qty", "unit", "unit_price", "total", "currency"]],
        },
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "generator_log", "CSV"): {
        "field_ids": _GENERATOR_LOG_FIELDS,
        "column_orders": {"default": _GENERATOR_LOG_FIELDS, "realistic": [_GENERATOR_LOG_FIELDS], "balanced": [_GENERATOR_LOG_FIELDS], "stress": [_GENERATOR_LOG_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "generator_log", "XLSX"): {
        "field_ids": _GENERATOR_LOG_FIELDS,
        "column_orders": {"default": _GENERATOR_LOG_FIELDS, "realistic": [_GENERATOR_LOG_FIELDS], "balanced": [_GENERATOR_LOG_FIELDS], "stress": [_GENERATOR_LOG_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "bems_equipment_report", "CSV"): {
        "field_ids": _BEMS_EQUIPMENT_FIELDS,
        "column_orders": {"default": _BEMS_EQUIPMENT_FIELDS, "realistic": [_BEMS_EQUIPMENT_FIELDS], "balanced": [_BEMS_EQUIPMENT_FIELDS], "stress": [_BEMS_EQUIPMENT_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "bems_equipment_report", "XLSX"): {
        "field_ids": _BEMS_EQUIPMENT_FIELDS,
        "column_orders": {"default": _BEMS_EQUIPMENT_FIELDS, "realistic": [_BEMS_EQUIPMENT_FIELDS], "balanced": [_BEMS_EQUIPMENT_FIELDS], "stress": [_BEMS_EQUIPMENT_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "bems_time_series", "CSV"): {
        "field_ids": _BEMS_TIME_SERIES_FIELDS,
        "column_orders": {"default": _BEMS_TIME_SERIES_FIELDS, "realistic": [_BEMS_TIME_SERIES_FIELDS], "balanced": [_BEMS_TIME_SERIES_FIELDS], "stress": [_BEMS_TIME_SERIES_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("stationary_combustion", "bems_time_series", "XLSX"): {
        "field_ids": _BEMS_TIME_SERIES_FIELDS,
        "column_orders": {"default": _BEMS_TIME_SERIES_FIELDS, "realistic": [_BEMS_TIME_SERIES_FIELDS], "balanced": [_BEMS_TIME_SERIES_FIELDS], "stress": [_BEMS_TIME_SERIES_FIELDS]},
        "header_row_offsets": {"default": 0, "realistic": [0], "balanced": [0, 1], "stress": [0, 1, 2]},
        "preamble_templates": {"default": ["none"], "realistic": ["none"], "balanced": ["none", "title_only"], "stress": ["none", "title_period", "audit_banner"]},
    },
    ("mobile_combustion", "fuel_card_statement", "CSV"): _simple_tabular_spec(
        _MOBILE_FUEL_CARD_FIELDS,
        ["card_no", "date", "time", "merchant", "site", "vehicle_reg", "driver", "product", "qty", "unit", "unit_price", "total", "currency", "receipt_no", "odometer"],
        ["receipt_no", "date", "time", "card_no", "vehicle_reg", "odometer", "merchant", "site", "driver", "product", "qty", "unit", "unit_price", "total", "currency"],
    ),
    ("mobile_combustion", "fuel_card_statement", "XLSX"): _simple_tabular_spec(
        _MOBILE_FUEL_CARD_FIELDS,
        ["card_no", "date", "time", "merchant", "site", "vehicle_reg", "driver", "product", "qty", "unit", "unit_price", "total", "currency", "receipt_no", "odometer"],
    ),
    ("mobile_combustion", "telematics_fuel", "CSV"): _simple_tabular_spec(
        _MOBILE_TELEMATICS_FUEL_FIELDS,
        ["vehicle_reg", "vehicle_name", "vehicle_type", "fuel_type", "period_start", "period_end", "fuel_used", "unit", "distance", "distance_unit", "idle_fuel", "engine_hours", "avg_consumption"],
    ),
    ("mobile_combustion", "telematics_fuel", "XLSX"): _simple_tabular_spec(
        _MOBILE_TELEMATICS_FUEL_FIELDS,
        ["vehicle_reg", "vehicle_name", "vehicle_type", "fuel_type", "period_start", "period_end", "fuel_used", "unit", "distance", "distance_unit", "idle_fuel", "engine_hours", "avg_consumption"],
    ),
    ("mobile_combustion", "telematics_trips", "CSV"): _simple_tabular_spec(
        _MOBILE_TELEMATICS_TRIP_FIELDS,
        ["vehicle_reg", "driver", "trip_start", "trip_end", "purpose", "start_location", "end_location", "distance", "distance_unit", "fuel_type"],
    ),
    ("mobile_combustion", "telematics_trips", "XLSX"): _simple_tabular_spec(
        _MOBILE_TELEMATICS_TRIP_FIELDS,
        ["vehicle_reg", "driver", "trip_start", "trip_end", "purpose", "start_location", "end_location", "distance", "distance_unit", "fuel_type"],
    ),
    ("mobile_combustion", "telematics_odometer", "CSV"): _simple_tabular_spec(
        _MOBILE_TELEMATICS_ODOMETER_FIELDS,
        ["vehicle_reg", "vehicle_name", "fuel_type", "odometer_start", "odometer_end", "distance", "distance_unit", "period_start", "period_end"],
    ),
    ("mobile_combustion", "telematics_odometer", "XLSX"): _simple_tabular_spec(
        _MOBILE_TELEMATICS_ODOMETER_FIELDS,
        ["vehicle_reg", "vehicle_name", "fuel_type", "odometer_start", "odometer_end", "distance", "distance_unit", "period_start", "period_end"],
    ),
}

_DOCUMENT_SPECS: dict[tuple[str, str, str], dict[str, Any]] = {
    ("heat", "utility_bill", "PDF"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "stacked-meta"], "stress": ["classic", "stacked-meta", "charges-first"]},
        "section_orders": {
            "default": ["addresses", "meta", "billing_fields", "charges", "footer"],
            "realistic": [["addresses", "meta", "billing_fields", "charges", "footer"]],
            "balanced": [["meta", "addresses", "billing_fields", "charges", "footer"], ["addresses", "meta", "charges", "billing_fields", "footer"]],
            "stress": [["meta", "addresses", "charges", "billing_fields", "footer"], ["addresses", "charges", "meta", "billing_fields", "footer"]],
        },
        "table_transforms": {"default": {"billing_fields": "vertical", "charges": "vertical"}, "realistic": [{"billing_fields": "vertical", "charges": "vertical"}], "balanced": [{"billing_fields": "vertical", "charges": "vertical"}, {"billing_fields": "transposed", "charges": "vertical"}], "stress": [{"billing_fields": "transposed", "charges": "vertical"}]},
    },
    ("heat", "utility_bill", "DOCX"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "stacked-meta"], "stress": ["classic", "stacked-meta", "charges-first"]},
        "section_orders": {
            "default": ["addresses", "meta", "billing_fields", "charges", "footer"],
            "realistic": [["addresses", "meta", "billing_fields", "charges", "footer"]],
            "balanced": [["meta", "addresses", "billing_fields", "charges", "footer"], ["addresses", "meta", "charges", "billing_fields", "footer"]],
            "stress": [["meta", "addresses", "charges", "billing_fields", "footer"], ["addresses", "charges", "meta", "billing_fields", "footer"]],
        },
        "table_transforms": {"default": {"billing_fields": "vertical", "charges": "vertical"}, "realistic": [{"billing_fields": "vertical", "charges": "vertical"}], "balanced": [{"billing_fields": "vertical", "charges": "vertical"}, {"billing_fields": "transposed", "charges": "vertical"}], "stress": [{"billing_fields": "transposed", "charges": "vertical"}]},
    },
    ("electricity", "electricity_bill", "PDF"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "grid-first"], "stress": ["classic", "grid-first", "totals-early"]},
        "section_orders": {
            "default": ["addresses", "period_meta", "meter_table", "grid_table", "tariff_table", "total_box", "footer"],
            "realistic": [["addresses", "period_meta", "meter_table", "grid_table", "tariff_table", "total_box", "footer"]],
            "balanced": [["period_meta", "addresses", "meter_table", "grid_table", "tariff_table", "total_box", "footer"], ["addresses", "period_meta", "grid_table", "meter_table", "tariff_table", "total_box", "footer"]],
            "stress": [["period_meta", "addresses", "grid_table", "meter_table", "total_box", "tariff_table", "footer"], ["addresses", "period_meta", "meter_table", "tariff_table", "grid_table", "total_box", "footer"]],
        },
        "table_transforms": {"default": {"meter_table": "vertical", "grid_table": "vertical"}, "realistic": [{"meter_table": "vertical", "grid_table": "vertical"}], "balanced": [{"meter_table": "vertical", "grid_table": "vertical"}, {"meter_table": "transposed", "grid_table": "vertical"}], "stress": [{"meter_table": "transposed", "grid_table": "vertical"}]},
    },
    ("electricity", "electricity_bill", "DOCX"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "grid-first"], "stress": ["classic", "grid-first", "totals-early"]},
        "section_orders": {
            "default": ["addresses", "period_meta", "meter_table", "grid_table", "tariff_table", "total_box", "footer"],
            "realistic": [["addresses", "period_meta", "meter_table", "grid_table", "tariff_table", "total_box", "footer"]],
            "balanced": [["period_meta", "addresses", "meter_table", "grid_table", "tariff_table", "total_box", "footer"], ["addresses", "period_meta", "grid_table", "meter_table", "tariff_table", "total_box", "footer"]],
            "stress": [["period_meta", "addresses", "grid_table", "meter_table", "total_box", "tariff_table", "footer"], ["addresses", "period_meta", "meter_table", "tariff_table", "grid_table", "total_box", "footer"]],
        },
        "table_transforms": {"default": {"meter_table": "vertical", "grid_table": "vertical"}, "realistic": [{"meter_table": "vertical", "grid_table": "vertical"}], "balanced": [{"meter_table": "vertical", "grid_table": "vertical"}, {"meter_table": "transposed", "grid_table": "vertical"}], "stress": [{"meter_table": "transposed", "grid_table": "vertical"}]},
    },
    ("stationary_combustion", "fuel_invoice", "PDF"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "meta-first"], "stress": ["classic", "meta-first"]},
        "section_orders": {"default": ["addresses", "meta", "line_items", "totals", "footer"], "realistic": [["addresses", "meta", "line_items", "totals", "footer"]], "balanced": [["meta", "addresses", "line_items", "totals", "footer"], ["addresses", "line_items", "meta", "totals", "footer"]], "stress": [["meta", "addresses", "totals", "line_items", "footer"]]},
    },
    ("stationary_combustion", "fuel_invoice", "DOCX"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "meta-first"], "stress": ["classic", "meta-first"]},
        "section_orders": {"default": ["addresses", "meta", "line_items", "totals", "footer"], "realistic": [["addresses", "meta", "line_items", "totals", "footer"]], "balanced": [["meta", "addresses", "line_items", "totals", "footer"], ["addresses", "line_items", "meta", "totals", "footer"]], "stress": [["meta", "addresses", "totals", "line_items", "footer"]]},
    },
    ("stationary_combustion", "delivery_note", "PDF"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "details-first"], "stress": ["classic", "details-first"]},
        "section_orders": {"default": ["header", "addresses", "delivery_details", "footer"], "realistic": [["header", "addresses", "delivery_details", "footer"]], "balanced": [["header", "delivery_details", "addresses", "footer"]], "stress": [["delivery_details", "header", "addresses", "footer"]]},
    },
    ("stationary_combustion", "delivery_note", "DOCX"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "details-first"], "stress": ["classic", "details-first"]},
        "section_orders": {"default": ["header", "addresses", "delivery_details", "footer"], "realistic": [["header", "addresses", "delivery_details", "footer"]], "balanced": [["header", "delivery_details", "addresses", "footer"]], "stress": [["delivery_details", "header", "addresses", "footer"]]},
    },
    ("stationary_combustion", "fuel_card", "PDF"): {
        "base_families": {"default": "statement", "realistic": ["statement"], "balanced": ["statement", "summary-first"], "stress": ["statement", "summary-first"]},
        "section_orders": {"default": ["header", "summary", "transactions", "footer"], "realistic": [["header", "summary", "transactions", "footer"]], "balanced": [["summary", "header", "transactions", "footer"]], "stress": [["header", "transactions", "summary", "footer"]]},
    },
    ("stationary_combustion", "fuel_card", "DOCX"): {
        "base_families": {"default": "statement", "realistic": ["statement"], "balanced": ["statement", "summary-first"], "stress": ["statement", "summary-first"]},
        "section_orders": {"default": ["header", "summary", "transactions", "footer"], "realistic": [["header", "summary", "transactions", "footer"]], "balanced": [["summary", "header", "transactions", "footer"]], "stress": [["header", "transactions", "summary", "footer"]]},
    },
    ("stationary_combustion", "bems_equipment_report", "PDF"): {
        "base_families": {"default": "report", "realistic": ["report"], "balanced": ["report", "table-first"], "stress": ["report", "table-first"]},
        "section_orders": {"default": ["header", "meta", "table", "footer"], "realistic": [["header", "meta", "table", "footer"]], "balanced": [["header", "table", "meta", "footer"]], "stress": [["table", "header", "meta", "footer"]]},
    },
    ("stationary_combustion", "bems_equipment_report", "DOCX"): {
        "base_families": {"default": "report", "realistic": ["report"], "balanced": ["report", "table-first"], "stress": ["report", "table-first"]},
        "section_orders": {"default": ["header", "meta", "table", "footer"], "realistic": [["header", "meta", "table", "footer"]], "balanced": [["header", "table", "meta", "footer"]], "stress": [["table", "header", "meta", "footer"]]},
    },
    ("stationary_combustion", "bems_time_series", "PDF"): {
        "base_families": {"default": "report", "realistic": ["report"], "balanced": ["report"], "stress": ["report"]},
        "section_orders": {"default": ["header", "meta", "table", "footer"], "realistic": [["header", "meta", "table", "footer"]], "balanced": [["header", "table", "meta", "footer"]], "stress": [["table", "header", "meta", "footer"]]},
    },
    ("stationary_combustion", "bems_time_series", "DOCX"): {
        "base_families": {"default": "report", "realistic": ["report"], "balanced": ["report"], "stress": ["report"]},
        "section_orders": {"default": ["header", "meta", "table", "footer"], "realistic": [["header", "meta", "table", "footer"]], "balanced": [["header", "table", "meta", "footer"]], "stress": [["table", "header", "meta", "footer"]]},
    },
    ("mobile_combustion", "fuel_invoice", "PDF"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "meta-first"], "stress": ["classic", "meta-first"]},
        "section_orders": {"default": ["addresses", "meta", "line_items", "totals", "footer"], "realistic": [["addresses", "meta", "line_items", "totals", "footer"]], "balanced": [["meta", "addresses", "line_items", "totals", "footer"], ["addresses", "line_items", "meta", "totals", "footer"]], "stress": [["meta", "addresses", "totals", "line_items", "footer"]]},
    },
    ("mobile_combustion", "fuel_invoice", "DOCX"): {
        "base_families": {"default": "classic", "realistic": ["classic"], "balanced": ["classic", "meta-first"], "stress": ["classic", "meta-first"]},
        "section_orders": {"default": ["addresses", "meta", "line_items", "totals", "footer"], "realistic": [["addresses", "meta", "line_items", "totals", "footer"]], "balanced": [["meta", "addresses", "line_items", "totals", "footer"], ["addresses", "line_items", "meta", "totals", "footer"]], "stress": [["meta", "addresses", "totals", "line_items", "footer"]]},
    },
    ("mobile_combustion", "fuel_card_statement", "PDF"): {
        "base_families": {"default": "statement", "realistic": ["statement"], "balanced": ["statement", "summary-first"], "stress": ["statement", "summary-first"]},
        "section_orders": {"default": ["summary", "transactions", "footer"], "realistic": [["summary", "transactions", "footer"]], "balanced": [["summary", "transactions", "footer"], ["transactions", "summary", "footer"]], "stress": [["transactions", "summary", "footer"]]},
    },
    ("mobile_combustion", "fuel_card_statement", "DOCX"): {
        "base_families": {"default": "statement", "realistic": ["statement"], "balanced": ["statement", "summary-first"], "stress": ["statement", "summary-first"]},
        "section_orders": {"default": ["summary", "transactions", "footer"], "realistic": [["summary", "transactions", "footer"]], "balanced": [["summary", "transactions", "footer"], ["transactions", "summary", "footer"]], "stress": [["transactions", "summary", "footer"]]},
    },
}


def normalize_layout_settings(document_cfg: dict, random_seed: int) -> dict:
    layout_cfg = document_cfg.get("layout") or {}
    preset = str(layout_cfg.get("preset") or "balanced").lower()
    if preset not in _VALID_PRESETS:
        preset = "balanced"

    resolved_seed = layout_cfg.get("seed")
    if resolved_seed in (None, ""):
        resolved_seed = int(random_seed)
    else:
        resolved_seed = int(resolved_seed)

    return {
        "enabled": bool(layout_cfg.get("enabled", False)),
        "preset": preset,
        "seed": resolved_seed,
    }


def resolve_layout_plan(
    document_cfg: dict,
    *,
    random_seed: int,
    category: str,
    document_type: str,
    output_format: str,
    artifact_key: str = "default",
    context: dict | None = None,
) -> dict:
    layout = normalize_layout_settings(document_cfg, random_seed)
    context = context or {}
    lang = str(context.get("language") or document_cfg.get("language") or "en").lower()
    lookup_document_type = document_type
    if category == "stationary_combustion" and document_type == "bems":
        report_type = str(context.get("bems_report_type") or document_cfg.get("bems_report_type") or "equipment_trend_report")
        lookup_document_type = "bems_time_series" if report_type == "time_series_trend_export" else "bems_equipment_report"
    plan = {
        "enabled": layout["enabled"],
        "preset": layout["preset"],
        "seed": layout["seed"],
        "slug": None,
        "base_family": "default",
        "header_aliases": {},
        "column_order": [],
        "header_row_offset": 0,
        "preamble_rows": [],
        "sheet_order": [],
        "company_sheet_order": [],
        "section_order": [],
        "table_transforms": {},
        "totals_position": None,
        "tariff_block_position": "tail",
    }

    tabular_spec = _TABULAR_SPECS.get((category, lookup_document_type, output_format))
    document_spec = _DOCUMENT_SPECS.get((category, lookup_document_type, output_format))

    material = {
        "seed": layout["seed"],
        "preset": layout["preset"],
        "category": category,
        "document_type": lookup_document_type,
        "output_format": output_format,
        "artifact_key": artifact_key,
    }
    rng = random.Random(_material_seed(material))

    if tabular_spec is not None and layout["enabled"]:
        field_ids = _resolve_tabular_field_ids(tabular_spec, context)
        selected_order = _select_option(tabular_spec.get("column_orders", {}), layout["preset"], rng, default=field_ids)
        plan["column_order"] = _ordered_field_ids(selected_order, field_ids)
        plan["header_row_offset"] = _select_scalar(tabular_spec.get("header_row_offsets", {}), layout["preset"], rng, default=0)
        preamble_name = _select_scalar(tabular_spec.get("preamble_templates", {}), layout["preset"], rng, default="none")
        plan["preamble_rows"] = _render_preamble_rows(preamble_name, document_cfg, layout["preset"], context)
        plan["header_aliases"] = _resolve_header_aliases(plan["column_order"], lang, layout["enabled"], layout["preset"], rng)
        plan["tariff_block_position"] = _select_scalar(tabular_spec.get("tariff_block_positions", {}), layout["preset"], rng, default="tail")
        plan.update(_resolve_sheet_plan(tabular_spec, layout["preset"], rng, context))

    if document_spec is not None and layout["enabled"]:
        plan["base_family"] = _select_scalar(document_spec.get("base_families", {}), layout["preset"], rng, default="default")
        plan["section_order"] = _select_option(document_spec.get("section_orders", {}), layout["preset"], rng, default=document_spec.get("section_orders", {}).get("default", []))
        plan["table_transforms"] = _select_option(document_spec.get("table_transforms", {}), layout["preset"], rng, default=document_spec.get("table_transforms", {}).get("default", {}))

    if layout["enabled"]:
        material.update({
            "base_family": plan["base_family"],
            "column_order": plan["column_order"],
            "header_row_offset": plan["header_row_offset"],
            "preamble_rows": plan["preamble_rows"],
            "sheet_order": plan["sheet_order"],
            "section_order": plan["section_order"],
            "table_transforms": plan["table_transforms"],
            "tariff_block_position": plan["tariff_block_position"],
        })
        plan["slug"] = f"layout-{hashlib.sha1(json.dumps(material, sort_keys=True).encode('utf-8')).hexdigest()[:8]}"

    return plan


def apply_layout_suffix(stem: str, layout_slug: str | None) -> str:
    if not layout_slug:
        return stem
    return f"{stem}_{layout_slug}"


def _material_seed(material: dict[str, Any]) -> int:
    digest = hashlib.sha1(json.dumps(material, sort_keys=True).encode("utf-8")).hexdigest()
    return int(digest[:16], 16)


def _resolve_tabular_field_ids(spec: dict[str, Any], context: dict[str, Any]) -> list[str]:
    mode_field_ids = spec.get("mode_field_ids")
    if not isinstance(mode_field_ids, dict):
        return list(spec.get("field_ids", []))

    smart_meter_mode = str(context.get("smart_meter_mode") or "monthly").lower()
    if smart_meter_mode == "interval":
        interval_mode = str(context.get("smart_meter_value_mode") or "consumption_diff").lower()
        return list(mode_field_ids.get(f"interval:{interval_mode}", mode_field_ids.get("interval:consumption_diff", spec.get("field_ids", []))))
    return list(mode_field_ids.get("monthly", spec.get("field_ids", [])))


def _select_option(spec: dict[str, Any], preset: str, rng: random.Random, *, default: Any) -> Any:
    if not spec:
        return _clone(default)
    options = spec.get(preset)
    if not options:
        return _clone(spec.get("default", default))
    return _clone(rng.choice(options))


def _select_scalar(spec: dict[str, Any], preset: str, rng: random.Random, *, default: Any) -> Any:
    if not spec:
        return default
    options = spec.get(preset)
    if not options:
        return spec.get("default", default)
    return rng.choice(options)


def _resolve_header_aliases(
    field_ids: list[str],
    language: str,
    enabled: bool,
    preset: str,
    rng: random.Random,
) -> dict[str, str]:
    if not enabled:
        return {}

    alias_rate = _ALIAS_RATES.get(preset, _ALIAS_RATES["balanced"])
    aliases: dict[str, str] = {}
    for field_id in field_ids:
        pool = _FIELD_ALIAS_POOLS.get(field_id, {}).get(language)
        if not pool:
            continue
        if rng.random() <= alias_rate:
            aliases[field_id] = rng.choice(pool)
    return aliases


def _render_preamble_rows(template_name: str, document_cfg: dict, preset: str, context: dict[str, Any]) -> list[list[str]]:
    template = _PREAMBLE_TEMPLATES.get(template_name, [])
    values = {
        "title": document_cfg.get("title", "Document"),
        "subject": document_cfg.get("subject", ""),
        "period_label": context.get("period_label", ""),
        "preset": preset,
    }
    rendered: list[list[str]] = []
    for row in template:
        rendered.append([str(cell).format(**values).strip() for cell in row])
    return rendered


def _resolve_sheet_plan(spec: dict[str, Any], preset: str, rng: random.Random, context: dict[str, Any]) -> dict[str, Any]:
    include_summary = bool(context.get("include_summary", False))
    split_by_company = bool(context.get("split_by_company", False))
    company_labels = list(context.get("company_labels", []))
    company_sheet_order = company_labels[:]

    order_mode = _select_scalar(spec.get("company_sheet_orders", {}), preset, rng, default="forward")
    if split_by_company and order_mode == "reverse":
        company_sheet_order = list(reversed(company_sheet_order))

    summary_position = _select_scalar(spec.get("summary_sheet_positions", {}), preset, rng, default="first")

    if split_by_company:
        sheet_order = company_sheet_order[:]
    else:
        sheet_order = ["detail"] if spec.get("summary_sheet_positions") is not None else []

    if include_summary:
        if summary_position == "last":
            sheet_order = sheet_order + ["summary"]
        else:
            sheet_order = ["summary"] + sheet_order

    return {
        "sheet_order": sheet_order,
        "company_sheet_order": company_sheet_order,
    }


def _clone(value: Any) -> Any:
    if isinstance(value, list):
        return [
            _clone(item)
            for item in value
        ]
    if isinstance(value, dict):
        return {key: _clone(item) for key, item in value.items()}
    return value


def _ordered_field_ids(selected_order: list[str], available_ids: list[str]) -> list[str]:
    ordered = [field_id for field_id in selected_order if field_id in available_ids]
    ordered.extend(field_id for field_id in available_ids if field_id not in ordered)
    return ordered