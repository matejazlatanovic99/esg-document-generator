from __future__ import annotations

import copy
import json
import unittest
from io import BytesIO
from zipfile import ZipFile

from generators.electricity_generator import build_smart_meter_rows
from utils.document_catalog import DOCUMENT_TYPES
from utils.document_catalog import (
    document_type_requires_company_currency,
    supports_monthly_zip_export,
)
from utils.generator import (
    _DOCUMENT_GENERATOR_DISPATCH,
    _build_electricity_config,
    _prepare_electricity_sections,
    _prepare_heat_sections,
    build_document_download_filename,
    generate_document_bytes,
    generate_json_ground_truth,
)


_CATEGORY_DISPATCH_KEY = {
    "purchased_heat_steam_cooling": "heat",
    "electricity": "electricity",
    "stationary_combustion": "stationary_combustion",
    "mobile_combustion": "mobile_combustion",
}


class GeneratorContractTests(unittest.TestCase):
    def test_document_catalog_capabilities_cover_current_special_cases(self) -> None:
        self.assertTrue(supports_monthly_zip_export("purchased_heat_steam_cooling", "utility_bill", "PDF"))
        self.assertTrue(supports_monthly_zip_export("electricity", "electricity_bill", "DOCX"))
        self.assertFalse(supports_monthly_zip_export("electricity", "supplier_portal_data", "CSV"))
        self.assertFalse(supports_monthly_zip_export("stationary_combustion", "fuel_invoice", "PDF"))

        self.assertTrue(document_type_requires_company_currency("stationary_combustion", "fuel_invoice"))
        self.assertTrue(document_type_requires_company_currency("stationary_combustion", "fuel_card"))
        self.assertFalse(document_type_requires_company_currency("stationary_combustion", "generator_log"))

    def test_dispatch_matrix_matches_sidebar_metadata(self) -> None:
        expected = {
            (_CATEGORY_DISPATCH_KEY[category], document_type, output_format)
            for category, document_types in DOCUMENT_TYPES.items()
            for document_type, config in document_types.items()
            if config.get("implemented", False)
            for output_format in config.get("formats", [])
        }

        self.assertEqual(expected, set(_DOCUMENT_GENERATOR_DISPATCH))

    def test_supported_generation_matrix_smoke(self) -> None:
        cases = [
            ("heat supplier portal csv", self._heat_config("supplier_portal_data"), "CSV"),
            ("heat supplier portal xlsx", self._heat_config("supplier_portal_data"), "XLSX"),
            ("heat utility bill pdf", self._heat_config("utility_bill"), "PDF"),
            ("heat utility bill docx", self._heat_config("utility_bill"), "DOCX"),
            ("electricity supplier portal csv", self._electricity_config("supplier_portal_data"), "CSV"),
            ("electricity supplier portal xlsx", self._electricity_config("supplier_portal_data"), "XLSX"),
            ("electricity smart meter csv", self._electricity_config("smart_meter_data", smart_meter_data_granularity="monthly"), "CSV"),
            ("electricity smart meter xlsx", self._electricity_config("smart_meter_data", smart_meter_data_granularity="monthly"), "XLSX"),
            ("electricity bill pdf", self._electricity_config("electricity_bill"), "PDF"),
            ("electricity bill docx", self._electricity_config("electricity_bill"), "DOCX"),
            ("stationary fuel invoice pdf", self._stationary_config("fuel_invoice"), "PDF"),
            ("stationary fuel invoice docx", self._stationary_config("fuel_invoice"), "DOCX"),
            ("stationary delivery note pdf", self._stationary_config("delivery_note"), "PDF"),
            ("stationary delivery note docx", self._stationary_config("delivery_note"), "DOCX"),
            ("stationary fuel card csv", self._stationary_config("fuel_card"), "CSV"),
            ("stationary fuel card xlsx", self._stationary_config("fuel_card"), "XLSX"),
            ("stationary fuel card pdf", self._stationary_config("fuel_card"), "PDF"),
            ("stationary fuel card docx", self._stationary_config("fuel_card"), "DOCX"),
            ("stationary generator log csv", self._stationary_config("generator_log"), "CSV"),
            ("stationary generator log xlsx", self._stationary_config("generator_log"), "XLSX"),
            ("stationary bems equipment pdf", self._stationary_config("bems", bems_report_type="equipment_trend_report"), "PDF"),
            ("stationary bems equipment docx", self._stationary_config("bems", bems_report_type="equipment_trend_report"), "DOCX"),
            ("stationary bems equipment csv", self._stationary_config("bems", bems_report_type="equipment_trend_report"), "CSV"),
            ("stationary bems equipment xlsx", self._stationary_config("bems", bems_report_type="equipment_trend_report"), "XLSX"),
            ("stationary bems time series pdf", self._stationary_config("bems", bems_report_type="time_series_trend_export"), "PDF"),
            ("stationary bems time series docx", self._stationary_config("bems", bems_report_type="time_series_trend_export"), "DOCX"),
            ("stationary bems time series csv", self._stationary_config("bems", bems_report_type="time_series_trend_export"), "CSV"),
            ("stationary bems time series xlsx", self._stationary_config("bems", bems_report_type="time_series_trend_export"), "XLSX"),
            ("mobile fuel invoice pdf", self._mobile_config("fuel_invoice"), "PDF"),
            ("mobile fuel invoice docx", self._mobile_config("fuel_invoice"), "DOCX"),
            ("mobile fuel card statement pdf", self._mobile_config("fuel_card_statement"), "PDF"),
            ("mobile fuel card statement docx", self._mobile_config("fuel_card_statement"), "DOCX"),
            ("mobile fuel card statement xlsx", self._mobile_config("fuel_card_statement"), "XLSX"),
            ("mobile fuel card statement csv", self._mobile_config("fuel_card_statement"), "CSV"),
            ("mobile telematics fuel xlsx", self._mobile_config("telematics_fuel"), "XLSX"),
            ("mobile telematics fuel csv", self._mobile_config("telematics_fuel"), "CSV"),
            ("mobile telematics trips xlsx", self._mobile_config("telematics_trips"), "XLSX"),
            ("mobile telematics trips csv", self._mobile_config("telematics_trips"), "CSV"),
            ("mobile telematics odometer xlsx", self._mobile_config("telematics_odometer"), "XLSX"),
            ("mobile telematics odometer csv", self._mobile_config("telematics_odometer"), "CSV"),
        ]

        for label, raw_config, output_format in cases:
            with self.subTest(case=label):
                payload = generate_document_bytes(raw_config, output_format)
                self.assertTrue(payload)
                self._assert_payload_shape(payload, output_format)
                self.assertTrue(
                    build_document_download_filename(raw_config, output_format).endswith(self._expected_extension(output_format))
                )

    def test_monthly_zip_exports_for_bill_documents(self) -> None:
        cases = [
            (
                "heat",
                self._with_monthly_zip(self._heat_config("utility_bill"), start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026"),
                "utility_bill",
                "heat",
            ),
            (
                "electricity",
                self._with_monthly_zip(self._electricity_config("electricity_bill"), start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026"),
                "electricity_bill",
                "electricity",
            ),
        ]

        for category_label, raw_config, document_type, filename_prefix in cases:
            for output_format in ("PDF", "DOCX"):
                with self.subTest(category=category_label, output_format=output_format):
                    filename = build_document_download_filename(raw_config, output_format)
                    self.assertTrue(filename.endswith(".zip"))

                    archive_payload = generate_document_bytes(raw_config, output_format)
                    with ZipFile(BytesIO(archive_payload)) as archive:
                        names = sorted(archive.namelist())
                        expected_names = [
                            f"{filename_prefix}_{document_type}_2026-01.{output_format.lower()}",
                            f"{filename_prefix}_{document_type}_2026-02.{output_format.lower()}",
                            f"{filename_prefix}_{document_type}_2026-03.{output_format.lower()}",
                        ]
                        self.assertEqual(expected_names, names)

                        for name in names:
                            self._assert_payload_shape(archive.read(name), output_format)

    def test_ground_truth_counts_track_resolved_generation_shapes(self) -> None:
        heat_config = self._heat_config("supplier_portal_data")
        electricity_portal_config = self._electricity_config("supplier_portal_data")
        electricity_bill_config = self._electricity_config("electricity_bill")
        smart_meter_config = self._electricity_config("smart_meter_data", smart_meter_data_granularity="monthly")

        _, heat_sections, _ = _prepare_heat_sections(heat_config, output_format="CSV")
        heat_ground_truth = json.loads(generate_json_ground_truth(heat_config).decode("utf-8"))
        self.assertEqual(sum(len(section["records"]) for section in heat_sections), len(heat_ground_truth))
        self.assertTrue(all("company_label" in row and "site_label" in row for row in heat_ground_truth))

        _, electricity_sections = _prepare_electricity_sections(electricity_portal_config, output_format="CSV")
        electricity_ground_truth = json.loads(generate_json_ground_truth(electricity_portal_config).decode("utf-8"))
        self.assertEqual(len(electricity_sections), len(electricity_ground_truth))
        self.assertTrue(all("company_label" in row and "site_label" in row for row in electricity_ground_truth))

        _, electricity_bill_sections = _prepare_electricity_sections(electricity_bill_config, output_format="PDF")
        electricity_bill_ground_truth = json.loads(generate_json_ground_truth(electricity_bill_config).decode("utf-8"))
        self.assertEqual(len(electricity_bill_sections), len(electricity_bill_ground_truth))

        smart_meter_runtime_config, smart_meter_sections = _build_electricity_config(smart_meter_config, output_format="CSV")
        smart_meter_rows = build_smart_meter_rows(smart_meter_runtime_config, smart_meter_sections)
        smart_meter_ground_truth = json.loads(generate_json_ground_truth(smart_meter_config).decode("utf-8"))
        self.assertEqual(len(smart_meter_rows), len(smart_meter_ground_truth))

    def test_mobile_ground_truth_is_non_empty_and_category_tagged(self) -> None:
        for document_type in [
            "fuel_invoice",
            "fuel_card_statement",
            "telematics_fuel",
            "telematics_trips",
            "telematics_odometer",
        ]:
            with self.subTest(document_type=document_type):
                ground_truth = json.loads(
                    generate_json_ground_truth(self._mobile_config(document_type)).decode("utf-8")
                )
                self.assertTrue(ground_truth)
                self.assertTrue(all(row["Scope"] == "Scope 1" for row in ground_truth))
                self.assertTrue(all(row["Category"] == "mobile_combustion" for row in ground_truth))

    def test_cross_scope_line_items_are_tagged_with_true_category(self) -> None:
        mobile_config = self._mobile_config("fuel_invoice")
        mobile_config["document"]["cross_scope_items"] = True
        mobile_ground_truth = json.loads(generate_json_ground_truth(mobile_config).decode("utf-8"))
        self.assertEqual(
            {"mobile_combustion", "stationary_combustion"},
            {row["Category"] for row in mobile_ground_truth},
        )

        stationary_config = self._stationary_config("fuel_invoice")
        stationary_config["document"]["cross_scope_items"] = True
        stationary_ground_truth = json.loads(generate_json_ground_truth(stationary_config).decode("utf-8"))
        self.assertEqual(
            {"mobile_combustion", "stationary_combustion"},
            {row["Category"] for row in stationary_ground_truth},
        )
        payload = generate_document_bytes(stationary_config, "PDF")
        self.assertTrue(payload.startswith(b"%PDF-"))

    def test_mobile_multi_document_invoices_export_as_zip(self) -> None:
        raw_config = self._mobile_config("fuel_invoice")
        raw_config["document"]["doc_count"] = 3
        filename = build_document_download_filename(raw_config, "PDF")
        self.assertTrue(filename.endswith(".zip"))

        payload = generate_document_bytes(raw_config, "PDF")
        with ZipFile(BytesIO(payload)) as archive:
            names = archive.namelist()
            self.assertEqual(len(names), 3)
            for name in names:
                self.assertTrue(name.endswith(".pdf"))
                self._assert_payload_shape(archive.read(name), "PDF")

    def test_stationary_ground_truth_is_non_empty_for_each_document_family(self) -> None:
        cases = [
            self._stationary_config("fuel_invoice"),
            self._stationary_config("delivery_note"),
            self._stationary_config("fuel_card"),
            self._stationary_config("generator_log"),
            self._stationary_config("bems", bems_report_type="equipment_trend_report"),
            self._stationary_config("bems", bems_report_type="time_series_trend_export"),
        ]

        for raw_config in cases:
            with self.subTest(document_type=raw_config["document_type"], report_type=raw_config["document"].get("bems_report_type")):
                ground_truth = json.loads(generate_json_ground_truth(raw_config).decode("utf-8"))
                self.assertTrue(ground_truth)
                self.assertIsInstance(ground_truth, list)

    def _assert_payload_shape(self, payload: bytes, output_format: str) -> None:
        if output_format == "PDF":
            self.assertTrue(payload.startswith(b"%PDF-"))
            return
        if output_format == "CSV":
            text = payload.decode("utf-8-sig")
            self.assertIn("\n", text)
            return

        self.assertTrue(payload.startswith(b"PK"))
        with ZipFile(BytesIO(payload)) as archive:
            xml_entries = [name for name in archive.namelist() if name.endswith(".xml")]
            self.assertTrue(xml_entries)

    def _expected_extension(self, output_format: str) -> str:
        return f".{output_format.lower()}"

    def _with_monthly_zip(self, config: dict, *, start_date: str, end_date: str, label: str) -> dict:
        config = copy.deepcopy(config)
        config["financial_period"] = {
            "label": label,
            "start_date": start_date,
            "end_date": end_date,
        }
        config.setdefault("document", {})["monthly_zip"] = True
        return config

    def _heat_config(self, document_type: str) -> dict:
        return {
            "random_seed": 5,
            "financial_period": {
                "label": "Jan 2026",
                "start_date": "2026-01-01",
                "end_date": "2026-01-31",
            },
            "document_type": document_type,
            "document": {
                "language": "en",
                "distractor_fields": {"enabled": False},
            },
            "companies": [
                {
                    "label": "North Heat",
                    "supplier": "HeatCo",
                    "supplier_code": "HCO",
                    "supplier_address": ["1 Heat Way", "London"],
                    "customer": "Acme",
                    "customer_code": "ACM",
                    "currency": "GBP (£)",
                    "sites": [
                        {
                            "label": "HQ",
                            "customer_address": ["1 Main St", "London"],
                            "city": "London",
                            "postcode": "SW1A 1AA",
                            "meter_id": "MTR-001",
                            "capacity_kw": 150,
                            "capacity_rate": "3.2",
                            "supplier_ef": "0.21",
                            "base_consumption": 12000,
                            "unit_price_base": "0.065",
                            "start_reading": 1000,
                        }
                    ],
                }
            ],
        }

    def _electricity_config(
        self,
        document_type: str,
        *,
        smart_meter_data_granularity: str | None = None,
    ) -> dict:
        document = {
            "language": "en",
            "distractor_fields": {"enabled": False},
        }
        if smart_meter_data_granularity is not None:
            document["smart_meter_data_granularity"] = smart_meter_data_granularity
        return {
            "_category": "electricity",
            "random_seed": 5,
            "financial_period": {
                "label": "Jan 2026",
                "start_date": "2026-01-01",
                "end_date": "2026-01-31",
            },
            "document_type": document_type,
            "document": document,
            "companies": [
                {
                    "label": "Grid Power",
                    "supplier": "GridCo",
                    "supplier_code": "GCO",
                    "supplier_address": ["2 Grid Way", "London"],
                    "customer": "Acme",
                    "customer_code": "ACM",
                    "currency": "GBP (£)",
                    "sites": [
                        {
                            "label": "HQ",
                            "customer_address": ["1 Main St", "London"],
                            "city": "London",
                            "postcode": "SW1A 1AA",
                            "meter_id": "ELEC-001",
                            "supplier_ef": "0.18",
                            "unit": "kWh",
                            "start_reading": 5000,
                            "total_quantity": "8200",
                            "total_cost": "1640",
                            "tariffs": [{"name": "Day"}],
                        }
                    ],
                }
            ],
        }

    def _mobile_config(self, document_type: str) -> dict:
        return {
            "_category": "mobile_combustion",
            "random_seed": 5,
            "financial_period": {
                "label": "Jan 2026",
                "start_date": "2026-01-01",
                "end_date": "2026-01-31",
            },
            "document_type": document_type,
            "document": {
                "language": "en",
                "distractor_fields": {"enabled": False},
                "vat_rate": 20,
                "distance_unit": "km",
                "trips_per_vehicle": 3,
                "doc_count": 4 if document_type == "fuel_card_statement" else 1,
            },
            "companies": [
                {
                    "label": "Fleet Ops",
                    "supplier": "Roadside Fuel Supplies",
                    "supplier_code": "RFS",
                    "supplier_address": ["4 Carriage Way", "Leeds"],
                    "customer": "Acme Fleet",
                    "customer_code": "ACF",
                    "currency": "GBP (£)",
                    "card_number": "CARD-9",
                    "merchants": ["Applegreen M1"],
                    "vehicles": [
                        {
                            "registration": "AB12 CDE",
                            "make_model": "Ford Transit 350",
                            "vehicle_type": "Van (LGV)",
                            "fuel": "Diesel",
                            "unit": "L",
                            "driver": "J. Smith",
                            "card_number": "7002 3411 XXXX 9012",
                            "site": "London Depot",
                            "quantity": "55",
                            "unit_price": "1.52",
                            "monthly_distance_km": "2600",
                            "efficiency_l_per_100km": "9.8",
                            "odometer_start": "42000",
                        }
                    ],
                }
            ],
        }

    def _stationary_config(
        self,
        document_type: str,
        *,
        bems_report_type: str = "equipment_trend_report",
    ) -> dict:
        return {
            "_category": "stationary_combustion",
            "random_seed": 5,
            "financial_period": {
                "label": "Jan 2026",
                "start_date": "2026-01-01",
                "end_date": "2026-01-31",
            },
            "document_type": document_type,
            "document": {
                "language": "en",
                "distractor_fields": {"enabled": False},
                "bems_report_type": bems_report_type,
                "bems_interval_minutes": 60,
            },
            "companies": [
                {
                    "label": "Depot Ops",
                    "customer": "Acme Ops",
                    "supplier": "FuelCo",
                    "supplier_code": "FCO",
                    "customer_code": "AOP",
                    "supplier_address": ["3 Fuel Rd", "London"],
                    "currency": "GBP (£)",
                    "card_number": "CARD-1",
                    "merchant": "FuelCo Central",
                    "sites": [
                        {
                            "label": "Depot A",
                            "country": "UK",
                            "customer_address": ["Dock Road", "London"],
                            "merchant": "FuelCo Central",
                            "equipment_items": [
                                {
                                    "equipment": "Generator 1",
                                    "emission_source": "Backup Generator",
                                    "fuel": "Diesel",
                                    "quantity": "120",
                                    "unit": "L",
                                    "unit_price": "1.50",
                                    "delivery_charge": "25",
                                    "vat_rate": "20",
                                    "runs_per_month": 1,
                                    "fuel_used_per_hour": 10,
                                    "tank_capacity": 600,
                                    "run_hours_min": 1.0,
                                    "run_hours_max": 1.0,
                                }
                            ],
                            "assets": [
                                {
                                    "asset_tag": "AST-001",
                                    "equipment_name": "Boiler A",
                                    "emission_source": "Boiler",
                                    "fuel": "Gas Oil",
                                    "unit": "kWh",
                                    "sensor_name": "S-100",
                                    "quantity": "240",
                                    "operating_hours": "18",
                                }
                            ],
                        }
                    ],
                }
            ],
        }


if __name__ == "__main__":
    unittest.main()