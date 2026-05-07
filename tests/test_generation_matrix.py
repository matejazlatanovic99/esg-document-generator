from __future__ import annotations

import copy
import json
import unittest
from io import BytesIO
from zipfile import ZipFile

from generators.electricity_generator import build_smart_meter_rows
from utils.generator import (
    _build_electricity_config,
    _prepare_electricity_sections,
    _prepare_heat_sections,
    build_document_download_filename,
    generate_document_bytes,
    generate_json_ground_truth,
)


class GenerationMatrixTests(unittest.TestCase):
    def test_implemented_generation_matrix_returns_expected_payload_types(self) -> None:
        cases = [
            ("heat utility bill pdf", self._heat_config("utility_bill"), "PDF"),
            ("heat utility bill docx", self._heat_config("utility_bill"), "DOCX"),
            ("heat supplier portal csv", self._heat_config("supplier_portal_data"), "CSV"),
            ("heat supplier portal xlsx", self._heat_config("supplier_portal_data"), "XLSX"),
            ("electricity bill pdf", self._electricity_config("electricity_bill"), "PDF"),
            ("electricity bill docx", self._electricity_config("electricity_bill"), "DOCX"),
            ("electricity supplier portal csv", self._electricity_config("supplier_portal_data"), "CSV"),
            ("electricity supplier portal xlsx", self._electricity_config("supplier_portal_data"), "XLSX"),
            ("smart meter csv", self._electricity_config("smart_meter_data", smart_meter_data_granularity="monthly"), "CSV"),
            ("smart meter xlsx", self._electricity_config("smart_meter_data", smart_meter_data_granularity="monthly"), "XLSX"),
            ("stationary fuel invoice pdf", self._stationary_config("fuel_invoice"), "PDF"),
            ("stationary fuel invoice docx", self._stationary_config("fuel_invoice"), "DOCX"),
            ("stationary delivery note pdf", self._stationary_config("delivery_note"), "PDF"),
            ("stationary delivery note docx", self._stationary_config("delivery_note"), "DOCX"),
            ("stationary fuel card pdf", self._stationary_config("fuel_card"), "PDF"),
            ("stationary fuel card docx", self._stationary_config("fuel_card"), "DOCX"),
            ("stationary fuel card csv", self._stationary_config("fuel_card"), "CSV"),
            ("stationary fuel card xlsx", self._stationary_config("fuel_card"), "XLSX"),
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
        ]

        for label, raw_config, output_format in cases:
            with self.subTest(label=label):
                payload = generate_document_bytes(raw_config, output_format)
                self._assert_payload_type(payload, output_format)

    def test_monthly_zip_exports_use_expected_archive_names(self) -> None:
        cases = [
            (
                self._heat_config("utility_bill", start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026", monthly_zip=True),
                "PDF",
                "heat_utility_bill_2026-01_2026-03.zip",
                [
                    "heat_utility_bill_2026-01.pdf",
                    "heat_utility_bill_2026-02.pdf",
                    "heat_utility_bill_2026-03.pdf",
                ],
            ),
            (
                self._heat_config("utility_bill", start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026", monthly_zip=True),
                "DOCX",
                "heat_utility_bill_2026-01_2026-03.zip",
                [
                    "heat_utility_bill_2026-01.docx",
                    "heat_utility_bill_2026-02.docx",
                    "heat_utility_bill_2026-03.docx",
                ],
            ),
            (
                self._electricity_config("electricity_bill", start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026", monthly_zip=True),
                "PDF",
                "electricity_electricity_bill_2026-01_2026-03.zip",
                [
                    "electricity_electricity_bill_2026-01.pdf",
                    "electricity_electricity_bill_2026-02.pdf",
                    "electricity_electricity_bill_2026-03.pdf",
                ],
            ),
            (
                self._electricity_config("electricity_bill", start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026", monthly_zip=True),
                "DOCX",
                "electricity_electricity_bill_2026-01_2026-03.zip",
                [
                    "electricity_electricity_bill_2026-01.docx",
                    "electricity_electricity_bill_2026-02.docx",
                    "electricity_electricity_bill_2026-03.docx",
                ],
            ),
        ]

        for raw_config, output_format, expected_download_name, expected_members in cases:
            with self.subTest(output_format=output_format, expected_download_name=expected_download_name):
                self.assertEqual(
                    build_document_download_filename(raw_config, output_format),
                    expected_download_name,
                )
                payload = generate_document_bytes(raw_config, output_format)
                self.assertEqual(self._zip_member_names(payload), expected_members)

    def test_heat_ground_truth_row_count_matches_resolved_records(self) -> None:
        raw_config = self._heat_config("supplier_portal_data", start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026")
        _, sections, _ = _prepare_heat_sections(raw_config)
        expected_row_count = sum(len(section["records"]) for section in sections)

        rows = self._ground_truth_rows(raw_config)

        self.assertEqual(len(rows), expected_row_count)
        self.assertTrue(all("company_label" in row for row in rows))
        self.assertTrue(all("site_label" in row for row in rows))

    def test_electricity_ground_truth_row_count_matches_resolved_sections(self) -> None:
        raw_config = self._electricity_config("supplier_portal_data", start_date="2026-01-01", end_date="2026-03-31", label="Q1 2026")
        _, sections = _prepare_electricity_sections(raw_config)

        rows = self._ground_truth_rows(raw_config)

        self.assertEqual(len(rows), len(sections))
        self.assertTrue(all("company_label" in row for row in rows))
        self.assertTrue(all("site_label" in row for row in rows))

    def test_smart_meter_ground_truth_row_count_matches_generated_rows(self) -> None:
        raw_config = self._electricity_config(
            "smart_meter_data",
            start_date="2026-01-01",
            end_date="2026-03-31",
            label="Q1 2026",
            smart_meter_data_granularity="monthly",
        )
        config, sections = _build_electricity_config(raw_config)
        expected_rows = build_smart_meter_rows(config, sections)

        rows = self._ground_truth_rows(raw_config)

        self.assertEqual(len(rows), len(expected_rows))
        self.assertTrue(all("meter_id" in row for row in rows))

    def _assert_payload_type(self, payload: bytes, output_format: str) -> None:
        self.assertTrue(payload)
        if output_format == "PDF":
            self.assertTrue(payload.startswith(b"%PDF-"))
            return
        if output_format in {"DOCX", "XLSX"}:
            self.assertTrue(payload.startswith(b"PK"))
            return
        self.assertIn("\n", payload.decode("utf-8-sig"))

    def _zip_member_names(self, payload: bytes) -> list[str]:
        with ZipFile(BytesIO(payload)) as archive:
            return sorted(archive.namelist())

    def _ground_truth_rows(self, raw_config: dict) -> list[dict]:
        return json.loads(generate_json_ground_truth(raw_config).decode("utf-8"))

    def _heat_config(
        self,
        document_type: str,
        *,
        start_date: str = "2026-01-01",
        end_date: str = "2026-01-31",
        label: str = "Jan 2026",
        monthly_zip: bool = False,
    ) -> dict:
        return {
            "random_seed": 5,
            "financial_period": {
                "label": label,
                "start_date": start_date,
                "end_date": end_date,
            },
            "document_type": document_type,
            "document": {
                "language": "en",
                "monthly_zip": monthly_zip,
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
        start_date: str = "2026-01-01",
        end_date: str = "2026-01-31",
        label: str = "Jan 2026",
        monthly_zip: bool = False,
        smart_meter_data_granularity: str | None = None,
    ) -> dict:
        document = {
            "language": "en",
            "monthly_zip": monthly_zip,
            "distractor_fields": {"enabled": False},
        }
        if smart_meter_data_granularity is not None:
            document["smart_meter_data_granularity"] = smart_meter_data_granularity
        return {
            "_category": "electricity",
            "random_seed": 5,
            "financial_period": {
                "label": label,
                "start_date": start_date,
                "end_date": end_date,
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