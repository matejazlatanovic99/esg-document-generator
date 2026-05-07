from __future__ import annotations

import copy
import unittest
from io import BytesIO
from zipfile import ZipFile

from utils.distractor_fields import resolve_distractor_plan
from utils.generator import generate_document_bytes, generate_json_ground_truth


HEAT_LABELS = {"Connection ID", "Service Route", "Network Zone"}
ELECTRICITY_LABELS = {"Settlement Run", "Feeder Code", "Profile Class", "Meter Point Group"}
LOGISTICS_LABELS = {"Dispatch Ref", "Haulier Code", "Fleet Group", "Cost Centre"}
ADMIN_LABELS = {"Shift Code", "Plant Group", "Source System", "Report Batch"}


class DistractorFieldTests(unittest.TestCase):
    maxDiff = None

    def test_resolver_determinism_and_artifact_variation(self) -> None:
        disabled_plan = resolve_distractor_plan(
            {"language": "en", "distractor_fields": {"enabled": False}},
            random_seed=7,
            category="electricity",
            document_type="supplier_portal_data",
            output_format="CSV",
            context={"language": "en"},
        )
        self.assertFalse(disabled_plan.enabled)
        self.assertEqual(disabled_plan.document_fields, ())
        self.assertEqual(disabled_plan.tabular_fields, ())

        enabled_cfg = {"language": "en", "distractor_fields": {"enabled": True}}
        baseline = resolve_distractor_plan(
            enabled_cfg,
            random_seed=7,
            category="electricity",
            document_type="supplier_portal_data",
            output_format="CSV",
            artifact_key="root",
            context={"language": "en"},
        )
        repeat = resolve_distractor_plan(
            enabled_cfg,
            random_seed=7,
            category="electricity",
            document_type="supplier_portal_data",
            output_format="CSV",
            artifact_key="root",
            context={"language": "en"},
        )
        shifted = resolve_distractor_plan(
            enabled_cfg,
            random_seed=7,
            category="electricity",
            document_type="supplier_portal_data",
            output_format="CSV",
            artifact_key="2026-01",
            context={"language": "en"},
        )

        self.assertEqual(baseline, repeat)
        self.assertNotEqual(baseline, shifted)

    def test_heat_outputs_and_ground_truth(self) -> None:
        supplier_portal_config = self._heat_config("supplier_portal_data")
        utility_bill_config = self._heat_config("utility_bill")

        for output_format in ("CSV", "XLSX"):
            with self.subTest(document_type="supplier_portal_data", output_format=output_format):
                payload = generate_document_bytes(supplier_portal_config, output_format)
                self._assert_payload_contains_any(payload, output_format, HEAT_LABELS)

        docx_payload = generate_document_bytes(utility_bill_config, "DOCX")
        self._assert_payload_contains_any(docx_payload, "DOCX", HEAT_LABELS)

        pdf_payload = generate_document_bytes(utility_bill_config, "PDF")
        self.assertTrue(pdf_payload.startswith(b"%PDF-"))

        disabled_csv = generate_document_bytes(self._heat_config("supplier_portal_data", distractors_enabled=False), "CSV")
        self._assert_payload_excludes_all(disabled_csv, "CSV", HEAT_LABELS)

        ground_truth = generate_json_ground_truth(supplier_portal_config).decode("utf-8")
        tokens = self._plan_tokens(
            resolve_distractor_plan(
                supplier_portal_config["document"],
                random_seed=supplier_portal_config["random_seed"],
                category="heat",
                document_type="supplier_portal_data",
                output_format="CSV",
                context={"language": "en"},
            ),
            resolve_distractor_plan(
                utility_bill_config["document"],
                random_seed=utility_bill_config["random_seed"],
                category="heat",
                document_type="utility_bill",
                output_format="DOCX",
                context={"language": "en"},
            ),
        )
        self._assert_text_excludes_all(ground_truth, tokens)

    def test_electricity_outputs_and_ground_truth(self) -> None:
        supplier_portal_config = self._electricity_config("supplier_portal_data")
        bill_config = self._electricity_config("electricity_bill")
        smart_meter_config = self._electricity_config("smart_meter_data", smart_meter_data_granularity="monthly")

        for output_format in ("CSV", "XLSX"):
            with self.subTest(document_type="supplier_portal_data", output_format=output_format):
                payload = generate_document_bytes(supplier_portal_config, output_format)
                self._assert_payload_contains_any(payload, output_format, ELECTRICITY_LABELS)

        for output_format in ("CSV", "XLSX"):
            with self.subTest(document_type="smart_meter_data", output_format=output_format):
                payload = generate_document_bytes(smart_meter_config, output_format)
                self._assert_payload_contains_any(payload, output_format, ELECTRICITY_LABELS)

        docx_payload = generate_document_bytes(bill_config, "DOCX")
        self._assert_payload_contains_any(docx_payload, "DOCX", ELECTRICITY_LABELS)

        pdf_payload = generate_document_bytes(bill_config, "PDF")
        self.assertTrue(pdf_payload.startswith(b"%PDF-"))

        ground_truth = generate_json_ground_truth(supplier_portal_config).decode("utf-8")
        tokens = self._plan_tokens(
            resolve_distractor_plan(
                supplier_portal_config["document"],
                random_seed=supplier_portal_config["random_seed"],
                category="electricity",
                document_type="supplier_portal_data",
                output_format="CSV",
                context={"language": "en"},
            ),
            resolve_distractor_plan(
                smart_meter_config["document"],
                random_seed=smart_meter_config["random_seed"],
                category="electricity",
                document_type="smart_meter_data",
                output_format="CSV",
                context={"language": "en"},
            ),
            resolve_distractor_plan(
                bill_config["document"],
                random_seed=bill_config["random_seed"],
                category="electricity",
                document_type="electricity_bill",
                output_format="DOCX",
                context={"language": "en"},
            ),
        )
        self._assert_text_excludes_all(ground_truth, tokens)

    def test_site_labels_can_be_omitted_from_generated_outputs(self) -> None:
        heat_csv_config = self._omit_site_labels(self._heat_config("supplier_portal_data"))
        heat_docx_config = self._omit_site_labels(self._heat_config("utility_bill"))
        electricity_csv_config = self._omit_site_labels(self._electricity_config("supplier_portal_data"))
        electricity_docx_config = self._omit_site_labels(self._electricity_config("electricity_bill"))
        smart_meter_csv_config = self._omit_site_labels(
            self._electricity_config("smart_meter_data", smart_meter_data_granularity="monthly")
        )

        self._assert_text_excludes_all(
            self._payload_text(generate_document_bytes(heat_csv_config, "CSV"), "CSV"),
            {"HQ", "Site"},
        )
        self._assert_text_excludes_all(
            self._payload_text(generate_document_bytes(heat_docx_config, "DOCX"), "DOCX"),
            {"HQ"},
        )
        self._assert_text_excludes_all(
            self._payload_text(generate_document_bytes(electricity_csv_config, "CSV"), "CSV"),
            {"HQ", "Site"},
        )
        self._assert_text_excludes_all(
            self._payload_text(generate_document_bytes(electricity_docx_config, "DOCX"), "DOCX"),
            {"HQ"},
        )
        self._assert_text_excludes_all(
            self._payload_text(generate_document_bytes(smart_meter_csv_config, "CSV"), "CSV"),
            {"HQ", "Site"},
        )
        self._assert_text_excludes_all(
            self._payload_text(generate_document_bytes(smart_meter_csv_config, "XLSX"), "XLSX"),
            {"HQ", "Site"},
        )

    def test_stationary_outputs_and_ground_truth(self) -> None:
        for document_type in ("fuel_invoice", "delivery_note"):
            config = self._stationary_config(document_type)
            with self.subTest(document_type=document_type, output_format="DOCX"):
                docx_payload = generate_document_bytes(config, "DOCX")
                self._assert_payload_contains_any(docx_payload, "DOCX", LOGISTICS_LABELS)
            with self.subTest(document_type=document_type, output_format="PDF"):
                pdf_payload = generate_document_bytes(config, "PDF")
                self.assertTrue(pdf_payload.startswith(b"%PDF-"))

        fuel_card_config = self._stationary_config("fuel_card")
        for output_format in ("CSV", "XLSX", "DOCX"):
            with self.subTest(document_type="fuel_card", output_format=output_format):
                payload = generate_document_bytes(fuel_card_config, output_format)
                self._assert_payload_contains_any(payload, output_format, LOGISTICS_LABELS)
        self.assertTrue(generate_document_bytes(fuel_card_config, "PDF").startswith(b"%PDF-"))

        generator_log_config = self._stationary_config("generator_log")
        for output_format in ("CSV", "XLSX"):
            with self.subTest(document_type="generator_log", output_format=output_format):
                payload = generate_document_bytes(generator_log_config, output_format)
                self._assert_payload_contains_any(payload, output_format, ADMIN_LABELS)

        bems_equipment_config = self._stationary_config("bems", bems_report_type="equipment_trend_report")
        with self.subTest(document_type="bems_equipment", output_format="DOCX"):
            docx_payload = generate_document_bytes(bems_equipment_config, "DOCX")
            self._assert_payload_contains_any(docx_payload, "DOCX", ADMIN_LABELS)
        with self.subTest(document_type="bems_equipment", output_format="PDF"):
            self.assertTrue(generate_document_bytes(bems_equipment_config, "PDF").startswith(b"%PDF-"))

        bems_time_series_config = self._stationary_config("bems", bems_report_type="time_series_trend_export")
        for output_format in ("CSV", "XLSX"):
            with self.subTest(document_type="bems_time_series", output_format=output_format):
                payload = generate_document_bytes(bems_time_series_config, output_format)
                self._assert_payload_contains_any(payload, output_format, ADMIN_LABELS)

        logistics_ground_truth = generate_json_ground_truth(fuel_card_config).decode("utf-8")
        logistics_tokens = self._plan_tokens(
            resolve_distractor_plan(
                fuel_card_config["document"],
                random_seed=fuel_card_config["random_seed"],
                category="stationary_combustion",
                document_type="fuel_card",
                output_format="CSV",
                context={"language": "en", "bems_report_type": "equipment_trend_report"},
            ),
            resolve_distractor_plan(
                self._stationary_config("fuel_invoice")["document"],
                random_seed=fuel_card_config["random_seed"],
                category="stationary_combustion",
                document_type="fuel_invoice",
                output_format="DOCX",
                context={"language": "en", "bems_report_type": "equipment_trend_report"},
            ),
        )
        self._assert_text_excludes_all(logistics_ground_truth, logistics_tokens)

        bems_ground_truth = generate_json_ground_truth(bems_time_series_config).decode("utf-8")
        bems_tokens = self._plan_tokens(
            resolve_distractor_plan(
                bems_equipment_config["document"],
                random_seed=bems_equipment_config["random_seed"],
                category="stationary_combustion",
                document_type="bems",
                output_format="DOCX",
                context={"language": "en", "bems_report_type": "equipment_trend_report"},
            ),
            resolve_distractor_plan(
                bems_time_series_config["document"],
                random_seed=bems_time_series_config["random_seed"],
                category="stationary_combustion",
                document_type="bems",
                output_format="CSV",
                context={"language": "en", "bems_report_type": "time_series_trend_export"},
            ),
        )
        self._assert_text_excludes_all(bems_ground_truth, bems_tokens)

    def _assert_payload_contains_any(self, payload: bytes, output_format: str, labels: set[str]) -> None:
        text = self._payload_text(payload, output_format)
        self.assertTrue(any(label in text for label in labels), f"Expected one of {sorted(labels)} in {output_format} payload")

    def _assert_payload_excludes_all(self, payload: bytes, output_format: str, labels: set[str]) -> None:
        self._assert_text_excludes_all(self._payload_text(payload, output_format), labels)

    def _assert_text_excludes_all(self, text: str, labels: set[str]) -> None:
        unexpected = sorted(label for label in labels if label and label in text)
        self.assertEqual(unexpected, [], f"Unexpected distractor content in payload: {unexpected}")

    def _payload_text(self, payload: bytes, output_format: str) -> str:
        if output_format == "CSV":
            return payload.decode("utf-8-sig")
        if output_format in {"DOCX", "XLSX"}:
            return self._zip_xml_text(payload)
        raise ValueError(f"Text extraction is only supported for CSV/DOCX/XLSX, not {output_format}")

    def _zip_xml_text(self, payload: bytes) -> str:
        with ZipFile(BytesIO(payload)) as archive:
            return "\n".join(
                archive.read(name).decode("utf-8", errors="ignore")
                for name in archive.namelist()
                if name.endswith(".xml")
            )

    def _plan_tokens(self, *plans) -> set[str]:
        tokens: set[str] = set()
        for plan in plans:
            tokens.update(field.label for field in plan.document_fields)
            tokens.update(field.value for field in plan.document_fields)
            for field in plan.tabular_fields:
                tokens.add(field.label)
                tokens.update(field.value_options)
        return {token for token in tokens if token}

    def _heat_config(self, document_type: str, *, distractors_enabled: bool = True) -> dict:
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
                "distractor_fields": {"enabled": distractors_enabled},
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
        distractors_enabled: bool = True,
        smart_meter_data_granularity: str | None = None,
    ) -> dict:
        document = {
            "language": "en",
            "distractor_fields": {"enabled": distractors_enabled},
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

    def _stationary_config(
        self,
        document_type: str,
        *,
        distractors_enabled: bool = True,
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
                "distractor_fields": {"enabled": distractors_enabled},
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

    def _omit_site_labels(self, config: dict) -> dict:
        config = copy.deepcopy(config)
        for company in config.get("companies", []):
            for site in company.get("sites", []):
                site.setdefault("_omit", {})["label"] = True
        return config


if __name__ == "__main__":
    unittest.main()
