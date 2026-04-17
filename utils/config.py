from __future__ import annotations

from collections.abc import Callable


def build_raw_config(
    form_data: dict,
    *,
    category: str | None = None,
    document_overrides: dict | None = None,
) -> dict:
    """Convert form data into the generator-compatible config structure."""
    raw_config = {
        "_category": category or form_data.get("_category"),
        "document_type": form_data.get("document_type", "utility_bill"),
        "random_seed": form_data["doc_seed"],
        "document": {
            "output_dir": "./outputs",
            "background_dirname": "_page_backgrounds",
            "type": form_data.get("document_type", "utility_bill"),
            "title": form_data["doc_title"],
            "subject": form_data["doc_subject"],
            "language": form_data.get("doc_language", "en"),
            "noise_level": float(form_data.get("doc_noise", 0.0)),
            "monthly_zip": bool(form_data.get("doc_monthly_zip", False)),
            "inject_special_chars": bool(form_data.get("doc_inject_special_chars", False)),
        },
        "xlsx_include_summary": form_data.get("xlsx_include_summary", False),
        "financial_period": {
            "label": form_data["fp_label"],
            "start_date": form_data["fp_start"],
            "end_date": form_data["fp_end"],
        },
        "xlsx_split_by_company": form_data.get("xlsx_split_by_company", False),
        "companies": form_data["companies"],
    }
    if document_overrides:
        raw_config["document"].update(document_overrides)
    return raw_config


def build_raw_config_electricity(form_data: dict) -> dict:
    """Convert electricity form data into the generator config dict."""
    return build_raw_config(
        form_data,
        category="electricity",
        document_overrides={
            "smart_meter_data_granularity": form_data.get("smart_meter_data_granularity", "monthly"),
            "smart_meter_interval_minutes": int(form_data.get("smart_meter_interval_minutes", 30)),
            "smart_meter_interval_value_mode": form_data.get("smart_meter_interval_value_mode", "consumption_diff"),
            "smart_meter_timestamp_format": form_data.get("smart_meter_timestamp_format", "iso_8601_utc"),
        },
    )


def build_raw_config_stationary(form_data: dict) -> dict:
    """Convert stationary combustion form data into the generator config dict."""
    return build_raw_config(
        form_data,
        category="stationary_combustion",
        document_overrides={
            "bems_interval_minutes": int(form_data.get("bems_interval_minutes", 60)),
            "bems_report_type": form_data.get("bems_report_type", "equipment_trend_report"),
        },
    )


def _validate_common_financial_period(raw_config: dict) -> list[str]:
    errors: list[str] = []
    fp = raw_config.get("financial_period", {})
    start_str = fp.get("start_date", "")
    end_str = fp.get("end_date", "")

    if start_str and end_str and start_str > end_str:
        errors.append("Financial period: end date must be after start date.")
    if not fp.get("label", "").strip():
        errors.append("Financial period label is required.")
    return errors


def _validate_metered_scope_config(
    raw_config: dict,
    *,
    meter_label: str,
    extra_site_validation: Callable[[str, dict, list[str]], None] | None = None,
) -> list[str]:
    errors = _validate_common_financial_period(raw_config)

    companies: list[dict] = raw_config.get("companies", [])
    if not companies:
        errors.append("At least one company is required.")

    seen_meter_ids: set[str] = set()

    for i, company in enumerate(companies):
        prefix = f"Company {i + 1} ({company.get('label', '?')})"

        for field, label in [
            ("label", "Company label"),
            ("supplier", "Supplier name"),
            ("supplier_code", "Supplier code"),
            ("customer", "Customer name"),
            ("customer_code", "Customer code"),
        ]:
            if not company.get(field, "").strip():
                errors.append(f"{prefix}: {label} is required.")

        if not company.get("supplier_address"):
            errors.append(f"{prefix}: Supplier address is required.")

        sites: list[dict] = company.get("sites", [])
        if not sites:
            errors.append(f"{prefix}: At least one site is required.")

        for j, site in enumerate(sites):
            site_prefix = f"{prefix} > Site {j + 1} ({site.get('label', '?')})"
            site_omit: dict = site.get("_omit", {})

            for field, label in [
                ("city", "City"),
                ("postcode", "Postcode"),
            ]:
                if not site_omit.get(field, False) and not site.get(field, "").strip():
                    errors.append(f"{site_prefix}: {label} is required.")

            if not site.get("meter_id", "").strip():
                errors.append(f"{site_prefix}: {meter_label} is required.")

            meter_id = site.get("meter_id", "").strip()
            if meter_id:
                if meter_id in seen_meter_ids:
                    errors.append(f"{site_prefix}: Meter ID '{meter_id}' is duplicated.")
                else:
                    seen_meter_ids.add(meter_id)

            if not site.get("customer_address"):
                errors.append(f"{site_prefix}: Customer address is required.")

            if extra_site_validation is not None:
                extra_site_validation(site_prefix, site, errors)

    return errors


def _validate_heat_site(site_prefix: str, site: dict, errors: list[str]) -> None:
    billing_periods = site.get("billing_periods")
    if billing_periods is not None and len(billing_periods) == 0:
        errors.append(
            f"{site_prefix}: At least one billing month must be selected "
            "when using custom billing periods."
        )

    for field, label in [
        ("capacity_kw", "Contracted capacity"),
        ("capacity_rate", "Capacity rate"),
        ("base_consumption", "Base monthly consumption"),
        ("unit_price_base", "Base unit price"),
    ]:
        try:
            if float(site.get(field, 0)) <= 0:
                errors.append(f"{site_prefix}: {label} must be greater than zero.")
        except (TypeError, ValueError):
            errors.append(f"{site_prefix}: {label} must be a valid number.")

    try:
        if float(site.get("start_reading", 0)) < 0:
            errors.append(f"{site_prefix}: Start meter reading cannot be negative.")
    except (TypeError, ValueError):
        errors.append(f"{site_prefix}: Start meter reading must be a valid number.")

    try:
        if float(site.get("supplier_ef", 0)) < 0:
            errors.append(f"{site_prefix}: Supplier emission factor cannot be negative.")
    except (TypeError, ValueError):
        errors.append(f"{site_prefix}: Supplier emission factor must be a valid number.")


def _validate_electricity_site(site_prefix: str, site: dict, errors: list[str]) -> None:
    try:
        qty = float(site.get("total_quantity", 0))
        if qty <= 0:
            errors.append(f"{site_prefix}: Total quantity must be greater than zero.")
    except (ValueError, TypeError):
        errors.append(f"{site_prefix}: Total quantity must be a valid number.")


def validate_raw_config(raw_config: dict) -> list[str]:
    """Validation for purchased heat configs."""
    return _validate_metered_scope_config(
        raw_config,
        meter_label="Heat Meter ID",
        extra_site_validation=_validate_heat_site,
    )


def validate_raw_config_electricity(raw_config: dict) -> list[str]:
    """Validation for electricity configs."""
    return _validate_metered_scope_config(
        raw_config,
        meter_label="Electricity Meter ID",
        extra_site_validation=_validate_electricity_site,
    )


def validate_raw_config_stationary(raw_config: dict) -> list[str]:
    """Validation for stationary combustion configs."""
    errors = _validate_common_financial_period(raw_config)
    document_type = raw_config.get("document_type", "fuel_invoice")

    companies: list[dict] = raw_config.get("companies", [])
    if not companies:
        errors.append("At least one company is required.")

    for i, company in enumerate(companies):
        prefix = f"Company {i + 1} ({company.get('label', '?')})"
        required_company_fields = [
            ("label", "Company label"),
            ("supplier", "Supplier name"),
            ("supplier_code", "Supplier code"),
            ("customer", "Customer name"),
        ]
        if document_type != "delivery_note":
            required_company_fields.append(("currency", "Currency"))

        for field, label in required_company_fields:
            if not str(company.get(field, "")).strip():
                errors.append(f"{prefix}: {label} is required.")

        if document_type not in {"fuel_card"} and not company.get("supplier_address"):
            errors.append(f"{prefix}: Supplier address is required.")

        sites: list[dict] = company.get("sites", [])
        if not sites:
            errors.append(f"{prefix}: At least one site is required.")

        for j, site in enumerate(sites):
            site_prefix = f"{prefix} > Site {j + 1} ({site.get('label', '?')})"
            if document_type not in {"fuel_card"} and not site.get("customer_address"):
                errors.append(f"{site_prefix}: Address is required.")

            if document_type == "fuel_invoice":
                site_omit: dict = site.get("_omit", {})
                for field, label in [
                    ("label", "Site"),
                    ("fuel", "Fuel"),
                    ("unit", "Unit"),
                ]:
                    if not str(site.get(field, "")).strip():
                        errors.append(f"{site_prefix}: {label} is required.")

                if not site_omit.get("country", False) and not str(site.get("country", "")).strip():
                    errors.append(f"{site_prefix}: Country is required unless omitted.")
                if not site_omit.get("equipment", False) and not str(site.get("equipment", "")).strip():
                    errors.append(f"{site_prefix}: Equipment is required unless omitted.")
                if not site_omit.get("emission_source", False) and not str(site.get("emission_source", "")).strip():
                    errors.append(f"{site_prefix}: Emission source is required unless omitted.")

                try:
                    if float(site.get("quantity", 0)) <= 0:
                        errors.append(f"{site_prefix}: Quantity must be greater than zero.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Quantity must be a valid number.")

                try:
                    if float(site.get("unit_price", 0)) <= 0:
                        errors.append(f"{site_prefix}: Unit price must be greater than zero.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Unit price must be a valid number.")
            elif document_type == "delivery_note":
                site_omit: dict = site.get("_omit", {})
                for field, label in [
                    ("label", "Site"),
                    ("fuel", "Fuel"),
                    ("unit", "Unit"),
                ]:
                    if not str(site.get(field, "")).strip():
                        errors.append(f"{site_prefix}: {label} is required.")

                if not site_omit.get("country", False) and not str(site.get("country", "")).strip():
                    errors.append(f"{site_prefix}: Country is required unless omitted.")
                if not site_omit.get("equipment", False) and not str(site.get("equipment", "")).strip():
                    errors.append(f"{site_prefix}: Tank / equipment is required unless omitted.")

                try:
                    if float(site.get("quantity", 0)) <= 0:
                        errors.append(f"{site_prefix}: Delivered quantity must be greater than zero.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Delivered quantity must be a valid number.")
            elif document_type == "fuel_card":
                site_omit: dict = site.get("_omit", {})

                for field, label in [
                    ("merchant", "Merchant"),
                    ("card_number", "Card number"),
                    ("fuel", "Fuel"),
                    ("unit", "Unit"),
                ]:
                    if not str(site.get(field, "")).strip():
                        errors.append(f"{site_prefix}: {label} is required.")

                if not site_omit.get("label", False) and not str(site.get("label", "")).strip():
                    errors.append(f"{site_prefix}: Site is required unless omitted.")
                if not site_omit.get("country", False) and not str(site.get("country", "")).strip():
                    errors.append(f"{site_prefix}: Country is required unless omitted.")
                if not site_omit.get("equipment", False) and not str(site.get("equipment", "")).strip():
                    errors.append(f"{site_prefix}: Equipment is required unless omitted.")
                if not site_omit.get("emission_source", False) and not str(site.get("emission_source", "")).strip():
                    errors.append(f"{site_prefix}: Emission source is required unless omitted.")

                try:
                    if float(site.get("quantity", 0)) <= 0:
                        errors.append(f"{site_prefix}: Quantity must be greater than zero.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Quantity must be a valid number.")

                try:
                    if float(site.get("unit_price", 0)) <= 0:
                        errors.append(f"{site_prefix}: Unit price must be greater than zero.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Unit price must be a valid number.")
            elif document_type == "bems":
                site_omit: dict = site.get("_omit", {})
                if not str(site.get("label", "")).strip():
                    errors.append(f"{site_prefix}: Site is required.")
                if not site_omit.get("country", False) and not str(site.get("country", "")).strip():
                    errors.append(f"{site_prefix}: Country is required.")

                assets: list[dict] = site.get("assets", [])
                if not assets:
                    errors.append(f"{site_prefix}: At least one asset is required.")

                seen_asset_tags: set[str] = set()
                for asset_idx, asset in enumerate(assets):
                    asset_prefix = f"{site_prefix} > Asset {asset_idx + 1} ({asset.get('asset_tag', '?')})"
                    asset_omit: dict = asset.get("_omit", {})
                    for field, label in [
                        ("asset_tag", "Asset tag"),
                        ("unit", "Unit"),
                    ]:
                        if not str(asset.get(field, "")).strip():
                            errors.append(f"{asset_prefix}: {label} is required.")

                    for field, label in [
                        ("equipment_name", "Equipment name"),
                        ("emission_source", "Emission source"),
                        ("fuel", "Fuel"),
                        ("sensor_name", "Sensor name"),
                    ]:
                        if not asset_omit.get(field, False) and not str(asset.get(field, "")).strip():
                            errors.append(f"{asset_prefix}: {label} is required.")

                    asset_tag = str(asset.get("asset_tag", "")).strip()
                    if asset_tag:
                        if asset_tag in seen_asset_tags:
                            errors.append(f"{asset_prefix}: Asset tag '{asset_tag}' is duplicated within the site.")
                        else:
                            seen_asset_tags.add(asset_tag)

                    try:
                        if float(asset.get("quantity", 0)) <= 0:
                            errors.append(f"{asset_prefix}: Consumption must be greater than zero.")
                    except (TypeError, ValueError):
                        errors.append(f"{asset_prefix}: Consumption must be a valid number.")

                    try:
                        if (
                            not asset_omit.get("operating_hours", False)
                            and float(asset.get("operating_hours", 0)) <= 0
                        ):
                            errors.append(f"{asset_prefix}: Operating hours must be greater than zero.")
                    except (TypeError, ValueError):
                        if not asset_omit.get("operating_hours", False):
                            errors.append(f"{asset_prefix}: Operating hours must be a valid number.")
            else:
                site_omit: dict = site.get("_omit", {})
                for field, label in [
                    ("label", "Site"),
                    ("equipment", "Equipment"),
                    ("fuel", "Fuel"),
                    ("unit", "Unit"),
                ]:
                    if not str(site.get(field, "")).strip():
                        errors.append(f"{site_prefix}: {label} is required.")

                if not site_omit.get("country", False) and not str(site.get("country", "")).strip():
                    errors.append(f"{site_prefix}: Country is required unless omitted.")
                if not site_omit.get("emission_source", False) and not str(site.get("emission_source", "")).strip():
                    errors.append(f"{site_prefix}: Emission source is required unless omitted.")

                try:
                    if int(site.get("runs_per_month", 0)) <= 0:
                        errors.append(f"{site_prefix}: Runs per month must be greater than zero.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Runs per month must be a whole number.")

                try:
                    if float(site.get("fuel_used_per_hour", 0)) <= 0:
                        errors.append(f"{site_prefix}: Fuel used per hour must be greater than zero.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Fuel used per hour must be a valid number.")

                try:
                    min_hours = float(site.get("run_hours_min", 0))
                    max_hours = float(site.get("run_hours_max", 0))
                    if min_hours <= 0 or max_hours <= 0 or min_hours > max_hours:
                        errors.append(f"{site_prefix}: Run hour bounds must be valid and increasing.")
                except (TypeError, ValueError):
                    errors.append(f"{site_prefix}: Run hour bounds must be valid numbers.")

                if site.get("quantity_mode") == "tank_level_change":
                    try:
                        if float(site.get("tank_capacity", 0)) <= 0:
                            errors.append(f"{site_prefix}: Tank capacity must be greater than zero.")
                    except (TypeError, ValueError):
                        errors.append(f"{site_prefix}: Tank capacity must be a valid number.")

    return errors
