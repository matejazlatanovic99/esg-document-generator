from __future__ import annotations


def build_raw_config(form_data: dict) -> dict:
    """Convert form_data (from scope_forms) into the generator-compatible config structure."""
    return {
        "random_seed": form_data["doc_seed"],
        "document": {
            "output_dir": "./outputs",  # overridden by temp dir in generator
            "background_dirname": "_page_backgrounds",
            "title": form_data["doc_title"],
            "subject": form_data["doc_subject"],
            "language": form_data.get("doc_language", "en"),
            "noise_level": float(form_data.get("doc_noise", 0.3)),
            "inject_special_chars": bool(form_data.get("doc_inject_special_chars", False)),
        },
        "financial_period": {
            "label": form_data["fp_label"],
            "start_date": form_data["fp_start"],
            "end_date": form_data["fp_end"],
        },
        "xlsx_split_by_company": form_data.get("xlsx_split_by_company", False),
        "companies": form_data["companies"],
    }


def build_raw_config_electricity(form_data: dict) -> dict:
    """Convert electricity form_data into the electricity-generator-compatible config dict."""
    return {
        "_category": "electricity",
        "random_seed": form_data["doc_seed"],
        "document": {
            "output_dir": "./outputs",
            "title": form_data["doc_title"],
            "subject": form_data["doc_subject"],
            "language": form_data.get("doc_language", "en"),
            "noise_level": float(form_data.get("doc_noise", 0.3)),
            "inject_special_chars": bool(form_data.get("doc_inject_special_chars", False)),
        },
        "financial_period": {
            "label": form_data["fp_label"],
            "start_date": form_data["fp_start"],
            "end_date": form_data["fp_end"],
        },
        "xlsx_split_by_company": form_data.get("xlsx_split_by_company", False),
        "companies": form_data["companies"],
    }


def validate_raw_config(raw_config: dict) -> list[str]:
    """Return a list of human-readable validation error strings (empty if valid)."""
    errors: list[str] = []

    fp = raw_config.get("financial_period", {})
    start_str = fp.get("start_date", "")
    end_str = fp.get("end_date", "")

    if start_str and end_str and start_str > end_str:
        errors.append("Financial period: end date must be after start date.")

    if not fp.get("label", "").strip():
        errors.append("Financial period label is required.")

    companies: list[dict] = raw_config.get("companies", [])
    if not companies:
        errors.append("At least one company is required.")

    seen_meter_ids: set[str] = set()

    for i, company in enumerate(companies):
        prefix = f"Company {i + 1} ({company.get('label', '?')})"
        co_omit: dict = company.get("_omit", {})

        for field, label in [
            ("label", "Company label"),
            ("supplier", "Supplier name"),
            ("supplier_code", "Supplier code"),
            ("customer", "Customer name"),
            ("customer_code", "Customer code"),
        ]:
            if not co_omit.get(field, False) and not company.get(field, "").strip():
                errors.append(f"{prefix}: {label} is required.")

        if not co_omit.get("supplier_address", False) and not company.get("supplier_address"):
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
                ("meter_id", "Heat Meter ID"),
            ]:
                if not site_omit.get(field, False) and not site.get(field, "").strip():
                    errors.append(f"{site_prefix}: {label} is required.")

            meter_id = site.get("meter_id", "").strip()
            if meter_id:
                if meter_id in seen_meter_ids:
                    errors.append(f"{site_prefix}: Meter ID '{meter_id}' is duplicated.")
                else:
                    seen_meter_ids.add(meter_id)

            if not site_omit.get("address", False) and not site.get("customer_address"):
                errors.append(f"{site_prefix}: Customer address is required.")

            billing_periods = site.get("billing_periods")
            if billing_periods is not None and len(billing_periods) == 0:
                errors.append(
                    f"{site_prefix}: At least one billing month must be selected "
                    "when using custom billing periods."
                )

    return errors


def validate_raw_config_electricity(raw_config: dict) -> list[str]:
    """Validation for electricity form configs."""
    errors: list[str] = []

    fp = raw_config.get("financial_period", {})
    if fp.get("start_date", "") > fp.get("end_date", ""):
        errors.append("Financial period: end date must be after start date.")
    if not fp.get("label", "").strip():
        errors.append("Financial period label is required.")

    companies: list[dict] = raw_config.get("companies", [])
    if not companies:
        errors.append("At least one company is required.")

    seen_meter_ids: set[str] = set()

    for i, company in enumerate(companies):
        prefix = f"Company {i + 1} ({company.get('label', '?')})"
        co_omit: dict = company.get("_omit", {})

        for field, label in [
            ("label", "Company label"),
            ("supplier", "Supplier name"),
            ("supplier_code", "Supplier code"),
            ("customer", "Customer name"),
            ("customer_code", "Customer code"),
        ]:
            if not co_omit.get(field, False) and not company.get(field, "").strip():
                errors.append(f"{prefix}: {label} is required.")

        if not co_omit.get("supplier_address", False) and not company.get("supplier_address"):
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
                ("meter_id", "Electricity Meter ID"),
            ]:
                if not site_omit.get(field, False) and not site.get(field, "").strip():
                    errors.append(f"{site_prefix}: {label} is required.")

            if not site_omit.get("address", False) and not site.get("customer_address"):
                errors.append(f"{site_prefix}: Customer address is required.")

            meter_id = site.get("meter_id", "").strip()
            if meter_id:
                if meter_id in seen_meter_ids:
                    errors.append(f"{site_prefix}: Meter ID '{meter_id}' is duplicated.")
                else:
                    seen_meter_ids.add(meter_id)

            try:
                qty = float(site.get("total_quantity", 0))
                if qty <= 0:
                    errors.append(f"{site_prefix}: Total quantity must be greater than zero.")
            except (ValueError, TypeError):
                errors.append(f"{site_prefix}: Total quantity must be a valid number.")

    return errors


def validate_raw_config(raw_config: dict) -> list[str]:
    """Return a list of human-readable validation error strings (empty if valid)."""
    errors: list[str] = []

    fp = raw_config.get("financial_period", {})
    start_str = fp.get("start_date", "")
    end_str = fp.get("end_date", "")

    if start_str and end_str and start_str > end_str:
        errors.append("Financial period: end date must be after start date.")

    if not fp.get("label", "").strip():
        errors.append("Financial period label is required.")

    companies: list[dict] = raw_config.get("companies", [])
    if not companies:
        errors.append("At least one company is required.")

    seen_meter_ids: set[str] = set()

    for i, company in enumerate(companies):
        prefix = f"Company {i + 1} ({company.get('label', '?')})"
        co_omit: dict = company.get("_omit", {})

        for field, label in [
            ("label", "Company label"),
            ("supplier", "Supplier name"),
            ("supplier_code", "Supplier code"),
            ("customer", "Customer name"),
            ("customer_code", "Customer code"),
        ]:
            if not co_omit.get(field, False) and not company.get(field, "").strip():
                errors.append(f"{prefix}: {label} is required.")

        if not co_omit.get("supplier_address", False) and not company.get("supplier_address"):
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
                ("meter_id", "Heat Meter ID"),
            ]:
                if not site_omit.get(field, False) and not site.get(field, "").strip():
                    errors.append(f"{site_prefix}: {label} is required.")

            meter_id = site.get("meter_id", "").strip()
            if meter_id:
                if meter_id in seen_meter_ids:
                    errors.append(f"{site_prefix}: Meter ID '{meter_id}' is duplicated.")
                else:
                    seen_meter_ids.add(meter_id)

            if not site_omit.get("address", False) and not site.get("customer_address"):
                errors.append(f"{site_prefix}: Customer address is required.")

            billing_periods = site.get("billing_periods")
            if billing_periods is not None and len(billing_periods) == 0:
                errors.append(
                    f"{site_prefix}: At least one billing month must be selected "
                    "when using custom billing periods."
                )

    return errors
