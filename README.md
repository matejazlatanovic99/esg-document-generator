# ESG Document Generator

A Streamlit web application for generating ESG (Environmental, Social, Governance) billing documents across all GHG Protocol scopes.

## Features

- **Scope-based navigation** — browse all GHG Protocol scopes and categories via the sidebar
- **Dynamic forms** — fields update based on the selected scope and category
- **Pre-filled defaults** — form loads with example data so you can generate immediately
- **Validation** — required fields, duplicate meter IDs, and billing period checks before generation
- **PDF download** — generated document is returned as a browser download, no files saved to disk

## Scope support

| Scope | Category | Status |
|---|---|---|
| Scope 2: Indirect Energy | Purchased Heat / Steam / Cooling | Implemented |
| Scope 1: Direct Emissions | Stationary Combustion | Coming Soon |
| Scope 1: Direct Emissions | Mobile Combustion | Coming Soon |
| Scope 1: Direct Emissions | Fugitive Emissions | Coming Soon |
| Scope 2: Indirect Energy | Electricity | Coming Soon |
| Scope 3: Upstream | All categories (8) | Coming Soon |
| Scope 3: Downstream | All categories (7) | Coming Soon |

## Output formats

| Format | Status |
|---|---|
| PDF | Implemented |
| XLSX | Implemented |
| DOCX | Coming Soon |
| CSV | Coming Soon |

## Project structure

```
doc-generator/
├── app.py                          # Main Streamlit app
├── components/
│   ├── sidebar.py                  # Scope / category / format selectors
│   └── scope_forms.py             # Dynamic form rendering and data collection
├── utils/
│   ├── config.py                   # Config builder and validator
│   └── generator.py               # Integration with the PDF generator
├── generators/
│   ├── pdf-generator.py            # PDF generation engine (ReportLab)
│   └── pdf-generator.config.json  # Example configuration
├── requirements.txt
├── DEPLOY.md                       # Deployment instructions
└── README.md
```

## Getting started

### Prerequisites

Python 3.10 or later.

### Install dependencies

```bash
pip install -r requirements.txt
```

### Run locally

```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Configuration reference

The Purchased Heat form maps to the following configuration schema:

### Document settings

| Field | Description | Default |
|---|---|---|
| Document Title | PDF metadata title | `District Heating Billing Statement` |
| Document Subject | PDF metadata subject | `Purchased Heat billing statements` |
| Output Filename | Downloaded file name | `billing_statement.pdf` |
| Random Seed | Controls meter reading and price variations | `20260325` |

### Financial period

| Field | Description |
|---|---|
| Period Label | Human-readable label printed on each invoice |
| Start Date | First day of the billing year |
| End Date | Last day of the billing year |

### Company fields

| Field | Description |
|---|---|
| Company Label | Internal identifier for this company |
| Supplier Name | Energy supplier printed on invoices |
| Supplier Code | Short code used in invoice numbers |
| Supplier Address | Supplier address, one line per row |
| Customer Name | Customer entity printed on invoices |
| Customer Code | Short code used in invoice numbers |
| Currency *(advanced)* | Currency label, e.g. `GBP (£)` |
| Accent Colour *(advanced)* | Hex colour used for invoice styling |

### Site fields

| Field | Description |
|---|---|
| Site Label | Name of the physical location |
| Customer Address | Site address, one line per row |
| City | City for the site |
| Postcode | Postcode / ZIP for the site |
| Heat Meter ID | Unique meter identifier printed on invoices |
| Contracted Capacity (kW) | Capacity charge basis, 50–500 kW |
| Capacity Rate (£/kW/month) | Rate per kW of contracted capacity |
| Base Monthly Consumption (kWh) | Average monthly consumption, used with seasonal factors |
| Base Unit Price (£/kWh) | Starting heat unit price before seasonal adjustments |
| Start Meter Reading (kWh) | Opening meter value for the first billing period |
| Billing Periods | Full financial period (auto) or custom month selection |

## How it works

1. The Streamlit UI collects form data and builds a config dict (`utils/config.py`)
2. The config is validated before generation
3. `utils/generator.py` loads the PDF engine via `importlib`, normalises the config, and runs generation inside a temporary directory
4. The PDF bytes are returned to the browser via `st.download_button` — nothing is written to disk permanently

## Deployment

See [DEPLOY.md](DEPLOY.md) for Streamlit Cloud and Docker instructions.
