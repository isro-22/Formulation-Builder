# Formulation Sheet Automation

Access the live Streamlit app here:

https://formulation-builder-djdu7ukbsdzvgnepoobqdq.streamlit.app/

This project supports two main workflows in a single Streamlit workspace:

1. generating a `Formulation Sheet` from structured input,
2. scraping legacy Excel templates into normalized JSON/XLSX data.

The calculation and Excel rendering engine lives in `formulation_generator.py`, while `app.py` is the main entry point for both workflows.

## Overview

This system is designed to replace a legacy Excel/VBA process that is difficult to maintain, especially around:

- dynamic layouts,
- premix formulas and nested BOM logic,
- approval blocks,
- normalized output that can be reused as structured model data.

The final output is a stable Excel workbook, along with JSON/XLSX model exports for validation, backup, and future integration.

## Key Features

- Streamlit UI for product metadata, phase metadata, materials, and approval data
- Formulation workbook generation from `templates/Template_Generate.xlsm`
- Downloadable formulation model in JSON and XLSX formats
- Legacy Excel template scraping from `template_1.xlsx` or uploaded files
- Scraped output export to `outputs/scraped` as JSON and XLSX
- Header, phase, and material validation before generation
- Support for `mg/stick` and `%` input modes
- Dynamic material tables with preserved borders when phase rows expand
- Premix sections render only when material rows exist for that premix phase
- One visible blank separator row between rendered phase tables
- Online material lookup formulas use Excel-compatible `IFNA(XLOOKUP(...))` syntax
- Unit tests for the generator engine

## Project Structure

```text
03 Formulation Sheet Creation/
├── app.py
├── formulation_generator.py
├── template_scraper.py
├── requirements.txt
├── README.md
├── templates/
│   └── Template_Generate.xlsm
├── references/
│   ├── 02._SAAT_777.xlsx
│   └── BOL-SAAT List.xlsx
├── outputs/
│   ├── generated/
│   ├── scraped/
│   └── legacy/
├── assets/
└── tests/
```

## Main Files

### `app.py`

The main Streamlit UI. Users can choose between two workspaces:

- `Formulation Sheet Generator`
- `Data Extraction Engine`

### `formulation_generator.py`

The core engine for:

- formulation data models,
- input validation and transformation,
- dosage, ratio, premix, and price calculations,
- output workbook rendering,
- JSON/XLSX formulation model export,
- blank `INPUT` template generation.

### `template_scraper.py`

Reads legacy Excel templates and extracts:

- product headers,
- phase metadata,
- material rows,
- summaries,
- approval blocks.

This module can run independently or be rendered from `app.py`.

## Setup

### Requirements

- Python 3.8+
- terminal or shell access

### Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Running the App

Use the hosted version:

https://formulation-builder-djdu7ukbsdzvgnepoobqdq.streamlit.app/

Or run it locally:

```bash
streamlit run app.py
```

By default, Streamlit opens on localhost, usually:

```text
http://localhost:8501
```

## Workflow 1: Formulation Sheet Generator

1. Open the app.
2. Select `Formulation Sheet Generator`.
3. Upload an `INPUT` Excel file if you want to load an existing draft.
4. Complete the product metadata, approval data, phase metadata, and material table.
5. Click `Save Data / Validate Draft` to review issues.
6. Click `Generate Formulation Sheet` to create the final workbook.

Generated files are saved to:

- `outputs/generated`

## Workflow 2: Data Extraction Engine

1. Open the app.
2. Select `Data Extraction Engine`.
3. Use the default `template_1.xlsx` file or upload a legacy template.
4. Select the sheet to read.
5. Review the extracted data.
6. Download the output or use the automatically saved files.

Scraped files are saved to:

- `outputs/scraped`

## Reference Data

Before generating a workbook, users can choose one of two material lookup sources.

### Local Data

Material lookup and chemical prices are read from:

`references/BOL-SAAT List.xlsx`

The generator uses this file as the material master reference database.

### Online Data

The generated workbook writes SharePoint `XLOOKUP` formulas directly into the material table:

- `Physical State` looks up the online `Physical Form` value.
- `CAS Number` looks up the online CAS value.
- `Material Price (USD / KG)` looks up the online price value.

The formulas are written as `=IFNA(XLOOKUP(...), fallback)` without the implicit-intersection `@` operator. This avoids Excel parsing issues such as `=@IFNA(...)` in generated material lookup formulas.

### Layout Rules

Generated formulation sheets apply these layout rules:

- material table borders are applied across all active material rows for every phase,
- empty premix sections are removed instead of displayed as blank tables,
- rendered phase tables are separated by one visible blank row,
- sensory metadata cells `J13:J16` are formatted as numeric cells when numeric values are provided.

These rules apply to the main phases and supported premix phases:

- `Casing Rajangan`
- `Casing Krosok`
- `Top Flavor`
- `Casing Pre-Mix`
- `Casing Pre-Mix 2`
- `Casing Pre-Mix 3`
- `Flavor Pre-Mix 1`
- `Flavor Pre-Mix 2`
- `Flavor Pre-Mix 3`
- `Flavor Pre-Mix 4`
- `Flavor Pre-Mix 5`

## Database Model

The JSON/XLSX model export contains three main tables:

- `product`
- `phase_metadata`
- `materials`

Relationships are defined with these keys:

- `product.product_id` is the product primary key.
- `phase_metadata.product_id` references `product.product_id`.
- `phase_metadata.phase_id` is the phase primary key.
- `materials.product_id` references `product.product_id`.
- `materials.phase_id` references `phase_metadata.phase_id`.
- `materials.material_id` is the material row primary key.

With this structure, a database containing many products can be built by appending all `product`, `phase_metadata`, and `materials` rows from each exported model. To retrieve one product, filter by `product.product_id`, then load all phase and material rows with the same `product_id`.

## Testing

Run the generator test suite with:

```bash
python -m unittest tests/test_formulation_generator.py
```

## Technical Notes

- `app.py` combines the generator and scraper workflows.
- `template_scraper.py` is safe to import from `app.py`.
- `formulation_generator.py` remains the source of truth for workbook generation logic.
- Workbook output depends heavily on the Excel template layout, so template changes should be tested carefully.
- Change history and bug/action notes are tracked in `log.md`.

## Next Development

Recommended next areas:

- formulation versioning,
- database integration,
- PDF export,
- approval workflow,
- internal deployment hardening.
