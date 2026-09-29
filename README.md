# Inventory Reporting Automation

A Python-based Excel automation project for standardizing product
mappings, processing distributor inventory files, converting inventory
quantities into a common master-carton unit, and generating a
consolidated national inventory report.

## Overview

This repository contains two related automation scripts:

### 1. Product Dictionary Automation

`product_dictionary_automation_public.py`

This script helps maintain and expand a product dictionary when
different distributors use different product names or codes.

Key features: - Normalizes Persian/Arabic characters and digits -
Preserves existing valid mappings - Uses fuzzy matching with
`rapidfuzz` - Compares each unmatched product against known product
examples - Calculates match score and margin from the second-best
option - Assigns confidence statuses for manual review - Exports an
updated Excel dictionary

### 2. Inventory Reporting Automation

`inventory_reporting_automation_public.py`

This script processes multiple distributor inventory workbooks and
creates one standardized Excel report.

Key features: - Detects representative/distributor files by filename -
Reads `.xlsx` and `.xlsm` inventory files - Detects inventory tables
even when headers vary - Handles merged Excel headers - Matches
distributor product codes to a central dictionary - Supports three
inventory conversion levels - Converts quantities to master-carton
equivalents - Consolidates duplicate rows before final rounding - Avoids
guessing when mappings or structures are ambiguous - Generates a
hierarchical Excel report: `Brand → Product Group → Section → Flavor` -
Creates a dedicated error sheet for invalid or ambiguous records -
Supports Persian text and Jalali report dates

## Project Structure

``` text
inventory-reporting-automation/
├── product_dictionary_automation_public.py
├── inventory_reporting_automation_public.py
├── requirements.txt
├── .gitignore
├── inventories/
│   └── .gitkeep
└── README.md
```

> Real company data, internal dictionaries, distributor files, product
> aliases, and confidential business rules are intentionally excluded
> from the public repository.

## Requirements

-   Python 3.10+
-   pandas
-   numpy
-   openpyxl
-   rapidfuzz

Install dependencies:

``` bash
pip install -r requirements.txt
```

## Product Dictionary Workflow

Place these files beside `product_dictionary_automation_public.py`:

``` text
Dictionary_v2_matched.xlsx
نام کالا کد کالا.xlsx
```

Then run:

``` bash
python product_dictionary_automation_public.py
```

The default output is:

``` text
Dictionary_v3_matched_updated.xlsx
```

The script adds suggested internal product codes/names, matching scores,
score margins, and review statuses.

## Inventory Reporting Workflow

Place the dictionary file beside
`inventory_reporting_automation_public.py`.

Its filename should begin with:

``` text
Dictionary_v2_matched_updated
```

Create an `inventories` folder and place distributor Excel files inside
it:

``` text
inventories/
├── representative-a.xlsx
├── representative-b.xlsx
└── representative-c.xlsm
```

Run:

``` bash
python inventory_reporting_automation_public.py
```

The default output is:

``` text
inventory_report.xlsx
```

The workbook contains: - A consolidated inventory report - Hierarchical
product headers - Representative/company information - Total
quantities - An error sheet with detailed validation messages

## Inventory Conversion Logic

The reporting script supports three inventory levels:

-   **Level 1:** quantity is already expressed in the master-carton unit
-   **Level 2:** quantity is divided by the configured number of middle
    packs per master carton
-   **Level 3:** quantity is divided by the configured number of units
    per carton

Conversions are accumulated precisely before the final quantity is
rounded down.

## Error Handling

The project is intentionally conservative: when the input is ambiguous,
it records an error instead of silently guessing.

Examples include: - Unknown representative - Missing product code -
Product code not found in the dictionary - Ambiguous duplicate
dictionary definitions - Missing inventory columns - Invalid or negative
quantities - Invalid conversion factors - Inconsistent product metadata

## Privacy / Portfolio Version

This repository is intended as a portfolio-safe version of a real-world
automation workflow.

The public scripts intentionally omit: - Company-specific product
aliases - Confidential datasets - Internal distributor files - Sensitive
commercial information - Project-specific product/section rules

To use the scripts in another organization, supply your own dictionary
and normalization rules.

## Skills Demonstrated

-   Python automation
-   Excel processing with `openpyxl` and `pandas`
-   Data cleaning and normalization
-   Fuzzy string matching
-   Validation and defensive programming
-   Multi-file data consolidation
-   Persian-language data handling
-   Automated Excel report generation
-   Business process automation

## Future Improvements

Potential next steps: - Web dashboard for uploading inventory files -
Interactive validation/review screen - Automated charts and KPIs - Unit
tests - Command-line arguments and configuration files - Packaging as a
reusable Python application

## License

This repository is shared for portfolio and educational purposes. Add a
formal open-source license only if you intend to permit reuse,
modification, and redistribution.
