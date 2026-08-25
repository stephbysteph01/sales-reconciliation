# Sales Reconciliation

Python project to consolidate Excel sales files from three branches and reconcile them against a payments file, identifying anomalies such as missing payments, inconsistent amounts, duplicates, and non-standardized customer names.

## Goal

The project generates realistic fake data in Italian with Faker and saves it as `.xlsx` files, then normalizes and reconciles the information to highlight business discrepancies.

## Structure

- `src/sales_reconciliation/data_generator.py`: generates the input Excel files.
- `src/sales_reconciliation/reconcile.py`: cleans and reconciles the data.
- `src/sales_reconciliation/cli.py`: command-line entry point for generating and reconciling data.
- `tests/test_reconcile.py`: minimal automated tests.
- `data/raw/`: generated input Excel files.
- `data/output/`: exported reconciliation reports.

## Setup

```bash
python -m venv venv
venv\Scripts\python.exe -m pip install -e '.[dev]'
```

## Generate demo data

```bash
venv\Scripts\python.exe -m sales_reconciliation.data_generator
```

This creates:

- `data/raw/vendite_roma.xlsx`
- `data/raw/vendite_milano.xlsx`
- `data/raw/vendite_firenze.xlsx`
- `data/raw/incassi.xlsx`

## Reconciliation workflow

There are two ways to use the project:

1. Via the CLI, which is the fastest and simplest way to run the workflow without writing code.
2. Via a Python script, which is useful for customizing the flow, debugging, or integrating it into other programs.

### 1) Fastest path: CLI

```bash
venv\Scripts\python.exe -m sales_reconciliation.cli generate --output-dir data/raw
venv\Scripts\python.exe -m sales_reconciliation.cli reconcile --input-dir data/raw --output data/output/reconciliation_report.xlsx --csv-output data/output/reconciliation_details.csv
```

The CLI is useful for:

- generating input Excel files quickly;
- running reconciliation in a repeatable way without writing a one-off script;
- exporting Excel and CSV reports ready for analysis or sharing across teams;
- integrating the workflow into automation or scheduled jobs.

### 2) Direct Python usage

```python
from pathlib import Path
from sales_reconciliation.reconcile import export_reconciliation_report, read_payments_file, read_sales_files, reconcile_sales_and_payments

sales_df = read_sales_files(Path("data/raw"))
payments_df = read_payments_file(Path("data/raw/incassi.xlsx"))
report = reconcile_sales_and_payments(sales_df, payments_df)
print(report)
export_reconciliation_report(report, Path("data/output/report.xlsx"), sales_df, payments_df)
```

This approach is useful when you want to:

- see the processing flow explicitly;
- customize or extend the steps;
- access the dataframes directly for analysis or debugging;
- use the code as a base for a notebook or a larger module.

In short, the CLI is the preferred option for day-to-day use, while the Python script is the more transparent and customizable version of the same workflow.

The exported Excel file includes: a summary, ID lists, and dedicated sheets with detailed records for each anomaly category (missing payments, mismatched payments, duplicate sales, and duplicate cash entries). The same function also exports a CSV with all detailed records.

## "Dirty" data included

The generated data intentionally includes:

- about 15% of sales with no matching payment;
- payment amounts slightly different from the sale amount;
- some duplicated `ID_Vendita` values;
- inconsistent customer names across files;
- missing dates or dates in different formats.

## Testing

```bash
pytest -q
```
