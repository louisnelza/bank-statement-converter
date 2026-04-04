# bank-statement-converter

A command-line Python tool that converts bank statement exports (.xls / .xlsx) into clean CSV files ready for import into accounting software like [Wave](https://www.waveapps.com/accounting), Xero, or QuickBooks.

Driven entirely by a JSON config file — no code changes needed to support a new bank or accounting tool.

---

## The problem it solves

Most banks export statements in their own proprietary spreadsheet format. Accounting tools like Wave expect a specific CSV layout. Manually copying and reformatting transactions every month is tedious and error-prone.

**Source (Bank Zero `.xls` export):**

| Date | Day | Time | Type | Description 1 | Description 2 | Fee | Amount | Balance |
|------|-----|------|------|---------------|---------------|-----|--------|---------|
| 2025-12-12 | Fri | 08:16 | Pay in | ACME CLIENT | Invoice #1042 payment | 0 | 2500.00 | 3100.00 |
| 2025-12-12 | Fri | 15:14 | Card purchase | Office Supplies Co | Tap Transaction | 0 | -84.50 | 3015.50 |

**Output (Wave-ready CSV):**

| Date | Description | Deposits | Withdrawals |
|------|-------------|----------|-------------|
| 2025-12-12 | ACME CLIENT - Invoice #1042 payment | 2,500.00 | |
| 2025-12-12 | Office Supplies Co - Tap Transaction | | 84.50 |

---

## Features

- Reads `.xls` (legacy) and `.xlsx` / `.xlsm` files
- Auto-detects the transactions sheet by name (e.g. `Dec 25 Transactions`, `Jan 26 Transactions`)
- Merges multiple description columns into one clean field
- Splits a signed `Amount` column into separate `Deposits` and `Withdrawals` columns
- Configurable date format, thousands separator, column names, and output structure
- Cross-platform: works on Windows, macOS, and Linux
- Enforces `.csv` output extension regardless of what you type

---

## Requirements

- Python 3.8+
- [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/) — only needed for `.xls` files (free, cross-platform)

Install Python dependencies:

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
# Output saved as <source>_converted.csv in the same folder
python converter.py january.xls

# Specify output path
python converter.py january.xls wave_import.csv

# Use a custom config file
python converter.py january.xls wave_import.csv --config my_bank_config.json
```

---

## Configuration

All conversion logic lives in `config.json`. The default config maps **Bank Zero → Wave Accounting**.

```json
{
  "source": {
    "sheet": { "contains": "Transactions" },
    "description_separator": " - ",
    "columns": {
      "date":         "Date",
      "description1": "Description 1",
      "description2": "Description 2",
      "amount":       "Amount"
    }
  },
  "output": {
    "date_format": "%Y-%m-%d",
    "thousands_separator": true,
    "columns": [
      { "header": "Date",        "type": "date" },
      { "header": "Description", "type": "description", "sources": ["description1", "description2"] },
      { "header": "Deposits",    "type": "deposit" },
      { "header": "Withdrawals", "type": "withdrawal" }
    ]
  }
}
```

### Sheet selection

| Option | Example | Behaviour |
|--------|---------|-----------|
| `contains` | `"contains": "Transactions"` | First sheet whose name contains the string |
| `name` | `"name": "Sheet1"` | Exact sheet name match |
| `index` | `"index": 0` | Sheet by zero-based position |

### Output column types

| Type | Description |
|------|-------------|
| `date` | Formats the date column using `date_format` |
| `description` | Joins one or more source columns with `description_separator` |
| `deposit` | Amount if positive, blank if negative |
| `withdrawal` | Absolute amount if negative, blank if positive |
| `amount` | Raw signed amount |
| `passthrough` | Copies a source column as-is |
| `literal` | Fixed string value for every row |

### Adapting to another bank

Copy `config.json` to a new file (e.g. `fnb_config.json`) and update the `columns` keys to match your bank's header names. Pass it in with `--config fnb_config.json`.

---

## Project structure

```
bank-statement-converter/
├── converter.py       # Main script
├── config.json        # Bank Zero → Wave config (default)
├── requirements.txt
└── README.md
```

---

## Why LibreOffice for .xls?

The legacy `.xls` format (BIFF8) requires the `xlrd` library, which dropped support for it in v2.0 due to security concerns. LibreOffice converts `.xls` to `.xlsx` transparently, which `openpyxl` can then read safely. If your bank already exports `.xlsx`, LibreOffice is not needed at all.

---

## License

MIT
