# bank-statement-converter

A command-line Python toolkit that converts bank statement exports into clean CSV files ready for import into accounting software like [Wave](https://www.waveapps.com/accounting), [Sage One](https://www.sage.com/en-za/sage-business-cloud/accounting/), Xero, or QuickBooks.

Two scripts cover the most common workflows:

| Script | Input | Output | Target |
|--------|-------|--------|--------|
| `converter.py` | Bank Zero `.xls` export | 4-column CSV | Wave Accounting |
| `parse_standardbank_pdf.py` | Standard Bank PDF statement | 3-column CSV | Sage One |

Both are config-driven and easy to adapt to other banks or accounting tools.

---

## The problem they solve

Banks export statements in proprietary formats. Accounting tools expect specific CSV layouts. Manually copying and reformatting transactions every month is tedious and error-prone.

### Bank Zero → Wave (`converter.py`)

**Source (Bank Zero `.xls` export):**

| Date | Day | Time | Type | Description 1 | Description 2 | Fee | Amount | Balance |
|------|-----|------|------|---------------|---------------|-----|--------|---------|
| 2025-12-12 | Fri | 08:16 | Pay in | ACME CLIENT | Invoice #1042 payment | 0 | 2500.00 | 3100.00 |
| 2025-12-12 | Fri | 15:14 | Card purchase | Office Supplies Co | Tap Transaction | 0 | -84.50 | 3015.50 |

**Output (Wave-ready 4-column CSV):**

| Date | Description | Deposits | Withdrawals |
|------|-------------|----------|-------------|
| 2025-12-12 | ACME CLIENT - Invoice #1042 payment | 2,500.00 | |
| 2025-12-12 | Office Supplies Co - Tap Transaction | | 84.50 |

### Standard Bank → Sage One (`parse_standardbank_pdf.py`)

**Source (Standard Bank PDF — two-column layout, may be password-protected):**

| Date | Description | Amount |
|------|-------------|--------|
| 2 Mar 26 | Apo Payment, Thank You | -4,497.14 |
| 16 Feb 26 | Rtmc | 1,119.00 |

**Output (Sage One-ready 3-column CSV):**

| Date | Description | Amount |
|------|-------------|--------|
| 2026-03-02 | Apo Payment, Thank You | -4497.14 |
| 2026-02-16 | Rtmc | 1119.00 |

---

## Features

### converter.py (Bank Zero → Wave)
- Reads `.xls` (legacy) and `.xlsx` / `.xlsm` files
- Auto-detects the transactions sheet by name (e.g. `Dec 25 Transactions`, `Jan 26 Transactions`)
- Merges multiple description columns into one clean field
- Splits a signed `Amount` column into separate `Deposits` and `Withdrawals` columns
- Fully config-driven — adapt to any bank or accounting tool via `config.json`
- Enforces `.csv` output extension

### parse_standardbank_pdf.py (Standard Bank → Sage One)
- Parses Standard Bank credit card PDF statements
- Handles the two-column transaction table layout
- Handles foreign currency rows — strips foreign amount, uses ZAR amount
- Unlocks password-protected PDFs (Standard Bank locks statements with your ID number)
- Uses `qpdf` for decryption if available, falls back to `pypdf`

---

## Requirements

- Python 3.8+
- [LibreOffice](https://www.libreoffice.org/download/) — only needed for `.xls` files in `converter.py`
- `qpdf` — optional but recommended for PDF decryption (`winget install qpdf` / `brew install qpdf` / `apt install qpdf`)

Install Python dependencies:

```bash
pip install -r requirements.txt
```

---

## Usage

### converter.py

```bash
# Output saved as <source>_converted.csv in the same folder
python converter.py january.xls

# Specify output path
python converter.py january.xls wave_import.csv

# Use a custom config file
python converter.py january.xls wave_import.csv --config my_bank_config.json
```

### parse_standardbank_pdf.py

```bash
# Unlocked PDF
python parse_standardbank_pdf.py statement.pdf

# Password-protected PDF (Standard Bank uses your SA ID number as the password)
python parse_standardbank_pdf.py statement.pdf --ask-password

# Specify output path
python parse_standardbank_pdf.py statement.pdf sage_import.csv --ask-password
```

---

## Configuration (converter.py)

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

Copy `config.json` to a new file (e.g. `fnb_config.json`) and update the column names to match your bank's export headers. Pass it in with `--config fnb_config.json`.

---

## Project structure

```
bank-statement-converter/
├── converter.py                  # Bank Zero .xls → Wave CSV
├── parse_standardbank_pdf.py     # Standard Bank PDF → Sage One CSV
├── config.json                   # Bank Zero → Wave column mapping
├── requirements.txt
└── README.md
```

---

## Notes

**Why LibreOffice for .xls?**
The legacy `.xls` format (BIFF8) requires `xlrd`, which dropped support for it in v2.0 due to security concerns. LibreOffice converts `.xls` to `.xlsx` transparently, which `openpyxl` can then read safely. If your bank exports `.xlsx`, LibreOffice is not needed.

**Why is the Standard Bank PDF password-protected?**
Standard Bank locks PDF statements with the account holder's South African ID number. Pass it via `--password` and the script decrypts it automatically before parsing.

**Sage One import steps:**
1. In Sage One, go to **Banking** → select your account → **Actions** → **Import Statement**
2. Upload the CSV generated by `parse_standardbank_pdf.py`
3. Select date format `YYYY-MM-DD`
4. Review and confirm

---

## License

MIT
