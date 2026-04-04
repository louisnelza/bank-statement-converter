"""
converter.py
------------
Converts a bank statement spreadsheet (.xls / .xlsx) into a clean CSV file
ready for import into accounting software such as Wave, Xero, or QuickBooks.

The conversion is fully driven by a JSON config file, making it easy to adapt
to any bank export format or accounting tool without touching this code.

Usage:
    python converter.py <source.xls> [output.csv] [--config config.json]

    source.xls   Path to your bank export file.
    output.csv   (Optional) Output path. Defaults to <source>_converted.csv
    --config     (Optional) Path to config file. Defaults to config.json in
                 the same directory as this script.

Examples:
    python converter.py january.xls
    python converter.py january.xls wave_import.csv
    python converter.py january.xls wave_import.csv --config my_config.json

Requirements:
    pip install -r requirements.txt
    LibreOffice (for .xls files): https://www.libreoffice.org/download/
"""

import sys
import os
import csv
import json
import shutil
import argparse
import subprocess
import tempfile
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# Config loading
# ─────────────────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG = os.path.join(SCRIPT_DIR, "config.json")


def load_config(path):
    if not os.path.isfile(path):
        raise FileNotFoundError(
            f"Config file not found: {path}\n"
            "Make sure config.json is in the same folder as converter.py, "
            "or pass --config <path>."
        )
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ─────────────────────────────────────────────────────────────────────────────
# LibreOffice: .xls → .xlsx conversion
# ─────────────────────────────────────────────────────────────────────────────

def find_libreoffice():
    """Locate the LibreOffice executable across platforms."""
    exe = shutil.which("libreoffice") or shutil.which("soffice")
    if exe:
        return exe
    # Common Windows install locations (not always on PATH)
    win_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in win_paths:
        if os.path.isfile(p):
            return p
    return None


def convert_xls_to_xlsx(xls_path):
    """
    Use LibreOffice to convert a legacy .xls file to .xlsx.
    Returns (xlsx_path, tmp_dir) — caller must delete tmp_dir when done.
    """
    lo_exe = find_libreoffice()
    if lo_exe is None:
        raise RuntimeError(
            "LibreOffice could not be found. It is required to read .xls files.\n"
            "Install it from: https://www.libreoffice.org/download/download-libreoffice/\n"
            "Alternatively, resave your file as .xlsx in Excel and run again."
        )

    tmp_dir = tempfile.mkdtemp()
    result = subprocess.run(
        [lo_exe, "--headless", "--convert-to", "xlsx", xls_path, "--outdir", tmp_dir],
        capture_output=True, text=True, timeout=60
    )
    if result.returncode != 0:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise RuntimeError(f"LibreOffice conversion failed:\n{result.stderr}")

    base = os.path.splitext(os.path.basename(xls_path))[0]
    xlsx_path = os.path.join(tmp_dir, base + ".xlsx")
    if not os.path.exists(xlsx_path):
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise RuntimeError(
            f"Converted file not found at expected path: {xlsx_path}\n{result.stdout}"
        )
    return xlsx_path, tmp_dir


# ─────────────────────────────────────────────────────────────────────────────
# Workbook loading
# ─────────────────────────────────────────────────────────────────────────────

def open_workbook(path):
    """
    Open any Excel file with openpyxl.
    Handles .xls via LibreOffice conversion; .xlsx/.xlsm directly.
    Returns (workbook, tmp_dir_or_None).
    """
    import openpyxl
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return openpyxl.load_workbook(path), None
    if ext == ".xls":
        xlsx_path, tmp_dir = convert_xls_to_xlsx(path)
        return openpyxl.load_workbook(xlsx_path), tmp_dir
    raise ValueError(f"Unsupported file type: '{ext}'. Supported: .xls, .xlsx, .xlsm")


# ─────────────────────────────────────────────────────────────────────────────
# Sheet selection
# ─────────────────────────────────────────────────────────────────────────────

def find_sheet(workbook, sheet_config):
    """
    Find the right worksheet based on config.
    sheet_config can be:
      - {"name": "Transactions"}        → exact match
      - {"contains": "Transactions"}    → first sheet whose name contains the value
      - {"index": 0}                    → sheet by zero-based index
    """
    if "name" in sheet_config:
        name = sheet_config["name"]
        if name not in workbook.sheetnames:
            raise ValueError(
                f"Sheet '{name}' not found. Available: {workbook.sheetnames}"
            )
        return workbook[name]

    if "contains" in sheet_config:
        needle = sheet_config["contains"].lower()
        for sname in workbook.sheetnames:
            if needle in sname.lower():
                return workbook[sname]
        raise ValueError(
            f"No sheet containing '{sheet_config['contains']}' found. "
            f"Available: {workbook.sheetnames}"
        )

    if "index" in sheet_config:
        idx = sheet_config["index"]
        return workbook.worksheets[idx]

    raise ValueError("Config 'sheet' must have 'name', 'contains', or 'index'.")


# ─────────────────────────────────────────────────────────────────────────────
# Value formatters
# ─────────────────────────────────────────────────────────────────────────────

def format_date(value, fmt):
    """Format a date cell value using the given strftime format string."""
    if isinstance(value, datetime):
        return value.strftime(fmt)
    if value is None:
        return ""
    # Try parsing common string date formats
    for parse_fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(str(value).strip(), parse_fmt).strftime(fmt)
        except ValueError:
            continue
    return str(value)


def format_amount(value, thousands_sep):
    """Format a numeric amount, optionally with thousands separator."""
    if value is None:
        return ""
    if thousands_sep:
        return f"{value:,.2f}"
    return f"{value:.2f}"


def safe_float(value):
    """Safely convert a cell value to float."""
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def resolve_cell_value(row, col_map, key, headers):
    """
    Get a cell value from a row using the column mapping in config.
    col_map[key] can be a column name (str) or zero-based index (int).
    Returns None if key not in col_map.
    """
    if key not in col_map:
        return None
    ref = col_map[key]
    if isinstance(ref, int):
        return row[ref] if ref < len(row) else None
    # String: look up by header name
    ref_l = ref.lower()
    for i, h in enumerate(headers):
        if h.lower() == ref_l:
            return row[i] if i < len(row) else None
    raise ValueError(f"Column '{ref}' not found in headers: {headers}")


# ─────────────────────────────────────────────────────────────────────────────
# Core transformation
# ─────────────────────────────────────────────────────────────────────────────

def transform(source_path, output_path, config):
    src_cfg    = config["source"]
    out_cfg    = config["output"]
    col_map    = src_cfg["columns"]
    date_fmt   = out_cfg.get("date_format", "%Y-%m-%d")
    thousands  = out_cfg.get("thousands_separator", True)
    desc_sep   = src_cfg.get("description_separator", " - ")

    print(f"Opening source file: {source_path}")
    workbook, tmp_dir = open_workbook(source_path)

    try:
        sheet = find_sheet(workbook, src_cfg["sheet"])
        print(f"  → Using sheet: '{sheet.title}'")

        all_rows = list(sheet.iter_rows(values_only=True))
        if not all_rows:
            raise ValueError("The selected sheet is empty.")

        # Auto-detect header row: first row containing the date column name
        date_col_name = col_map.get("date", "Date")
        header_idx = None
        for i, row in enumerate(all_rows):
            if any(
                str(c).strip().lower() == str(date_col_name).lower()
                for c in row if c is not None
            ):
                header_idx = i
                break
        if header_idx is None:
            raise ValueError(
                f"Could not find header row containing '{date_col_name}'."
            )

        headers   = [str(c).strip() if c is not None else "" for c in all_rows[header_idx]]
        data_rows = all_rows[header_idx + 1:]
        print(f"  → {len(data_rows)} row(s) found after header.")

        # Build output column list from config
        output_columns = out_cfg["columns"]  # ordered list of output column definitions

        count = 0
        # Enforce .csv extension on output
        if not output_path.lower().endswith(".csv"):
            output_path = os.path.splitext(output_path)[0] + ".csv"

        with open(output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow([c["header"] for c in output_columns])

            for src_row in data_rows:
                if all(c is None for c in src_row):
                    continue  # skip blank rows

                out_row = []
                amount = safe_float(resolve_cell_value(src_row, col_map, "amount", headers))

                for col_def in output_columns:
                    col_type = col_def["type"]

                    if col_type == "date":
                        raw = resolve_cell_value(src_row, col_map, "date", headers)
                        out_row.append(format_date(raw, date_fmt))

                    elif col_type == "description":
                        # Support one or more source columns joined by separator
                        parts = []
                        for src_key in col_def.get("sources", ["description"]):
                            val = resolve_cell_value(src_row, col_map, src_key, headers)
                            if val is not None and str(val).strip():
                                parts.append(str(val).strip())
                        out_row.append(desc_sep.join(parts))

                    elif col_type == "deposit":
                        # Positive amounts only
                        if amount is not None and amount >= 0:
                            out_row.append(format_amount(amount, thousands))
                        else:
                            out_row.append("")

                    elif col_type == "withdrawal":
                        # Negative amounts, stored as positive value
                        if amount is not None and amount < 0:
                            out_row.append(format_amount(abs(amount), thousands))
                        else:
                            out_row.append("")

                    elif col_type == "amount":
                        # Raw signed amount
                        out_row.append(format_amount(amount, thousands) if amount is not None else "")

                    elif col_type == "literal":
                        # A fixed value defined in config
                        out_row.append(col_def.get("value", ""))

                    elif col_type == "passthrough":
                        # Copy a source column as-is
                        raw = resolve_cell_value(src_row, col_map, col_def["source"], headers)
                        out_row.append(str(raw).strip() if raw is not None else "")

                    else:
                        out_row.append("")

                writer.writerow(out_row)
                count += 1

        print(f"\n✅  Done! {count} transaction(s) written to: {output_path}")

    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert a bank statement spreadsheet to a clean accounting CSV.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("source", help="Path to the bank export file (.xls or .xlsx)")
    parser.add_argument("output", nargs="?", help="Output CSV path (optional)")
    parser.add_argument(
        "--config", default=DEFAULT_CONFIG,
        help=f"Path to JSON config file (default: {DEFAULT_CONFIG})"
    )
    args = parser.parse_args()

    if not os.path.isfile(args.source):
        print(f"Error: source file not found — {args.source}")
        sys.exit(1)

    if args.output:
        output = args.output
        # Always enforce .csv extension
        if not output.lower().endswith(".csv"):
            output = os.path.splitext(output)[0] + ".csv"
            print(f"Note: output path corrected to '{output}' (must be .csv)")
    else:
        base, _ = os.path.splitext(args.source)
        output = base + "_converted.csv"

    try:
        config = load_config(args.config)
        transform(args.source, output, config)
    except Exception as e:
        print(f"\n❌  Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
