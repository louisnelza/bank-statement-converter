"""
parse_standardbank_pdf.py
-------------------------
Parses a Standard Bank credit card PDF statement and outputs a
3-column CSV ready for import into Sage One (Sage Accounting):

    Date,Description,Amount

- Credits (payments in) → negative amounts
- Debits (purchases/charges) → positive amounts
- Skips Balance Brought Forward and Closing Balance rows
- Handles the two-column transaction table layout
- Handles foreign currency rows (e.g. "Distrokid Musician 24.99 405.02")
  by stripping the foreign amount and using only the ZAR amount
- Password-protected PDFs: use --ask-password (or STATEMENT_PASSWORD env var)

Usage:
    python parse_standardbank_pdf.py <statement.pdf> [output.csv]
    python parse_standardbank_pdf.py <statement.pdf> [output.csv] --ask-password

Requirements:
    pip install -r requirements.txt   (pdfplumber, pypdf)
"""

import re
import sys
import csv
import os
import argparse
import getpass
import subprocess
import tempfile
import shutil


# ─────────────────────────────────────────────────────────────────────────────
# PDF unlocking
# ─────────────────────────────────────────────────────────────────────────────

def unlock_pdf(pdf_path, password):
    """
    Decrypt a password-protected PDF using qpdf (preferred) or pypdf (fallback).
    Returns (unlocked_path, tmp_dir) — caller must delete tmp_dir when done.
    """
    tmp_dir = tempfile.mkdtemp()
    unlocked_path = os.path.join(tmp_dir, "unlocked.pdf")

    if shutil.which("qpdf"):
        # Pass the password via a private temp file, not argv (argv is visible in `ps`)
        pw_path = os.path.join(tmp_dir, "pw.txt")
        fd = os.open(pw_path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "w") as f:
            f.write(password)
        try:
            result = subprocess.run(
                ["qpdf", f"--password-file={pw_path}", "--decrypt", pdf_path, unlocked_path],
                capture_output=True, text=True
            )
        finally:
            os.remove(pw_path)
        if result.returncode == 0 and os.path.exists(unlocked_path):
            return unlocked_path, tmp_dir

    # pypdf fallback
    try:
        from pypdf import PdfReader, PdfWriter
        reader = PdfReader(pdf_path)
        if reader.is_encrypted:
            reader.decrypt(password)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        with open(unlocked_path, "wb") as f:
            writer.write(f)
        return unlocked_path, tmp_dir
    except Exception as e:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise RuntimeError(f"Could not unlock PDF: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

MONTH_MAP = {
    "jan": "01", "feb": "02", "mar": "03", "apr": "04",
    "may": "05", "jun": "06", "jul": "07", "aug": "08",
    "sep": "09", "oct": "10", "nov": "11", "dec": "12"
}

def parse_date(day, month, year_2digit):
    """Convert ('16', 'Feb', '26') → '16/02/2026'."""
    month_num = MONTH_MAP.get(month.lower())
    if not month_num:
        return None
    year = f"20{year_2digit.zfill(2)}"
    return f"{int(day):02d}/{month_num}/{year}"


AMOUNT_RE = re.compile(r'^-?[\d,]+\.\d{2}$')

def is_amount(token):
    return bool(AMOUNT_RE.match(token))

def parse_amount(token):
    return float(token.replace(",", ""))

SKIP_ROWS = {
    "balance brought forward", "closing balance",
    "credits", "debit", "debits",
}


# ─────────────────────────────────────────────────────────────────────────────
# Column layout detection
# ─────────────────────────────────────────────────────────────────────────────

def detect_columns(words):
    """
    Locate the transaction table header row and derive the x-coordinate
    that separates the left and right columns.

    The header has 6 cells: Date | Description | Amount || Date | Description | Amount
    col_split = midpoint between left-Amount-x and right-Date-x
    """
    header_candidates = [
        w for w in words
        if w["text"].lower() in ("date", "description", "amount")
    ]
    by_y = {}
    for w in header_candidates:
        key = round(w["top"])
        by_y.setdefault(key, []).append(w)

    # Pick the row with the most header tokens (should be 6)
    header_row = sorted(
        max(by_y.values(), key=len),
        key=lambda w: w["x0"]
    )

    xs = [w["x0"] for w in header_row]
    if len(xs) >= 6:
        left_amount_x = xs[2]
        right_date_x  = xs[3]
    else:
        left_amount_x = 250
        right_date_x  = 300

    col_split = (left_amount_x + right_date_x) / 2
    table_top = min(w["top"] for w in header_row) if header_row else 0

    return col_split, table_top


# ─────────────────────────────────────────────────────────────────────────────
# Row assembly and token parsing
# ─────────────────────────────────────────────────────────────────────────────

def group_into_rows(words, table_top, y_tolerance=4):
    """Group words below the table header into horizontal rows."""
    table_words = [w for w in words if w["top"] >= table_top]
    rows = {}
    for w in table_words:
        key = round(w["top"] / y_tolerance) * y_tolerance
        rows.setdefault(key, []).append(w)
    return [
        sorted(rows[y], key=lambda w: w["x0"])
        for y in sorted(rows.keys())
    ]


def parse_tokens(tokens):
    """
    Parse one column's token list into a section marker or transaction.

    Handles normal rows, foreign-currency rows (strip foreign amount,
    keep ZAR), and signed credits.

    Returns:
      {"type": "section", "name": "credit"|"debit"}
      {"type": "transaction", "date", "description", "amount"}
      None — row should be skipped
    """
    if not tokens:
        return None

    joined_lower = " ".join(tokens).lower()

    if joined_lower in ("credits", "credit"):
        return {"type": "section", "name": "credit"}
    if joined_lower in ("debit", "debits"):
        return {"type": "section", "name": "debit"}

    if len(tokens) < 5:
        return None

    date = parse_date(tokens[0], tokens[1], tokens[2])
    if not date:
        return None

    remaining = tokens[3:]
    amount_positions = [i for i, t in enumerate(remaining) if is_amount(t)]
    if not amount_positions:
        return None

    # ZAR amount = last amount token on the line (handles FX rows)
    zar_amount  = parse_amount(remaining[amount_positions[-1]])
    first_amt   = amount_positions[0]
    description = " ".join(remaining[:first_amt]).strip()

    if not description or description.lower() in SKIP_ROWS:
        return None

    return {
        "type":        "transaction",
        "date":        date,
        "description": description,
        "amount":      zar_amount,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Main extraction
# ─────────────────────────────────────────────────────────────────────────────

def find_transaction_page(pages_words):
    for i, words in enumerate(pages_words):
        texts = " ".join(w["text"].lower() for w in words)
        if "transaction" in texts and "details" in texts:
            return i
    raise ValueError(
        "Could not find a 'Transaction details' page. "
        "Is this a Standard Bank credit card statement?"
    )


def extract_transactions(pdf_path):
    import pdfplumber

    pages_words = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages_words.append(page.extract_words())

    page_idx          = find_transaction_page(pages_words)
    words             = pages_words[page_idx]
    col_split, table_top = detect_columns(words)
    rows              = group_into_rows(words, table_top)

    transactions    = []
    current_section = "debit"

    for row_words in rows:
        left_tokens  = [w["text"] for w in row_words if w["x0"] <  col_split]
        right_tokens = [w["text"] for w in row_words if w["x0"] >= col_split]

        for tokens in (left_tokens, right_tokens):
            result = parse_tokens(tokens)
            if result is None:
                continue
            if result["type"] == "section":
                current_section = result["name"]
            else:
                transactions.append({
                    "date":        result["date"],
                    "description": result["description"],
                    "amount":      result["amount"],
                })

    return transactions


# ─────────────────────────────────────────────────────────────────────────────
# CSV output
# ─────────────────────────────────────────────────────────────────────────────

SAGE_MIN_ROWS = 16  # Sage One ZA pads to at least 16 data rows

def write_sage_csv(transactions, output_path):
    """Write Sage One ZA v2.0.0-compatible CSV: Date (DD/MM/YYYY), Description, Amount.
    Pads to a minimum of 16 data rows with blank trailing rows as required by Sage.
    """
    padding = max(0, SAGE_MIN_ROWS - len(transactions))
    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Date", "Description", "Amount"])
        for tx in transactions:
            writer.writerow([tx["date"], tx["description"], f"{tx['amount']:.2f}"])
        for _ in range(padding):
            writer.writerow(["", "", ""])
    print(f"\n✅  Done! {len(transactions)} transaction(s) written to: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert a Standard Bank PDF statement to a Sage One CSV.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("source", help="Path to the Standard Bank PDF statement")
    parser.add_argument("output", nargs="?", help="Output CSV path (optional)")
    parser.add_argument(
        "--password",
        help="(Deprecated: visible in shell history/process list) PDF password. "
             "Prefer --ask-password or the STATEMENT_PASSWORD environment variable.",
        default=None
    )
    parser.add_argument(
        "--ask-password", action="store_true",
        help="Prompt for the PDF password without echoing it"
    )
    args = parser.parse_args()

    if args.ask_password:
        args.password = getpass.getpass("PDF password: ")
    elif args.password:
        print("Warning: --password exposes the password in shell history; use --ask-password instead.",
              file=sys.stderr)
    else:
        args.password = os.environ.get("STATEMENT_PASSWORD") or None

    if not os.path.isfile(args.source):
        print(f"Error: file not found — {args.source}")
        sys.exit(1)

    output = args.output
    if output:
        if not output.lower().endswith(".csv"):
            output = os.path.splitext(output)[0] + ".csv"
            print(f"Note: output path corrected to '{output}' (must be .csv)")
    else:
        base, _ = os.path.splitext(args.source)
        output = base + "_sage.csv"

    tmp_dir  = None
    pdf_path = args.source

    try:
        if args.password:
            print("Unlocking PDF...")
            pdf_path, tmp_dir = unlock_pdf(args.source, args.password)
        else:
            try:
                import pdfplumber
                with pdfplumber.open(pdf_path) as pdf:
                    _ = pdf.pages[0].extract_words()
            except Exception:
                print("Error: This PDF appears to be password-protected.")
                print("Re-run with: --ask-password")
                sys.exit(1)

        print(f"Parsing: {args.source}")
        transactions = extract_transactions(pdf_path)

        if not transactions:
            print("Warning: No transactions found. Check that this is a Standard Bank credit card statement.")
            sys.exit(1)

        write_sage_csv(transactions, output)

    except Exception as e:
        print(f"\n❌  Error: {e}")
        sys.exit(1)
    finally:
        if tmp_dir:
            shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
