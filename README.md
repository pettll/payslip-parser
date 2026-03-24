# payslip-parser

Python CLI that reads payslip PDFs and writes a structured Excel workbook — useful for tracking salary progression, tax paid, and pay review history over time.

Handles three document types:

| Type | Filename pattern | Example |
|---|---|---|
| Monthly payslip | `YYYY-MM-cap.pdf` | `2025-03-cap.pdf` |
| P60 annual certificate | `p60-YYYY-YYYY.pdf` | `p60-2024-2025.pdf` |
| Pay / reward letter | contains `pay letter`, `reward letter`, or `benefit funding` | `2024 Reward Letter 1.pdf` |

## Output

Running the script produces `payslip_data.xlsx` with three sheets:

- **Monthly Payslips** — one row per payslip: date, tax code, base pay, all allowances, deductions, net pay, and YTD totals
- **P60 Annual** — one row per tax year: gross pay, tax deducted, NI breakdown, final tax code
- **Pay & Reward Letters** — one row per letter: type, effective date, salary table (current / new / increase / %)

## Requirements

Python 3.9+

```
pip install pdfplumber openpyxl
```

## Usage

```bash
# Parse all PDFs in ./payslips/ and write payslip_data.xlsx
python3 parse.py

# Dump raw extracted text from a single PDF (useful for debugging a new format)
python3 parse.py --debug 2025-03-cap.pdf
```

Drop your PDFs into `payslips/` and run. The script classifies each file automatically by filename pattern.

## Project Structure

```
payslip-parser/
├── parse.py          Parser and Excel writer
├── payslips/         Drop PDFs here (gitignored)
└── payslip_data.xlsx Output workbook (gitignored)
```

## Payslip Format

The parser is calibrated for **SDWorx**-generated payslip PDFs (used by Capgemini). It uses `pdfplumber` word-level coordinate extraction to split the three-column table layout (Pay & Allowances | Deductions | Totals & Balances) rather than relying on raw text order.

Pay items (e.g. overtime, bonus, allowances) are discovered dynamically — any item that appears in the Allowances column is added as a column in the output, so the sheet adapts to payslips with different structures.

To adapt to a different payslip provider, use `--debug` to inspect the raw text, then adjust the column x-boundaries (`_COL2_X`, `_COL3_X`) and header/field regex patterns in `parse.py`.

## Privacy

`payslips/` and all `.xlsx` files are gitignored. No data leaves your machine.

## License

MIT
