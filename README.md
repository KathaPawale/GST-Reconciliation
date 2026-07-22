# GST-Reconciliation

Compares GST Portal B2B data against Zoho Purchase/Bill data and generates a color-coded reconciliation workbook.

## What it does

- Accepts either:
  - One Excel file containing both sheets:
    - `GST Portal B2B`
    - `Zoho Purchase/Bill`
  - Or two CSV files (`--b2b-csv` and `--zoho-csv`)
- Cleans data by trimming spaces, uppercasing GSTIN, normalizing dates, parsing numeric tax values, and treating blank/NA values as missing.
- Matches records by GSTIN + Invoice/Bill Number + Invoice/Bill Date.
- Compares IGST/CGST/SGST with tolerance (default `1.0`).
- Produces statuses and reasons for every row:
  - `MATCH` (yellow tax cells)
  - `PREVIOUS_MONTH` (green in B2B rows)
  - `ZOHO_TAX_BLANK` (red rows)
  - `MISMATCH` with reason (no color)
- Exports one workbook with 4 sheets:
  - `B2B`
  - `Zoho`
  - `Available in B2B but Not in Zoho`
  - `Available in Zoho but Not in B2B`
- Adds per-sheet summary totals and counts.

## Usage

```bash
python /home/runner/work/GST-Reconciliation/GST-Reconciliation/gst_reconciliation.py \
  --input /absolute/path/to/input.xlsx \
  --output /absolute/path/to/reconciliation-output.xlsx
```

CSV mode:

```bash
python /home/runner/work/GST-Reconciliation/GST-Reconciliation/gst_reconciliation.py \
  --b2b-csv /absolute/path/to/b2b.csv \
  --zoho-csv /absolute/path/to/zoho.csv \
  --output /absolute/path/to/reconciliation-output.xlsx
```

Optional mapping for non-standard column names:

```bash
--mapping-json '{"b2b":{"invoice_number":"Bill No"},"zoho":{"invoice_number":"Invoice #"}}'
```
