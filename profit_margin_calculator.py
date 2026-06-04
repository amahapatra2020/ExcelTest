"""profit_margin_calculator.py

Robust profit margin calculator for the repository Excel file.

Usage:
  - Default: writes a new file named <original>_with_margin.xlsx (does not overwrite original)
  - To overwrite original file, pass --inplace

It detects header row, finds the columns containing Revenue and Cost (case-insensitive, accepts headers containing the words "revenue" and "cost") and writes a new column "Profit Margin (%)" with numeric percentage values formatted as 2 decimals.

Handles divide-by-zero and non-numeric cells gracefully.
"""

import argparse
import os
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill


def normalize(s):
    return ''.join(c for c in (s or '').strip().lower())


def find_header_row(ws, max_rows_to_check=10):
    # Return first row number that contains any non-empty cell
    for r in range(1, min(max_rows_to_check, ws.max_row) + 1):
        if any(cell.value is not None and str(cell.value).strip() != '' for cell in ws[r]):
            return r
    return 1


def find_column_index(headers_map, keywords):
    # keywords: list of keywords to match (all must be present in header text)
    for name, idx in headers_map.items():
        if all(k in name for k in keywords):
            return idx
    return None


def main():
    parser = argparse.ArgumentParser(description='Add Profit Margin (%) column to an Excel file')
    parser.add_argument('file', nargs='?', default='GitHub_Copilot_Test.xlsx', help='Path to the Excel file')
    parser.add_argument('--inplace', action='store_true', help='Overwrite original file instead of creating a new file')
    args = parser.parse_args()

    infile = args.file
    if not os.path.exists(infile):
        print(f'Error: file not found: {infile}', file=sys.stderr)
        sys.exit(2)

    wb = load_workbook(infile)
    ws = wb.active

    header_row = find_header_row(ws)

    # Build headers map: normalized header text -> column index (1-based)
    headers = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            headers[normalize(str(cell.value))] = cell.column

    # Try to find revenue and cost columns by matching keywords
    revenue_col = find_column_index(headers, ['revenue'])
    cost_col = find_column_index(headers, ['cost'])

    if revenue_col is None or cost_col is None:
        # Try looser matching: header contains the word revenue/cost anywhere
        for name, idx in headers.items():
            if 'revenue' in name and revenue_col is None:
                revenue_col = idx
            if 'cost' in name and cost_col is None:
                cost_col = idx

    if revenue_col is None or cost_col is None:
        print('Could not automatically determine Revenue and/or Cost columns. Found headers:')
        for name, idx in headers.items():
            print(f'  - {name!r} -> column {get_column_letter(idx)}')
        print('\nPlease ensure your headers contain the words "Revenue" and "Cost" (case-insensitive).', file=sys.stderr)
        sys.exit(3)

    # Determine where to add Profit Margin column (append after last used column)
    last_col = max(cell.column for cell in ws[header_row])
    profit_col = last_col + 1

    # Write header
    profit_header_cell = ws.cell(row=header_row, column=profit_col)
    profit_header_cell.value = 'Profit Margin (%)'
    profit_header_cell.font = Font(bold=True)
    profit_header_cell.fill = PatternFill(start_color='FFEEEEEE', end_color='FFEEEEEE', fill_type='solid')

    rows_processed = 0
    rows_computed = 0
    rows_skipped = 0

    for r in range(header_row + 1, ws.max_row + 1):
        rows_processed += 1
        rev_cell = ws.cell(row=r, column=revenue_col)
        cost_cell = ws.cell(row=r, column=cost_col)
        out_cell = ws.cell(row=r, column=profit_col)

        # Read numeric values
        try:
            rev_val = float(rev_cell.value) if rev_cell.value is not None and str(rev_cell.value).strip() != '' else None
        except Exception:
            rev_val = None
        try:
            cost_val = float(cost_cell.value) if cost_cell.value is not None and str(cost_cell.value).strip() != '' else None
        except Exception:
            cost_val = None

        if rev_val is None:
            out_cell.value = None
            rows_skipped += 1
            continue

        if rev_val == 0:
            # Avoid division by zero; set cell text to None or a string
            out_cell.value = None
            rows_skipped += 1
            continue

        if cost_val is None:
            # If cost missing, treat as 0
            cost_val = 0.0

        margin = (rev_val - cost_val) / rev_val
        out_cell.value = margin
        out_cell.number_format = '0.00%'
        rows_computed += 1

    # Save workbook
    if args.inplace:
        outpath = infile
    else:
        base, ext = os.path.splitext(infile)
        outpath = f"{base}_with_margin{ext}"

    wb.save(outpath)

    print(f'Input file: {infile}')
    print(f'Output file: {outpath}')
    print(f'Header row detected: {header_row}')
    print(f'Revenue column: {get_column_letter(revenue_col)} (index {revenue_col})')
    print(f'Cost column: {get_column_letter(cost_col)} (index {cost_col})')
    print(f'Added Profit Margin column at: {get_column_letter(profit_col)} (index {profit_col})')
    print(f'Rows processed: {rows_processed}; computed: {rows_computed}; skipped: {rows_skipped}')


if __name__ == '__main__':
    main()
