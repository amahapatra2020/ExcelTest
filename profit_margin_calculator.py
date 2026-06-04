import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment

# Load the workbook
wb = openpyxl.load_workbook('GitHub_Copilot_Test.xlsx')
ws = wb.active

# Find the header row
header_row = 1
headers = {}
for col_num, cell in enumerate(ws[header_row], 1):
    if cell.value:
        headers[cell.value] = col_num

# Get column indices
revenue_col = headers.get('Revenue (USD)')
cost_col = headers.get('Cost (USD)')

# Find the next available column for "Profit Margin (%)"
last_col = max(headers.values()) if headers else 1
profit_margin_col = last_col + 1

# Add header
ws.cell(row=header_row, column=profit_margin_col).value = 'Profit Margin (%)'
ws.cell(row=header_row, column=profit_margin_col).font = Font(bold=True)
ws.cell(row=header_row, column=profit_margin_col).fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")

# Add profit margin formula for each data row
for row in range(header_row + 1, ws.max_row + 1):
    revenue_cell = ws.cell(row=row, column=revenue_col)
    cost_cell = ws.cell(row=row, column=cost_col)
    profit_margin_cell = ws.cell(row=row, column=profit_margin_col)
    
    # Only add formula if revenue and cost cells have values
    if revenue_cell.value is not None and cost_cell.value is not None:
        # Formula: ((Revenue - Cost) / Revenue) * 100
        profit_margin_cell.value = f'=((C{row}-D{row})/C{row})*100'
        profit_margin_cell.number_format = '0.00'

# Save the workbook
wb.save('GitHub_Copilot_Test.xlsx')
print("Profit Margin column added successfully!")
print(f"Total rows processed: {ws.max_row - 1}")
