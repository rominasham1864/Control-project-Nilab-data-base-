import openpyxl
from openpyxl.styles import PatternFill

# Load the Excel file
workbook = openpyxl.load_workbook('main.xlsx')

# Select the worksheet
worksheet = workbook.active

# Change the color of row 2 to red
row_number = 2
fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
for cell in worksheet[row_number]:
    cell.fill = fill

# Save the modified Excel file
workbook.save('main.xlsx')