import openpyxl

def write_table_to_excel(table_data, file_path):
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    for row_idx, row_data in enumerate(table_data, start=1):
        for col_idx, cell_value in enumerate(row_data, start=1):
            worksheet.cell(row=row_idx, column=col_idx).value = cell_value

    workbook.save(file_path)

# Usage example
table_data = [
    ["Nam", "Age", "Email"],
    ["John Doe", "30", "john@example.com"],
    ["Jane Smith", "25", "jane@example.com"]
]
file_path = "output.xlsx"
write_table_to_excel(table_data, file_path)