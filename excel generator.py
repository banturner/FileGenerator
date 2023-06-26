import openpyxl

def generate_excel_sheet(filename, data):
    # Create a new workbook and select the active sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write data to the sheet
    for row_num, row_data in enumerate(data, start=1):
        for col_num, cell_data in enumerate(row_data, start=1):
            sheet.cell(row=row_num, column=col_num).value = cell_data

    # Save the workbook
    workbook.save(filename)
    print(f"Excel sheet '{filename}' generated successfully.")

# Example usage
data = [
    ["Name", "Age", "Country"],
    ["John", 25, "USA"],
    ["Alice", 30, "Canada"],
    ["Bob", 22, "UK"]
]

filename = "example.xlsx"
generate_excel_sheet(filename, data)
