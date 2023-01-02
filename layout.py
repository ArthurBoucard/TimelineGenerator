import openpyxl
import json

# Open the JSON file
with open('config.json', 'r') as f:
    # Load the JSON data into a Python object
    config = json.load(f)

# Create a new Excel workbook
wb = openpyxl.load_workbook("timeline.xlsx")
sheet = wb.active

def apply_layout():
    # Iterate through the first 2 columns in the sheet
    for cell in sheet.iter_cols(min_col=1, max_col=2):
        for i in range(len(cell)):
            # Make column size fit text length
            if len(cell[i].value) > sheet.column_dimensions[chr(ord('@') + cell[i].column)].width:
                sheet.column_dimensions[chr(ord('@') + cell[i].column)].width = len(cell[i].value)

    # Save the Excel workbook
    wb.save("timeline.xlsx")