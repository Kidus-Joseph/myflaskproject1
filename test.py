import sys
from openpyxl import workbook, load_workbook
import pandas as pd
import numpy as np

dataset1 = pd.read_excel("data/Sample UPC Data.xlsx")
UPC1 = dataset1["UPC"]
MODEL1 = dataset1["Model Number"]
CATEGORY1 = dataset1["Category"]
SUBCATEGORY1 = dataset1["Subcategory"]
BRAND1 = dataset1["Brand"]
# approved_Denied = dataset1['Approve/Denial Status']
# dataset1.set_index("Item ID")

workbook1 = load_workbook("data/Sample UPC Data.xlsx")
sheet1 = workbook1.active

upc = workbook1["Confirmed_UPC"]

code = int(input("Enter your code:"))
m = str(input("Enter your model number:"))
c = str(input("Enter your category:"))
s = str(input("Enter your subcategory:"))
b = str(input("Enter your brand:"))

column_letter = 'O'
success_column_letter = 'N'
row_number = dataset1[dataset1["UPC"] == code].index.to_list()
try:
    row_number = str(row_number)[1:-1]
    row_number = int(row_number) + 2
    print(row_number)
except ValueError as e:
    error_message = f"Error: Invalid row number format. Reason: {str(e)}"
    cell = sheet1[success_column_letter + str(row_number)]
    cell.value = "Denied"
    cell1 = sheet1[column_letter + str(row_number)]
    cell1.value = "UPC Not Found"
    filename = "Verification_Report.xlsx"
    workbook1.save(filename)
    raise ValueError(error_message)
    sys.exit(1)
print(row_number)

match_found = False

for index, u in enumerate(UPC1):
    if code == u:
        model = MODEL1[index]
        category = CATEGORY1[index]
        subcategory = SUBCATEGORY1[index]
        brand = BRAND1[index]

        if m == model and c == category and s == subcategory and b == brand:
            cell = sheet1[success_column_letter + str(row_number)]
            cell.value = "Approved"
            filename = "Verification_Report.xlsx"
            workbook1.save(filename)
            match_found = True
            break

if not match_found:
    cell = sheet1[success_column_letter + str(row_number)]
    cell.value = "Denied"
    cell1 = sheet1[column_letter + str(row_number)]
    cell1.value = "UPC Not Found"
    filename = "Verification_Report.xlsx"
    workbook1.save(filename)
