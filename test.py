from openpyxl import workbook, load_workbook
import pandas as pd

dataset1 = pd.read_excel("data/Sample UPC Data.xlsx")
UPC1 = dataset1["UPC"]
# approved_Denied = dataset1['Approve/Denial Status']
# dataset1.set_index("Item ID")

workbook1 = load_workbook("data/Sample UPC Data.xlsx")
sheet1 = workbook1.active

upc = workbook1["Confirmed_UPC"]

code = int(input("Enter your code:"))

column_letter = 'P'
row_number = 2

for u in UPC1:
    if code == u:
        print("Success")
        cell = sheet1[column_letter + str(row_number)]
        cell.value = "Approved"
        filename = "Verification_Report.xlsx"
        workbook1.save(filename)
        break
      # first_element = [u]
      # print("Success")
      # id = dataset1[dataset1['Approve/Denial Status'] == code].index
      # dataset1.loc[id:"Approve/Denial Status"] == "Approved"
      # print("Approved")
      # dataset1["Approve/Denial Status"] == "Approved"
    else:
        print("Failure")
        break
