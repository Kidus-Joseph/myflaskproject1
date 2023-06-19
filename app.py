import os

from flask import Flask, render_template, request, send_file, make_response

from fileinput import filename

import pandas as pd

from openpyxl import workbook, load_workbook

app = Flask(__name__)


@app.route('/', methods=["GET", "POST"])
def form():
    if request.method == "POST" and "z" in request.form:
        x = request.form.get('z')
        x = int(x)
        dataset = pd.read_excel("data/Sample UPC Data.xlsx")
        upc1 = dataset["UPC"]
        workbook1 = load_workbook("data/Sample UPC Data.xlsx")
        sheet1 = workbook1.active
        upc = workbook1["Confirmed_UPC"]
        column_letter = 'P'
        row_number = dataset[dataset["UPC"] == x].index.to_list()
        row_number = str(row_number)[1:-1]
        row_number = int(row_number) + 2
        filename = "Verification_Report.xlsx"
        for u in upc1:
            if x == u:
                cell = sheet1[column_letter + str(row_number)]
                cell.value = "Approved"
                filename = "Verification_Report.xlsx"
                workbook1.save(filename)
        return "Data has been Updated Successfully in Verification_Report.excel!! Check it out!!"
    return render_template('verification.html')


@app.route('/uploadVerification')
def uploadVerification():
    return render_template('verification.html')


@app.route('/download')
def download():
    file_path = r"C:\Users\kidus\myflaskproject\data\UPCUploadTemplate.xlsx"
    filename = "UPCUploadTemplate.xlsx"
    # Create the response object
    response = make_response(send_file(file_path))
    # Set the "Content-Disposition" header to trigger file download
    response.headers["Content-Disposition"] = "attachment; filename=" + filename
    # Get the user's home directory
    home_dir = os.path.expanduser("~")
    # Set the path to the "Downloads" folder
    downloads_folder = os.path.join(home_dir, "Downloads")
    # Set the "X-Sendfile" header to suggest the default save location
    response.headers["X-Sendfile"] = os.path.join(downloads_folder, filename)
    return response


@app.route('/success', methods=['POST'])
def success():
    if request.method == 'POST':
        f = request.files['file']
        f.save(f.filename)
        uploadedFile = f
        batchExcel_1 = pd.read_excel(uploadedFile)
        batchColumn_2 = set(batchExcel_1.iloc[:, 1])

        sampleExcel_2 = pd.read_excel("data/Sample UPC Data.xlsx")
        sampleColumn_2 = set(sampleExcel_2.iloc[:, 1])

        workbook1 = load_workbook(uploadedFile)
        sheet1 = workbook1.active

        upc = workbook1["Confirmed_UPC"]

        match_found = False
        for value in batchColumn_2:
            column_letter = 'P'
            column_letter1 = 'Q'
            if value in sampleColumn_2:
                match_found = True
                # print(f"UPC '{value}' found in both Excel files.")
                row_number = batchExcel_1[batchExcel_1["UPC"]
                                          == value].index.to_numpy()
                row_number = str(row_number)[1:-1]
                row_number = int(row_number) + 2
                cell = sheet1[column_letter + str(row_number)]
                cell.value = "Approved"
                outputFileName = "Verification_Report2.xlsx"
                workbook1.save(outputFileName)
            else:
                # print(f"UPC '{value}' invalid")
                row_number = batchExcel_1[batchExcel_1["UPC"]
                                          == value].index.to_numpy()
                row_number = str(row_number)[1:-1]
                row_number = int(row_number) + 2
                cell = sheet1[column_letter + str(row_number)]
                cell.value = "Denied"
                cell = sheet1[column_letter1 + str(row_number)]
                cell.value = "UPC Invalid/Not Found"
                outputFileName = "Verification_Report2.xlsx"
                workbook1.save(outputFileName)

        if not match_found:
            print("No values found in the second Excel file.")

        return render_template('checkout.html')
    return render_template('verification.html')


if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000, server_name='GS1 Scout')
