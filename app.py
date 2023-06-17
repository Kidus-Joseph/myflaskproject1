from flask import Flask, render_template, request

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
        return "Data has been Updated Successfully in Verification_Report.excel,!! Check it out!!"
    return render_template('form.html')


if __name__ == "__main__":
    app.run(debug=True)

# if x.all() == upc:
    #     approved_Denied == "Approved"
    #     dataset.to_excel(filename)
    # else:
    #     approved_Denied == "Denied"
    #     denial_Reason == 'Invalid UPC'
    #     dataset.to_excel(filename)
    # approved_Denied = dataset['Approve/Denial Status']
    # denial_Reason = dataset['Denial Reason']

# from flask import Flask, jsonify, request
# import requests
# import json


# app = Flask(__name__)


# @app.route('/api/tag/epc/30F4257BF46DB64000000190?apikey=JbhZI7fNiMnNsS5t', methods=['POST'])
# def gs1_endpoint():
#     data = request.get_json()
#     # Do something with the data
#     print(data)

#     response = {
#         'status': 'success',
#         'message': 'Data processed successfully',
#     }

#     return jsonify(response), 200


# if __name__ == 'main':
#     app.run(port=5000)

# url = "https://gs1-eu1-pd-rfidcoder-app.azurewebsites.net/api/tag/epc/30F4257BF46DB64000000190?apikey=JbhZI7fNiMnNsS5t"

# data = {
#     'upc': '622356532419',
#     'model_#': 'BL770'
# }

# response = requests.post(url, data=json.dumps(data), headers={
#                          'Content-Type': 'application/json'})

# if response.status_code == 200:
#     print(response.json())
# else:
#     print('Error:', response.status_code)
