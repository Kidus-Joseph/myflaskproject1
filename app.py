from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)


@app.route('/')
def index():
    return render_template('index.html')


if __name__ == "__main__":
    app.run(debug=True)

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
