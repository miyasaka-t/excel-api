import os
from flask import Flask, request, jsonify
import openpyxl

app = Flask(__name__)

@app.route("/")
def home():
    return "Flask Excel API is running!"

# 実際のAPI処理の例
@app.route("/extract", methods=["POST"])
def extract():
    file = request.files["file"]
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(values_only=True):
        data[row[0]] = row[1]
    return jsonify(data)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Render用PORT取得
    app.run(host="0.0.0.0", port=port, debug=True)
