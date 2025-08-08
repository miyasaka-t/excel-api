from flask import Flask, request, jsonify
from openpyxl import load_workbook
import tempfile

app = Flask(__name__)

@app.route('/extract', methods=['POST'])
def extract():
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "ファイルが見つかりません"}), 400

    # 一時ファイルとして保存
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        file.save(tmp.name)
        wb = load_workbook(tmp.name, data_only=True)
        ws = wb.active

        result = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    result.append(f"({cell.row},{cell.column}): \"{str(cell.value)}\"")

    return jsonify({"text": "\n".join(result)})

if __name__ == '__main__':
    print("✅ Flask API starting...")
    app.run(debug=True, port=5000)
