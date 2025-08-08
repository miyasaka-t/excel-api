import os
import tempfile
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/")
def home():
    return "Flask Excel API is running!"

@app.route("/health")
def health():
    return "ok", 200

@app.route("/extract", methods=["POST"])
def extract():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "missing 'file' form field"}), 400

    # 一時ファイルに保存（ファイルライクのままでも可だが安定性重視）
    suffix = ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        f.save(tmp.name)
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path, data_only=True)
        ws = wb.active  # 必要なら request.args でシート名を受け取る拡張も可

        lines = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    # openpyxlの列番号は 1 始まり
                    lines.append(f'({cell.row},{cell.column}): "{str(cell.value)}"')

        return jsonify({
            "sheet": ws.title,
            "text": "\n".join(lines)
        })
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass

if __name__ == "__main__":
    # Render では PORT 環境変数で受ける
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
