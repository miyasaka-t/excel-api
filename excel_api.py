# app.py
import os
from flask import Flask, request, jsonify, Response
from openpyxl import load_workbook

app = Flask(__name__)

# ここだけ変えれば上限を調整できる
MAX_ROWS = 50
MAX_COLS = 20

def to_str(v):
    if v is None:
        return ""
    s = str(v)
    # タブ/改行を潰して行崩れ・肥大化防止
    return s.replace("\t", " ").replace("\r\n", " ").replace("\n", " ").replace("\r", " ").strip()

@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "message": "excel-api (simple)", "endpoint": "/extract"})

@app.route("/extract", methods=["POST"])
def extract():
    """
    使い方: curl -X POST https://<your-app>.onrender.com/extract -F "file=@C:/path/to/file.xlsx"
    返却: 先頭 50行×20列 の TSV（固定）
    """
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "file is required (multipart/form-data)"}), 400

    try:
        # read_only=True で軽量に読む / data_only=True で数式は値に
        wb = load_workbook(f, data_only=True, read_only=True)
    except Exception as e:
        return jsonify({"error": f"failed to read workbook: {e}"}), 400

    ws = wb.active  # 最初のシートだけ
    lines = []

    # 先頭 50x20 だけを values_only で取得（マージ展開しない）
    for row in ws.iter_rows(min_row=1, max_row=MAX_ROWS, min_col=1, max_col=MAX_COLS, values_only=True):
        lines.append("\t".join(to_str(v) for v in row))

    tsv = "\n".join(lines)
    return Response(tsv, mimetype="text/plain; charset=utf-8")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
