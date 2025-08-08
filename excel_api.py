# app.py
import os
from flask import Flask, request, jsonify, Response
from openpyxl import load_workbook

app = Flask(__name__)

# ざっくり上限（必要なら増やす）
MAX_ROWS = 200
MAX_COLS = 50
MAX_NONEMPTY = 2000  # 返す非空セルの最大数（安全弁）

def to_str(v):
    if v is None:
        return ""
    s = str(v)
    return s.replace("\t", " ").replace("\r\n", " ").replace("\n", " ").replace("\r", " ").strip()

@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "message": "excel-api (sparse non-empty)", "endpoint": "/extract"})

@app.route("/extract", methods=["POST"])
def extract():
    """
    返却: 非空セルだけを 'A1<TAB>値' 形式で1行ずつ。
    空セルは出さないので“横に同じ文字が伸びる”現象は起きません。
    """
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "file is required (multipart/form-data)"}), 400

    try:
        wb = load_workbook(f, data_only=True, read_only=True)
    except Exception as e:
        return jsonify({"error": f"failed to read workbook: {e}"}), 400

    ws = wb.active  # 最初のシート
    lines = []
    count = 0

    for row in ws.iter_rows(min_row=1, max_row=MAX_ROWS, min_col=1, max_col=MAX_COLS, values_only=False):
        for cell in row:
            v = cell.value
            if v is None:
                continue
            txt = to_str(v)
            if not txt:
                continue
            lines.append(f"{cell.coordinate}\t{txt}")
            count += 1
            if count >= MAX_NONEMPTY:
                lines.append("# ...truncated...")
                break
        if count >= MAX_NONEMPTY:
            break

    return Response("\n".join(lines), mimetype="text/plain; charset=utf-8")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
