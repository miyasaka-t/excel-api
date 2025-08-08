import os
from flask import Flask, request, jsonify, Response
from openpyxl import load_workbook

app = Flask(__name__)

MAX_ROWS = 200
MAX_COLS = 50
MAX_NONEMPTY = 2000  # 安全弁

def to_str(v):
    if v is None:
        return ""
    s = str(v)
    return (
        s.replace("_x000D_", " ")  # Excel改行トークン除去
         .replace("\t", " ")
         .replace("\r\n", " ")
         .replace("\n", " ")
         .replace("\r", " ")
         .strip()
    )

@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "message": "excel-api (sparse non-empty, BOM)", "endpoint": "/extract"})

@app.route("/extract", methods=["POST"])
def extract():
    """
    非空セルだけ 'A1<TAB>値' 形式で返す。
    UTF-8 BOM付きなのでExcelやメモ帳で文字化けしない。
    """
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "file is required (multipart/form-data)"}), 400

    try:
        wb = load_workbook(f, data_only=True, read_only=True)
    except Exception as e:
        return jsonify({"error": f"failed to read workbook: {e}"}), 400

    ws = wb.active
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

    # UTF-8 BOM を先頭に付与
    bom_tsv = "\ufeff" + "\n".join(lines)
    return Response(
        bom_tsv,
        mimetype="text/tab-separated-values; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="extract.tsv"'}
    )

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
