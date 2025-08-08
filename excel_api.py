import os
import tempfile
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

def expand_merged_cells(ws):
    """
    結合セルを展開して全セルに同じ値を入れる（書き込み可能化のため先にunmerge）
    """
    ranges = list(ws.merged_cells.ranges)
    # 値と範囲を控える
    snapshot = []
    for rng in ranges:
        min_row, min_col, max_row, max_col = rng.bounds
        top_left = ws.cell(min_row, min_col).value
        snapshot.append((str(rng), (min_row, min_col, max_row, max_col), top_left))

    # 先に全部unmerge
    for rng_str, _, _ in snapshot:
        ws.unmerge_cells(rng_str)

    # 値を展開
    for _, (min_row, min_col, max_row, max_col), val in snapshot:
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(r, c, value=val)

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

    suffix = ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        f.save(tmp.name)
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path, data_only=True)
        ws = wb.active

        # 結合セル展開
        expand_merged_cells(ws)

        lines = []
        for row in ws.iter_rows():
            row_data = []
            for cell in row:
                if cell.value is not None:
                    val = str(cell.value).replace("\r\n", "\n").replace("\r", "\n")
                    row_data.append(val.strip())
            if row_data:
                lines.append("\t".join(row_data))

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
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
