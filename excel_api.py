import os
import tempfile
import re
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

def clean_cell_text(s: str) -> str:
    if s is None:
        return ""
    # 文字列化
    t = str(s)
    # Excel由来のリテラル改行トークンを本当の改行に
    t = t.replace("_x000D_", "\n")
    # CRLF/CR を LF に正規化
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    # タブは可視的に潰す（TSVに影響しないよう空白へ）
    t = t.replace("\t", " ")
    # 連続空白の軽減（全角空白も1個に圧縮）
    t = re.sub(r"[ \u3000]+", " ", t)
    # 行頭末の空白を削る & 空行の連続を1つに
    t = "\n".join([line.strip() for line in t.split("\n")])
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()

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

        lines = []
        for row in ws.iter_rows():
            row_vals = [clean_cell_text(cell.value) for cell in row]
            # すべて空ならスキップ
            if any(v.strip() for v in row_vals):
                lines.append("\t".join(row_vals))

        text = "\n".join(lines)

        # 連続する同一行を間引く
        deduped = []
        prev = None
        for line in text.split("\n"):
            if line != prev:
                deduped.append(line)
            prev = line
        text = "\n".join(deduped)

        return jsonify({
            "sheet": ws.title,
            "text": text
        })
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
