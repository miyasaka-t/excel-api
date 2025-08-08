# excel_api.py
import os
from flask import Flask, request, jsonify, Response
from openpyxl import load_workbook

app = Flask(__name__)

# 上限（必要に応じて調整）
MAX_ROWS = 200
MAX_COLS = 50
MAX_NONEMPTY = 2000  # 非空セルの最大数（安全弁）

def to_str(v):
    if v is None:
        return ""
    s = str(v)
    # Excelの _x000D_ 改行トークンや制御文字を潰して肥大化＆化け防止
    return (
        s.replace("_x000D_", " ")
         .replace("\t", " ")
         .replace("\r\n", " ")
         .replace("\n", " ")
         .replace("\r", " ")
         .strip()
    )

@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "message": "excel-api (sparse non-empty, inline/bom switch)", "endpoint": "/extract"})

@app.route("/extract", methods=["POST"])
def extract():
    """
    返却形式: 非空セルのみを 'A1<TAB>値' の1行フォーマットで列挙。
    切替パラメータ（multipart/form-data のフィールド）:
      - file: (必須) アップロードする .xlsx
      - bom: "true"/"false" 既定: true
      - inline: "true"/"false" 既定: true
        inline=true  → 本文で返す（LLM向け）
        inline=false → 添付(attachment)で返す（Excel/メモ帳配布向け）
      - sheet: 省略可。シート名 or "0"/"1"...（0/1始まりどちらでも可）
    """
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "file is required (multipart/form-data)"}), 400

    # オプション
    bom_on = (request.form.get("bom", "true").lower() != "false")
    inline_on = (request.form.get("inline", "true").lower() != "false")
    sheet_req = request.form.get("sheet")

    try:
        # 軽量に読む / 数式は値に
        wb = load_workbook(f, data_only=True, read_only=True)
    except Exception as e:
        return jsonify({"error": f"failed to read workbook: {e}"}), 400

    # シート解決（名前/0始まり/1始まりに対応）
    ws = None
    if sheet_req:
        try:
            idx = int(sheet_req)
            names = wb.sheetnames
            if 0 <= idx < len(names):
                ws = wb[names[idx]]          # 0-based
            elif 1 <= idx <= len(names):
                ws = wb[names[idx - 1]]      # 1-based
        except ValueError:
            if sheet_req in wb.sheetnames:
                ws = wb[sheet_req]
    ws = ws or wb.active

    # スパース出力（非空セルのみ）
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

    payload = "\n".join(lines)
    if bom_on:
        payload = "\ufeff" + payload  # UTF-8 BOM 付与（Excel/メモ帳向け）

    headers = {}
    # inline=false（配布用）のときだけ attachment で返す
    if not inline_on:
        headers["Content-Disposition"] = 'attachment; filename="extract.tsv"'

    # LLM向けは text/plain の方が扱いやすい
    return Response(payload, mimetype="text/plain; charset=utf-8", headers=headers)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
