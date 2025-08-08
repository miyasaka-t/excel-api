import io
import os
import json
from datetime import datetime, date
from typing import Any

from flask import Flask, request, Response, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

def to_str(v: Any) -> str:
    """TSV安全化して文字列化（タブ/改行はスペースに）"""
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.isoformat(sep=" ")
    s = str(v)
    return s.replace("\t", " ").replace("\r\n", " ").replace("\n", " ").replace("\r", " ").strip()

def sheet_to_grid(ws):
    """マージを展開した値で 2D グリッドを作る"""
    max_r = ws.max_row or 1
    max_c = ws.max_column or 1

    # まずは素の値を格納
    grid = [[None for _ in range(max_c)] for _ in range(max_r)]
    for r in range(1, max_r + 1):
        for c in range(1, max_c + 1):
            cell = ws.cell(row=r, column=c)
            grid[r-1][c-1] = cell.value

    # マージセルを値で埋める（元のシートは変更しない）
    for mr in ws.merged_cells.ranges:
        v = ws.cell(row=mr.min_row, column=mr.min_col).value
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                grid[r-1][c-1] = v

    return grid

def grid_to_tsv(grid):
    lines = []
    for row in grid:
        lines.append("\t".join(to_str(v) for v in row))
    return "\n".join(lines)

def grid_to_json(grid):
    # 行配列の配列（プレーン）で返す
    return [[to_str(v) for v in row] for row in grid]

@app.route("/", methods=["GET"])
def health():
    return jsonify({"ok": True, "message": "excel-api is up", "endpoints": ["/extract"]})

@app.route("/extract", methods=["POST"])
def extract():
    """
    フォーム: file=@xxx.xlsx
    任意: sheet=シート名 or 0始まり/1始まりどちらでも通す
    任意: format=tsv|json (デフォルト tsv)
    """
    if "file" not in request.files:
        return jsonify({"error": "file is required (multipart/form-data)"}), 400

    f = request.files["file"]
    data = f.read()
    if not data:
        return jsonify({"error": "empty file"}), 400

    # シート選択
    sheet_req = request.form.get("sheet")
    fmt = (request.form.get("format") or "tsv").lower()

    try:
        wb = load_workbook(io.BytesIO(data), data_only=True, read_only=False)
    except Exception as e:
        return jsonify({"error": f"failed to read workbook: {e}"}), 400

    # sheet の解決
    ws = None
    if sheet_req:
        # index でも名前でもOK
        try:
            # 数値っぽければ index として試す（0/1始まりどっちでも）
            idx = int(sheet_req)
            if idx < 0:
                # 0始まり負数は無効
                raise ValueError
            if idx < len(wb.sheetnames):
                ws = wb[wb.sheetnames[idx]]  # 0始まりとして解釈
            else:
                # 1始まりとして再解釈
                if 1 <= idx <= len(wb.sheetnames):
                    ws = wb[wb.sheetnames[idx - 1]]
        except ValueError:
            # 文字列名として解決
            if sheet_req in wb.sheetnames:
                ws = wb[sheet_req]
    if ws is None:
        # 指定なし or 解決失敗 → 最初のシート
        ws = wb[wb.sheetnames[0]]

    grid = sheet_to_grid(ws)

    if fmt == "json":
        payload = {
            "sheet": ws.title,
            "rows": grid_to_json(grid),
        }
        return Response(json.dumps(payload, ensure_ascii=False), mimetype="application/json; charset=utf-8")

    # 既定は TSV
    tsv = grid_to_tsv(grid)
    # 文字化け回避のため UTF-8 + BOM をつけることもあるが、ここは素のUTF-8で返す
    return Response(tsv, mimetype="text/plain; charset=utf-8")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
