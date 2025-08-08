# excel_api.py
import os
import re
import json
import tempfile
from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

# ==========
# ユーティリティ
# ==========
HEADER_SYNONYMS = {
    # 同義語マップ（必要に応じて拡張）
    "技術者名": ["氏名", "名前", "技術者名", "技術者 氏名"],
    "性別": ["性別", "Gender"],
    "年齢": ["年齢", "歳", "Age"],
    "最寄駅": ["最寄駅", "最寄り駅", "Nearest Station"],
    "資格": ["資格", "保有資格"],
    "学歴": ["学歴", "最終学歴"],
    "得意分野": ["得意分野"],
    "得意技術": ["得意技術", "スキル", "スキルセット", "技術"],
    "得意業務": ["得意業務"],
    "自己PR": ["自己PR", "自己紹介", "自己 アピール"],
    "期間": ["期間", "在籍期間"],
    "業務内容": ["業務内容", "仕事内容"],
    "役割": ["役割", "ポジション", "役割・規模"],
    "言語/OS/環境": ["言語/OS/環境", "言語", "環境", "OS", "ツール"],
}

# 見出しっぽさの判断（日本語コロン含む）
HEADER_LIKE = re.compile(r"[：:\uFF1A]\s*$")  # 末尾が : or ：
BR_TAGS = re.compile(r"(_x000D_|\r\n|\r|\n)")

def norm_text(v):
    """セル内容のノイズ除去＆正規化"""
    if v is None:
        return None
    s = str(v)
    # 改行ノイズ除去
    s = BR_TAGS.sub("\n", s)
    # 全角スペース正規化
    s = s.replace("\u3000", " ")
    # 連続空白の圧縮（改行は保持）
    s = re.sub(r"[ \t]+", " ", s)
    # 端の空白削除
    s = s.strip()
    return s if s else None

def expand_merged(ws):
    """マージセルを左上値で全セルへ展開（openpyxlは見た目上マージでも値は左上のみ持つため）"""
    for rng in ws.merged_cells.ranges:
        min_row, min_col, max_row, max_col = rng.min_row, rng.min_col, rng.max_row, rng.max_col
        top_left = ws.cell(min_row, min_col).value
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(r, c).value = top_left

def sheet_to_cells(ws):
    """セル配列（LLMに食わせやすい一次情報）"""
    cells = []
    for row in ws.iter_rows():
        for cell in row:
            v = norm_text(cell.value)
            if v is not None:
                cells.append({"r": cell.row, "c": cell.column, "v": v})
    return cells

def synonym_score(header_text):
    """ヘッダ同義語マッチの加点。最大 0.3 くらい。"""
    if not header_text:
        return 0.0
    score = 0.0
    for canonical, alts in HEADER_SYNONYMS.items():
        for alias in alts + [canonical]:
            if alias == header_text or (alias in header_text and len(alias) >= 2):
                score = max(score, 0.3 if alias == header_text else 0.15)
    return score

def is_headerish(text):
    """文字列が見出しっぽいか（末尾コロン、短め、日本語名詞っぽさなど）"""
    if not text:
        return False
    if HEADER_LIKE.search(text):
        return True
    # 長すぎる文章は見出しでない可能性が高い
    if len(text) <= 20 and re.search(r"[一-龥ぁ-んァ-ンA-Za-z0-9]", text):
        # “～について”のような名詞句をゆるく許容
        return True
    return False

def distance(r1, c1, r2, c2):
    return abs(r1 - r2) + abs(c1 - c2)

def find_kv_pairs(cells, search_right=6, search_down=4):
    """
    キー→値の候補をヒューリスティックで抽出。
    - キー候補：左端/上端/末尾コロン/短文 など
    - 値探索：右方向（同じ行で近い列）優先、無ければ下方向（同じ列で近い行）
    - 信頼度：近さ、同義語一致、見出し記号 の合成
    """
    # 索引用
    by_row = {}
    by_col = {}
    for cell in cells:
        by_row.setdefault(cell["r"], []).append(cell)
        by_col.setdefault(cell["c"], []).append(cell)
    for row in by_row.values():
        row.sort(key=lambda x: x["c"])
    for col in by_col.values():
        col.sort(key=lambda x: x["r"])

    pairs = []
    used_values = set()  # 値のセル重複を軽減（適度に）

    for cell in cells:
        k_text = cell["v"]
        if not is_headerish(k_text):
            continue

        base_conf = 0.2
        if HEADER_LIKE.search(k_text or ""):
            base_conf += 0.2
            k_text = k_text.rstrip("：:")  # コロン除去

        base_conf += synonym_score(k_text)

        r, c = cell["r"], cell["c"]

        # 1) 右方向優先（同一行）
        cand_value = None
        conf = base_conf
        if r in by_row:
            for neighbor in by_row[r]:
                if neighbor["c"] <= c:
                    continue
                if neighbor["c"] - c > search_right:
                    break
                if neighbor["v"] and (neighbor["r"], neighbor["c"]) not in used_values:
                    cand_value = neighbor
                    d = neighbor["c"] - c
                    conf_r = max(0.0, 0.4 - 0.06 * (d - 1))  # 近いほど高い
                    conf = base_conf + conf_r
                    break

        # 2) 無ければ下方向（同一列）
        if cand_value is None and c in by_col:
            for neighbor in by_col[c]:
                if neighbor["r"] <= r:
                    continue
                if neighbor["r"] - r > search_down:
                    break
                if neighbor["v"] and (neighbor["r"], neighbor["c"]) not in used_values:
                    cand_value = neighbor
                    d = neighbor["r"] - r
                    conf_d = max(0.0, 0.35 - 0.07 * (d - 1))
                    conf = base_conf + conf_d
                    break

        if cand_value:
            used_values.add((cand_value["r"], cand_value["c"]))
            pairs.append({
                "key": k_text,
                "key_cell": {"r": r, "c": c},
                "value": cand_value["v"],
                "value_cell": {"r": cand_value["r"], "c": cand_value["c"]},
                "confidence": round(min(conf, 0.95), 2),
                "strategy": "right-then-down"
            })

    # 重複キーの圧縮（高conf優先）
    dedup = {}
    for p in pairs:
        k = p["key"]
        if k not in dedup or p["confidence"] > dedup[k]["confidence"]:
            dedup[k] = p
    return list(dedup.values())

# ==========
# ルーティング
# ==========
@app.route("/")
def home():
    return "Flask Excel API is running!"

@app.route("/health")
def health():
    return "ok", 200

@app.route("/extract", methods=["POST"])
def extract():
    """
    multipart/form-data で `file`（.xlsx）を送ると、
    - cells: セルの一次情報（r,c,v）
    - pairs: ヒューリスティックなキー→値候補（confidence付き）
    - sheet: 対象シート名
    を返す。LLMには pairs を優先使用、欠落は cells から補完が推奨。
    """
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "missing 'file' form field"}), 400

    suffix = ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        f.save(tmp.name)
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path, data_only=True)
        ws = wb.active  # 必要なら ?sheet=xxx で指定に拡張してOK
        expand_merged(ws)

        cells = sheet_to_cells(ws)
        pairs = find_kv_pairs(cells)

        return jsonify({
            "sheet": ws.title,
            "pairs": pairs,           # LLMはまずこれを信頼
            "cells": cells,           # 欠落・補完・監査用（そのままLLMに渡してもOK）
            "meta": {
                "rows": ws.max_row,
                "cols": ws.max_column,
                "note": "pairsはヒューリスティック（右→下探索＋同義語加点）。欠落はcellsで補完してね。"
            }
        })
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass

if __name__ == "__main__":
    # Render は PORT 指定が来る
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
