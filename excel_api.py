from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import PlainTextResponse
import pandas as pd
from typing import Optional, List

app = FastAPI(title="Excel→TSV API", version="1.0.0")

def df_to_tsv(df: pd.DataFrame) -> str:
    # すべて文字列化・NaN除去
    df = df.astype(str).fillna("")
    # セル内タブ/改行をTSV安全に
    df = df.applymap(lambda x: x.replace("\t", " ").replace("\r\n", "\n").replace("\r", "\n"))
    # TSV出力
    return df.to_csv(sep="\t", index=False, header=True, line_terminator="\n")

@app.post("/extract", response_class=PlainTextResponse)
async def extract_tsv(
    file: UploadFile = File(..., description="Excelファイル（.xlsx, .xls）"),
    sheet: Optional[str] = Query(None, description="読み込むシート名。未指定なら全シート"),
    include_sheet_headers: bool = Query(True, description="複数シート結合時、シート見出し行を付けるか"),
) -> PlainTextResponse:
    try:
        content_type = (file.content_type or "").lower()
        if not (file.filename.endswith((".xlsx", ".xls")) or "excel" in content_type):
            raise HTTPException(400, "Excelファイル（.xlsx/.xls）をアップロードしてください。")

        # ファイルを直接読み込む（メモリ節約のためUploadFile.fileを渡す）
        if sheet:
            # 特定シートのみ
            df = pd.read_excel(file.file, sheet_name=sheet, dtype=str, engine=None)
            tsv = df_to_tsv(df)
            return PlainTextResponse(tsv, media_type="text/plain; charset=utf-8")

        # 全シート
        file.file.seek(0)
        xls = pd.ExcelFile(file.file, engine=None)
        parts: List[str] = []
        for name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=name, dtype=str)
            if include_sheet_headers:
                parts.append(f"# sheet: {name}")
            parts.append(df_to_tsv(df).rstrip("\n"))
        tsv_all = "\n".join(parts) + "\n"
        return PlainTextResponse(tsv_all, media_type="text/plain; charset=utf-8")

    except ValueError as e:
        # pandasがシート名を見つけられない等
        raise HTTPException(400, f"読み込みエラー: {e}")
    except Exception as e:
        raise HTTPException(500, f"サーバーエラー: {e}")
