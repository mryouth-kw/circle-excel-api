from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

import openpyxl
import pandas as pd
import zipfile
import tempfile
import os
import csv
import requests

from io import BytesIO

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

columns_to_search = [
    "患者ID（文字列型）",
    "患者ID"
]

CSV_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1Ys9iA8aIqyAgnOoUrnu1CZT6adT2E12hSuDmaH4blXQ/"
    "export?format=csv"
)

# =========================
# キャッシュ
# =========================
mapping_cache = None


def normalize(text):

    if text is None:
        return ""

    text = (
        str(text)
        .replace("\ufeff", "")
        .replace("\u3000", " ")
        .replace("\r", "")
        .replace("\n", "")
        .strip()
    )

    # 12345.0 → 12345
    try:

        num = float(text)

        if num.is_integer():
            return str(int(num))

    except:
        pass

    return text


# =========================
# Google Sheets 読込
# 初回のみアクセス
# =========================
def load_mapping():

    global mapping_cache

    # キャッシュ利用
    if mapping_cache is not None:
        return mapping_cache

    print("LOAD GOOGLE SHEET")

    res = requests.get(
        CSV_URL,
        timeout=15
    )

    res.encoding = "utf-8"

    reader = csv.reader(
        res.text.splitlines()
    )

    mapping_cache = {}

    for row in reader:

        if len(row) < 2:
            continue

        key = normalize(row[0])
        value = normalize(row[1])

        if key:
            mapping_cache[key] = value

    print(
        f"MAPPING COUNT: {len(mapping_cache)}"
    )

    return mapping_cache


def get_target_value(circle_id):

    mapping = load_mapping()

    return mapping.get(
        normalize(circle_id)
    )


def get_target_sheet(wb):
    """
    仕様:
    ・1枚目が「検索情報」なら2枚目
    ・それ以外なら1枚目
    """

    sheetnames = wb.sheetnames

    if len(sheetnames) == 0:
        return None

    first_sheet = sheetnames[0]

    if normalize(first_sheet) == "検索情報":

        if len(sheetnames) >= 2:
            return wb[sheetnames[1]]

        return None

    return wb[first_sheet]


# =========================
# xls → xlsx変換
# =========================
def convert_xls_to_xlsx(
    xls_path,
    xlsx_path
):

    excel_file = pd.ExcelFile(
        xls_path,
        engine="xlrd"
    )

    with pd.ExcelWriter(
        xlsx_path,
        engine="openpyxl"
    ) as writer:

        for sheet_name in excel_file.sheet_names:

            df = pd.read_excel(
                xls_path,
                sheet_name=sheet_name,
                engine="xlrd",
                header=None
            )

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                header=False,
                index=False
            )


@app.get("/")
def root():
    return {"message": "API OK"}


@app.post("/process")
async def process_excel(
    files: list[UploadFile] = File(...),
    circle_id: str = Form(...),
    visit: str = Form(...)
):

    target_value = get_target_value(circle_id)

    print("TARGET VALUE:", target_value)

    if not target_value:
        return {
            "error": (
                f"{circle_id} に対応する値が見つかりません"
            )
        }

    temp_dir = tempfile.mkdtemp()

    zip_buffer = BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=1
    ) as zipf:

        for file in files:

            try:

                print("=" * 50)
                print("FILE:", file.filename)

                input_path = os.path.join(
                    temp_dir,
                    file.filename
                )

                safe_target_value = (
                    str(target_value)
                    .replace("/", "_")
                    .replace("\\", "_")
                    .replace(":", "_")
                    .replace("*", "_")
                    .replace("?", "_")
                    .replace('"', "_")
                    .replace("<", "_")
                    .replace(">", "_")
                    .replace("|", "_")
                )

                output_path = os.path.join(
                    temp_dir,
                    f"processed_{file.filename}"
                    f"{safe_target_value}_{file.filename}"
                )

                # 保存
                with open(input_path, "wb") as f:
                    f.write(await file.read())

                ext = os.path.splitext(
                    file.filename
                )[1].lower()

                # =========================
                # xls
                # =========================
                if ext == ".xls":

                    print("XLS MODE")

                    converted_path = os.path.join(
                        temp_dir,
                        f"converted_{os.path.basename(file.filename)}x"
                    )

                    convert_xls_to_xlsx(
                        input_path,
                        converted_path
                    )

                    wb = openpyxl.load_workbook(
                        converted_path,
                        data_only=False
                    )

                # =========================
                # xlsx / xlsm
                # =========================
                else:

                    print("XLSX MODE")

                    wb = openpyxl.load_workbook(
                        input_path,
                        keep_vba=True,
                        data_only=False
                    )

                print(
                    "SHEETS:",
                    wb.sheetnames
                )

                # 対象シート取得
                ws = get_target_sheet(wb)

                # 対象シートなし
                if ws is None:

                    print("TARGET SHEET NONE")

                    wb.save(output_path)

                    zipf.write(
                        output_path,
                        arcname=file.filename
                        arcname=os.path.basename(output_path)
                    )

                    continue

                print(
                    "TARGET SHEET:",
                    ws.title
                )

                target_col_index = None
                header_row_index = None

                # ヘッダー探索
                for row in ws.iter_rows(
                    min_row=1,
                    max_row=min(10, ws.max_row)
                ):

                    headers = [
                        normalize(cell.value)
                        for cell in row
                    ]

                    for col_name in columns_to_search:

                        if col_name in headers:

                            target_col_index = (
                                headers.index(col_name) + 1
                            )

                            header_row_index = row[0].row

                            break

                    if target_col_index:
                        break

                print(
                    "TARGET COLUMN:",
                    target_col_index
                )

                # ID列なし
                if not target_col_index:

                    print(
                        "COLUMN NOT FOUND"
                    )

                    wb.save(output_path)

                    zipf.write(
                        output_path,
                        arcname=file.filename
                        arcname=os.path.basename(output_path)
                    )

                    continue

                delete_rows = []

                matched_count = 0

                # 行判定
                for row_idx in range(
                    header_row_index + 1,
                    ws.max_row + 1
                ):

                    raw_value = ws.cell(
                        row=row_idx,
                        column=target_col_index
                    ).value

                    cell_value = normalize(
                        raw_value
                    )

                    # 最初の10行だけログ
                    if row_idx <= (
                        header_row_index + 10
                    ):

                        print(
                            row_idx,
                            "RAW:",
                            raw_value,
                            "NORMALIZED:",
                            cell_value
                        )

                    if cell_value == target_value:
                        matched_count += 1
                    else:
                        delete_rows.append(row_idx)

                print(
                    "MATCHED:",
                    matched_count
                )

                print(
                    "DELETE:",
                    len(delete_rows)
                )

                # 後ろから削除
                for row_idx in reversed(delete_rows):
                    ws.delete_rows(row_idx, 1)

                wb.save(output_path)

                zipf.write(
                    output_path,
                    arcname=file.filename
                    arcname=os.path.basename(output_path)
                )

                print(
                    "SAVE OK:",
                    file.filename
                )

            except Exception as e:

                print(
                    f"ERROR: {file.filename}: {str(e)}"
                )

    zip_buffer.seek(0)

    zip_filename = (
        f"調査票2_{circle_id}_Visit-{visit}.zip"
    )

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition":
            f'attachment; filename="{zip_filename}"'
        }