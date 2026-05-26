from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

import openpyxl
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


def normalize(text):
    if text is None:
        return ""

    return (
        str(text)
        .replace("\ufeff", "")
        .replace("\u3000", " ")
        .strip()
    )


def get_target_value(circle_id):

    res = requests.get(CSV_URL, timeout=15)
    res.encoding = "utf-8"

    reader = csv.reader(res.text.splitlines())

    normalized_id = normalize(circle_id)

    for row in reader:

        if len(row) < 2:
            continue

        if normalize(row[0]) == normalized_id:
            return normalize(row[1])

    return None


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

                input_path = os.path.join(
                    temp_dir,
                    file.filename
                )

                output_path = os.path.join(
                    temp_dir,
                    f"processed_{file.filename}"
                )

                # 保存
                with open(input_path, "wb") as f:
                    f.write(await file.read())

                # フォーマット保持重視
                wb = openpyxl.load_workbook(
                    input_path,
                    keep_vba=True,
                    data_only=False
                )

                # 対象シート取得
                ws = get_target_sheet(wb)

                # 対象シートなし
                if ws is None:

                    wb.save(output_path)

                    zipf.write(
                        output_path,
                        arcname=file.filename
                    )

                    continue

                target_col_index = None
                header_row_index = None

                # ヘッダー探索
                # 最大10行まで
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

                # ID列が見つからない
                if not target_col_index:

                    wb.save(output_path)

                    zipf.write(
                        output_path,
                        arcname=file.filename
                    )

                    continue

                delete_rows = []

                # 行判定
                # openpyxl.cellアクセスを最小化
                for row_idx in range(
                    header_row_index + 1,
                    ws.max_row + 1
                ):

                    cell_value = normalize(
                        ws.cell(
                            row=row_idx,
                            column=target_col_index
                        ).value
                    )

                    if cell_value != target_value:
                        delete_rows.append(row_idx)

                # 後ろから削除
                # フォーマット維持に最も安全
                for row_idx in reversed(delete_rows):
                    ws.delete_rows(row_idx, 1)

                wb.save(output_path)

                zipf.write(
                    output_path,
                    arcname=file.filename
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
    )