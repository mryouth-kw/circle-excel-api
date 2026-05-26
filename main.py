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

columns_to_remove = [
    "患者ID（文字列型）",
    "患者ID"
]

CSV_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1Ys9iA8aIqyAgnOoUrnu1CZT6adT2E12hSuDmaH4blXQ/"
    "export?format=csv"
)


def normalize(text):
    return str(text).replace("\ufeff", "").strip()


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
                f"{circle_id} に対応する"
                "値が見つかりません"
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

            input_path = os.path.join(
                temp_dir,
                file.filename
            )

            output_path = os.path.join(
                temp_dir,
                f"processed_{file.filename}"
            )

            # upload保存
            with open(input_path, "wb") as f:
                f.write(await file.read())

            # keep_vba=True が重要
            wb = openpyxl.load_workbook(
                input_path,
                keep_vba=True,
                data_only=False
            )

            for ws in wb.worksheets:

                target_col_index = None
                header_row_index = None

                # header探索
                for row in ws.iter_rows(
                    min_row=1,
                    max_row=10
                ):

                    headers = [
                        normalize(cell.value)
                        for cell in row
                    ]

                    for col_name in columns_to_remove:

                        if col_name in headers:

                            target_col_index = (
                                headers.index(col_name) + 1
                            )

                            header_row_index = row[0].row

                            break

                    if target_col_index:
                        break

                # ID列なし
                if not target_col_index:
                    continue

                delete_rows = []

                # データ走査
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
                # これが超重要
                for row_idx in reversed(delete_rows):
                    ws.delete_rows(row_idx, 1)

            wb.save(output_path)

            zipf.write(
                output_path,
                arcname=file.filename
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