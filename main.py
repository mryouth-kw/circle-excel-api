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

CSV_URL = "https://docs.google.com/spreadsheets/d/1Ys9iA8aIqyAgnOoUrnu1CZT6adT2E12hSuDmaH4blXQ/export?format=csv"


def normalize(text):
    return str(text).replace("\ufeff", "").strip()


def get_target_value(circle_id):

    res = requests.get(CSV_URL)
    res.encoding = "utf-8"

    reader = csv.reader(res.text.splitlines())

    for row in reader:

        if len(row) < 2:
            continue

        if normalize(row[0]) == normalize(circle_id):
            return normalize(row[1])

    return None


@app.post("/process")
async def process_excel(
    files: list[UploadFile] = File(...),
    circle_id: str = Form(...),
    visit: str = Form(...)
):

    target_value = get_target_value(circle_id)

    if not target_value:
        return {
            "error": f"{circle_id} に対応する値が見つかりません"
        }

    temp_dir = tempfile.mkdtemp()

    zip_buffer = BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:

        for file in files:

            input_path = os.path.join(
                temp_dir,
                file.filename
            )

            with open(input_path, "wb") as f:
                f.write(await file.read())

            wb = openpyxl.load_workbook(
                input_path
            )

            for sheet_name in wb.sheetnames:

                ws = wb[sheet_name]

                header_row = None
                target_col_index = None

                # header探索
                for row in ws.iter_rows(
                    min_row=1,
                    max_row=5,
                    values_only=True
                ):

                    headers = [
                        str(v).strip() if v else ""
                        for v in row
                    ]

                    for col_name in columns_to_remove:

                        if col_name in headers:

                            header_row = headers
                            target_col_index = headers.index(col_name)
                            break

                    if target_col_index is not None:
                        break

                if target_col_index is None:
                    continue

                rows_to_keep = []

                # header
                rows_to_keep.append(header_row)

                for row in ws.iter_rows(
                    min_row=2,
                    values_only=True
                ):

                    if len(row) <= target_col_index:
                        continue

                    val = str(
                        row[target_col_index]
                    ).strip()

                    if val == target_value:
                        rows_to_keep.append(row)

                # 全削除
                ws.delete_rows(
                    1,
                    ws.max_row
                )

                # 再書込
                for row in rows_to_keep:
                    ws.append(list(row))

            output_excel = os.path.join(
                temp_dir,
                file.filename
            )

            wb.save(output_excel)

            zipf.write(
                output_excel,
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