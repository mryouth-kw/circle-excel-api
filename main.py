from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware

import pandas as pd
import openpyxl
import zipfile
import tempfile
import os

app = FastAPI()

# Studioからアクセス可能にする
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


@app.post("/process")
async def process_excel(
    files: list[UploadFile] = File(...),
    target_value: str = Form(...)
):

    temp_dir = tempfile.mkdtemp()

    zip_path = os.path.join(temp_dir, "result.zip")

    with zipfile.ZipFile(zip_path, "w") as zipf:

        for file in files:

            input_path = os.path.join(temp_dir, file.filename)

            with open(input_path, "wb") as f:
                f.write(await file.read())

            wb = openpyxl.load_workbook(input_path)

            for sheet_name in wb.sheetnames:

                ws = wb[sheet_name]

                rows = list(ws.values)

                if not rows:
                    continue

                headers = [str(h) if h else "" for h in rows[0]]

                target_col_index = None

                for col_name in columns_to_remove:
                    if col_name in headers:
                        target_col_index = headers.index(col_name)
                        break

                if target_col_index is None:
                    continue

                filtered_rows = [headers]

                for row in rows[1:]:

                    if len(row) <= target_col_index:
                        continue

                    cell_value = str(row[target_col_index]).strip()

                    if cell_value == target_value:
                        filtered_rows.append(row)

                new_wb = openpyxl.Workbook()
                new_ws = new_wb.active
                new_ws.title = sheet_name

                for row in filtered_rows:
                    new_ws.append(list(row))

                wb.remove(wb[sheet_name])
                wb._add_sheet(new_ws)

            output_excel = os.path.join(
                temp_dir,
                file.filename
            )

            wb.save(output_excel)

            zipf.write(
                output_excel,
                arcname=file.filename
            )

    return FileResponse(
        zip_path,
        media_type="application/zip",
        filename="result.zip"
    )