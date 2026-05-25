from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

import openpyxl
import zipfile

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


@app.get("/")
def root():
    return {"message": "API OK"}


@app.post("/process")
async def process_excel(
    files: list[UploadFile] = File(...),
    target_value: str = Form(...),
    circle_id: str = Form(...),
    visit: str = Form(...)
):


print("===== START PROCESS =====")
    print("target_value:", target_value)

    zip_buffer = BytesIO()

    with zipfile.ZipFile(
        zip_buffer,
        "w",
        zipfile.ZIP_DEFLATED
    ) as zipf:

        for file in files:

            print(f"Processing: {file.filename}")

            # UploadFile -> memory
            file_bytes = await file.read()

            excel_buffer = BytesIO(file_bytes)

            wb = openpyxl.load_workbook(excel_buffer)

            for sheet_name in wb.sheetnames:

                ws = wb[sheet_name]

                rows = list(ws.values)

                if not rows:
                    continue

                headers = [
                    str(h).strip() if h else ""
                    for h in rows[0]
                ]

                target_col_index = None

                for col_name in columns_to_remove:
                    if col_name in headers:
                        target_col_index = headers.index(col_name)
                        break

                # ID列が無いシートはスキップ
                if target_col_index is None:
                    continue

                filtered_rows = [headers]

                for row in rows[1:]:

                    if len(row) <= target_col_index:
                        continue

                    cell_value = str(
                        row[target_col_index]
                    ).strip()

                    if cell_value == target_value:
                        filtered_rows.append(row)

                # 元シート全削除
                ws.delete_rows(1, ws.max_row)

                # フィルタ後を書き戻し
                for row in filtered_rows:
                    ws.append(list(row))

                print(
                    f"{sheet_name}: {len(filtered_rows)-1} rows"
                )

            # 出力をメモリ化
            output_buffer = BytesIO()

            wb.save(output_buffer)

            output_buffer.seek(0)

            zipf.writestr(
                file.filename,
                output_buffer.read()
            )

            print(f"DONE: {file.filename}")

    zip_buffer.seek(0)

    zip_filename = (
        f"調査票2_{circle_id}_Visit-{visit}.zip"
    )

    print("===== FINISH =====")

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition":
            f'attachment; filename="{zip_filename}"'
        }
    )