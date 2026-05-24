from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse

from openpyxl import load_workbook
from fastapi.middleware.cors import CORSMiddleware

import tempfile
import zipfile
import os

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/process")
async def process_excel(
    files: list[UploadFile] = File(...)
):

    # 一時フォルダ
    temp_dir = tempfile.mkdtemp()

    processed_files = []

    for file in files:

        # 保存先
        input_path = os.path.join(
            temp_dir,
            file.filename
        )

        # upload保存
        with open(input_path, "wb") as f:
            f.write(await file.read())

        # Excel読込
        wb = load_workbook(input_path)

        # シート名確認
        print(wb.sheetnames)

        # 出力ファイル
        output_path = os.path.join(
            temp_dir,
            f"processed_{file.filename}"
        )

        # 保存
        wb.save(output_path)

        processed_files.append(output_path)

    # ZIP生成
    zip_path = os.path.join(
        temp_dir,
        "result.zip"
    )

    with zipfile.ZipFile(zip_path, "w") as zipf:
        for path in processed_files:
            zipf.write(
                path,
                os.path.basename(path)
            )

    return FileResponse(
        zip_path,
        filename="result.zip",
        media_type="application/zip"
    )