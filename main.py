from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware

import openpyxl
import zipfile
import tempfile
import os
import csv
import requests
import traceback

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

mapping_cache = None


def log(*args):
    print(*args, flush=True)


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

    try:

        num = float(text)

        if num.is_integer():
            return str(int(num))

    except:
        pass

    return text


def safe_filename(text):

    return (
        str(text)
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


def load_mapping():

    global mapping_cache

    if mapping_cache is not None:
        return mapping_cache

    log("LOAD GOOGLE SHEET")

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

    log(
        f"MAPPING COUNT: {len(mapping_cache)}"
    )

    return mapping_cache


def get_target_value(circle_id):

    mapping = load_mapping()

    return mapping.get(
        normalize(circle_id)
    )


def get_target_sheet(wb):

    sheetnames = wb.sheetnames

    if len(sheetnames) == 0:
        return None

    first_sheet = sheetnames[0]

    log("FIRST SHEET:", first_sheet)

    if normalize(first_sheet) == "検索情報":

        log("検索情報 sheet detected")

        if len(sheetnames) >= 2:

            log(
                "USE SECOND SHEET:",
                sheetnames[1]
            )

            return wb[sheetnames[1]]

        return None

    log("USE FIRST SHEET")

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

    log("")
    log("=" * 80)
    log("PROCESS START")
    log("=" * 80)

    log("FILES COUNT:", len(files))
    log("CIRCLE ID:", circle_id)
    log("VISIT:", visit)

    target_value = get_target_value(circle_id)

    log("TARGET VALUE:", target_value)

    if not target_value:

        log("TARGET VALUE NOT FOUND")

        return JSONResponse(
            status_code=400,
            content={
                "error":
                f"{circle_id} に対応する値が見つかりません"
            }
        )

    temp_dir = tempfile.mkdtemp()

    log("TEMP DIR:", temp_dir)

    zip_buffer = BytesIO()

    processed_file_count = 0

    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=1
    ) as zipf:

        for file in files:

            try:

                log("")
                log("=" * 80)
                log("FILE:", file.filename)
                log("=" * 80)

                input_path = os.path.join(
                    temp_dir,
                    file.filename
                )

                with open(input_path, "wb") as f:

                    content = await file.read()

                    log(
                        "FILE SIZE:",
                        len(content)
                    )

                    f.write(content)

                base_name, ext = os.path.splitext(
                    file.filename
                )

                ext = ext.lower()

                log("EXT:", ext)

                # xlsx / xlsm のみ許可
                if ext not in [".xlsx", ".xlsm"]:

                    log(
                        "UNSUPPORTED FILE:",
                        file.filename
                    )

                    continue

                safe_target_value = safe_filename(
                    target_value
                )

                output_filename = (
                    f"{safe_target_value}_{file.filename}"
                )

                output_path = os.path.join(
                    temp_dir,
                    output_filename
                )

                log("OPEN WORKBOOK")

                wb = openpyxl.load_workbook(
                    input_path,
                    keep_vba=True,
                    data_only=False
                )

                log(
                    "SHEETS:",
                    wb.sheetnames
                )

                ws = get_target_sheet(wb)

                if ws is None:

                    log(
                        "TARGET SHEET NOT FOUND"
                    )

                    wb.save(output_path)

                    zipf.write(
                        output_path,
                        arcname=output_filename
                    )

                    processed_file_count += 1

                    continue

                log(
                    "TARGET SHEET:",
                    ws.title
                )

                target_col_index = None
                header_row_index = None

                for row in ws.iter_rows(
                    min_row=1,
                    max_row=min(10, ws.max_row)
                ):

                    headers = [
                        normalize(cell.value)
                        for cell in row
                    ]

                    log(
                        "HEADER ROW:",
                        row[0].row,
                        headers
                    )

                    for col_name in columns_to_search:

                        if col_name in headers:

                            target_col_index = (
                                headers.index(col_name) + 1
                            )

                            header_row_index = row[0].row

                            log(
                                "FOUND COLUMN:",
                                col_name
                            )

                            break

                    if target_col_index:
                        break

                log(
                    "TARGET COLUMN INDEX:",
                    target_col_index
                )

                if not target_col_index:

                    log(
                        "COLUMN NOT FOUND"
                    )

                    wb.save(output_path)

                    zipf.write(
                        output_path,
                        arcname=output_filename
                    )

                    processed_file_count += 1

                    continue

                delete_rows = []

                matched_count = 0

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

                    if row_idx <= (
                        header_row_index + 20
                    ):

                        log(
                            "ROW:",
                            row_idx,
                            "RAW:",
                            raw_value,
                            "NORMALIZED:",
                            cell_value,
                            "TARGET:",
                            target_value
                        )

                    if cell_value == target_value:

                        matched_count += 1

                    else:

                        delete_rows.append(row_idx)

                log(
                    "MATCHED COUNT:",
                    matched_count
                )

                log(
                    "DELETE COUNT:",
                    len(delete_rows)
                )

                for row_idx in reversed(delete_rows):

                    ws.delete_rows(
                        row_idx,
                        1
                    )

                log(
                    "SAVE:",
                    output_filename
                )

                wb.save(output_path)

                zipf.write(
                    output_path,
                    arcname=output_filename
                )

                processed_file_count += 1

                log(
                    "ZIP ADD:",
                    output_filename
                )

            except Exception as e:

                log("")
                log("ERROR OCCURRED")
                log(str(e))

                traceback.print_exc()

    log("")
    log("=" * 80)
    log("PROCESSED FILE COUNT:", processed_file_count)
    log("=" * 80)

    if processed_file_count == 0:

        return JSONResponse(
            status_code=400,
            content={
                "error":
                "処理可能なxlsx/xlsmファイルがありませんでした"
            }
        )

    zip_buffer.seek(0)

    zip_filename = (
        f"調査票2_{circle_id}_Visit-{visit}.zip"
    )

    log(
        "RETURN ZIP:",
        zip_filename
    )

    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition":
            f'attachment; filename="{zip_filename}"'
        }
    )