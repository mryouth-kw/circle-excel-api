from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

import openpyxl
import zipfile
import tempfile
import os
import csv
import requests
import shutil
import gc

from urllib.parse import quote

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
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
# cache
# =========================
mapping_cache = None


# =========================
# normalize
# =========================
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


# =========================
# safe filename
# =========================
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


# =========================
# load mapping
# =========================
def load_mapping():

    global mapping_cache

    if mapping_cache is not None:
        return mapping_cache

    print("LOAD GOOGLE SHEET")

    try:

        res = requests.get(
            CSV_URL,
            timeout=15
        )

        res.raise_for_status()

    except Exception as e:

        print("GOOGLE SHEET ERROR:", str(e))

        raise HTTPException(
            status_code=500,
            detail=f"Google Sheets load failed: {str(e)}"
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


# =========================
# target sheet
# =========================
def get_target_sheet(wb):

    sheetnames = wb.sheetnames

    if len(sheetnames) == 0:
        return None

    first_sheet = sheetnames[0]

    print("FIRST SHEET:", first_sheet)

    if normalize(first_sheet) == "検索情報":

        print("検索情報 sheet detected")

        if len(sheetnames) >= 2:

            print(
                "USE SECOND SHEET:",
                sheetnames[1]
            )

            return wb[sheetnames[1]]

        return None

    print("USE FIRST SHEET")

    return wb[first_sheet]


@app.get("/")
@app.head("/")
def root():
    return {"message": "API OK"}


@app.post("/process")
async def process_excel(
    files: list[UploadFile] = File(...),
    circle_id: str = Form(...),
    visit: str = Form(...)
):

    print("=" * 80)
    print("PROCESS START")
    print("=" * 80)

    target_value = get_target_value(circle_id)

    print("CIRCLE ID:", circle_id)
    print("TARGET VALUE:", target_value)

    if not target_value:

        raise HTTPException(
            status_code=404,
            detail=f"{circle_id} に対応する値が見つかりません"
        )

    temp_dir = tempfile.mkdtemp()

    zip_path = os.path.join(
        temp_dir,
        f"調査票2_{circle_id}_Visit-{visit}.zip"
    )

    try:

        with zipfile.ZipFile(
            zip_path,
            mode="w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=0
        ) as zipf:

            for file in files:

                wb = None

                try:

                    print("=" * 80)
                    print("FILE:", file.filename)
                    print("=" * 80)

                    safe_input_name = safe_filename(
                        file.filename
                    )

                    input_path = os.path.join(
                        temp_dir,
                        safe_input_name
                    )

                    # =========================
                    # upload save
                    # =========================
                    with open(input_path, "wb") as f:

                        while True:

                            chunk = await file.read(1024 * 1024)

                            if not chunk:
                                break

                            f.write(chunk)

                    await file.close()

                    base_name, ext = os.path.splitext(
                        file.filename
                    )

                    ext = ext.lower()

                    # =========================
                    # xlsx only
                    # =========================
                    if ext != ".xlsx":

                        raise HTTPException(
                            status_code=400,
                            detail=f"{file.filename}: .xlsx のみ対応しています"
                        )

                    safe_target_value = safe_filename(
                        target_value
                    )

                    print("XLSX MODE")

                    wb = openpyxl.load_workbook(
                        input_path,
                        data_only=False,
                        keep_links=False
                    )

                    # =========================
                    # recalc OFF
                    # =========================
                    try:
                        wb.calculation.fullCalcOnLoad = False
                        wb.calculation.forceFullCalc = False
                    except:
                        pass

                    output_filename = (
                        f"{safe_target_value}_{safe_input_name}"
                    )

                    output_path = os.path.join(
                        temp_dir,
                        output_filename
                    )

                    # =========================
                    # target sheet
                    # =========================
                    ws = get_target_sheet(wb)

                    if ws is None:

                        wb.save(output_path)

                        wb.close()

                        zipf.write(
                            output_path,
                            arcname=output_filename
                        )

                        continue

                    target_col_index = None
                    header_row_index = None

                    # =========================
                    # header search
                    # =========================
                    for row_idx, row in enumerate(
                        ws.iter_rows(
                            min_row=1,
                            max_row=min(10, ws.max_row),
                            values_only=True
                        ),
                        start=1
                    ):

                        headers = [
                            normalize(v)
                            for v in row
                        ]

                        for col_name in columns_to_search:

                            if col_name in headers:

                                target_col_index = (
                                    headers.index(col_name) + 1
                                )

                                header_row_index = row_idx

                                break

                        if target_col_index:
                            break

                    if header_row_index is None:
                        header_row_index = 1

                    # =========================
                    # column not found
                    # =========================
                    if not target_col_index:

                        print("COLUMN NOT FOUND")

                        wb.save(output_path)

                        wb.close()

                        zipf.write(
                            output_path,
                            arcname=output_filename
                        )

                        continue

                    matched_count = 0

                    delete_rows = []

                    # =========================
                    # scan rows
                    # =========================
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

                        if cell_value == target_value:

                            matched_count += 1

                        else:

                            delete_rows.append(row_idx)

                    print(
                        "MATCHED COUNT:",
                        matched_count
                    )

                    # =========================
                    # delete reverse order
                    # =========================
                    if delete_rows:

                        ranges = []

                        start = delete_rows[0]
                        prev = delete_rows[0]

                        for row_idx in delete_rows[1:]:

                            if row_idx == prev + 1:

                                prev = row_idx

                            else:

                                ranges.append(
                                    (start, prev)
                                )

                                start = row_idx
                                prev = row_idx

                        ranges.append((start, prev))

                        # 後ろから削除
                        for start, end in reversed(ranges):

                            ws.delete_rows(
                                start,
                                end - start + 1
                            )

                    print(
                        "SAVE:",
                        output_filename
                    )

                    # =========================
                    # dimension/recalc suppression
                    # =========================
                    try:
                        ws.sheet_properties.outlinePr.summaryBelow = True
                    except:
                        pass

                    try:
                        ws._current_row = ws.max_row
                    except:
                        pass

                    try:
                        wb.calculation.fullCalcOnLoad = False
                        wb.calculation.forceFullCalc = False
                    except:
                        pass

                    wb.save(output_path)

                    wb.close()

                    del wb
                    del delete_rows

                    gc.collect()

                    zipf.write(
                        output_path,
                        arcname=output_filename
                    )

                    # =========================
                    # cleanup per file
                    # =========================
                    try:
                        os.remove(input_path)
                    except:
                        pass

                    try:
                        os.remove(output_path)
                    except:
                        pass

                    print(
                        "ZIP ADD:",
                        output_filename
                    )

                except Exception as e:

                    if wb:
                        try:
                            wb.close()
                        except:
                            pass

                    raise HTTPException(
                        status_code=500,
                        detail=f"{file.filename}: {str(e)}"
                    )

                finally:

                    gc.collect()

        # =========================
        # zip response
        # =========================
        zip_filename = os.path.basename(zip_path)

        encoded_filename = quote(zip_filename)

        return StreamingResponse(
            open(zip_path, "rb"),
            media_type="application/zip",
            headers={
                "Content-Disposition":
                f"attachment; filename*=UTF-8''{encoded_filename}"
            }
        )

    finally:

        gc.collect()

        try:
            shutil.rmtree(temp_dir)
        except:
            pass