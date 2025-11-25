# main.py
from fastapi import FastAPI, File, UploadFile, Form
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook
from io import BytesIO
import logging

from models.excel_params import ExcelProcessParams
from sheets.start_parameters import create_sheet_start_parameters
from sheets.smoothed_data import create_sheet_smoothed_data
from openpyxl.utils import column_index_from_string

app = FastAPI(title="Excel API — Pydantic + Form", description="Окремі поля")

@app.post("/process-excel/")
async def process_excel(
    file: UploadFile = File(...),
    column_year: str = Form("B"),
    column_month: str = Form("D"),
    range_data: str = Form("G-J"),
    row_title: int = Form(3),
    row_first_data: int = Form(4),
    row_last_data: int = Form(38),
    k: int = Form(2),
):
    # --- Створюємо Pydantic-модель вручну ---
    try:
        params = ExcelProcessParams(
            column_year=column_year,
            column_month=column_month,
            range_data=range_data,
            row_title=row_title,
            row_first_data=row_first_data,
            row_last_data=row_last_data,
            k=k,
        )
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))

    # --- Обробка файлу ---
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Підтримується лише .xlsx")

    content = await file.read()
    workbook = load_workbook(filename=BytesIO(content))
    active_sheet = workbook.active

    range_start, range_end = params.range_data.split("-")

    params_dict = {
        "filename": file.filename,
        "active_sheet_name": active_sheet.title,
        "active_sheet": active_sheet,
        "column_year": params.column_year,
        "column_month": params.column_month,
        "range_data": params.range_data,
        "range_start_col": column_index_from_string(range_start),
        "range_end_col": column_index_from_string(range_end),
        "row_title": params.row_title,
        "row_first_data": params.row_first_data,
        "row_last_data": params.row_last_data,
        "k": params.k,
    }

    create_sheet_start_parameters(workbook, params_dict)
    create_sheet_smoothed_data(workbook, params_dict)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename=processed_{file.filename}"}
    )