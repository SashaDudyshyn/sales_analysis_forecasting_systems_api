# sheets/start_parameters.py
from openpyxl.styles import Font, Alignment


def create_sheet_start_parameters(workbook, params):
    sheet_name = "Початкові налаштування"
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])

    ws = workbook.create_sheet(title=sheet_name)

    # Дані
    data = [
        ["Параметр", "Значення"],
        ["Файл", params["filename"]],
        ["Активний аркуш", params["active_sheet_name"]],
        ["", ""],
        ["column_year (рік)", params["column_year"]],
        ["column_month (місяць)", params["column_month"]],
        ["range_data (діапазон)", params["range_data"]],
        ["row_title (заголовки)", params["row_title"]],
        ["row_first_data", params["row_first_data"]],
        ["row_last_data", params["row_last_data"]],
        ["k (згладжування)", params["k"]],
    ]

    for row in data:
        ws.append(row)

    # === СТИЛІЗАЦІЯ (БЕЗ StyleProxy) ===
    bold_font = Font(bold=True)
    left_align = Alignment(horizontal="left")

    # Заголовки — жирний
    for cell in ws[1]:
        cell.font = bold_font

    # Колонка A — вирівнювання ліворуч
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=1):
        for cell in row:
            cell.alignment = left_align