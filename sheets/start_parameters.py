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
        ["Назва аркуша зі статистикою", params["sheet_stat"]],
        ["Назва аркуша з факторами", params["sheet_factor"]],
        ["Рік прогнозування прогнозу", params["model_year"]],
        ["", ""],
        ["column_year (рік)", params["column_year"]],
        ["column_month (місяць)", params["column_month"]],
        ["range_data (діапазон)", params["range_data"]],
        ["row_title (заголовки)", params["row_title"]],
        ["row_first_data", params["row_first_data"]],
        ["row_last_data", params["row_last_data"]],
        ["k (згладжування)", params["k"]],
        ["", ""],
        ["Налаштування факторів впливу", ""],
        ["", ""],
        ["Колонка року факторів", params["factor_column_year"]],
        ["Колонка місяця факторів", params["factor_column_month"]],
        ["Діапазон даних факторів", params["factor_row_range_data"]],
        ["Рядок опису факторів", params["factor_row_description"]],
        ["Рядок типу факторів", params["factor_row_type"]],
        ["Рядок заголовків факторів", params["factor_row_title"]],
        ["Перший рядок даних факторів", params["factor_row_first_data"]],
        ["Останній рядок даних факторів", params["factor_row_last_data"]],
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