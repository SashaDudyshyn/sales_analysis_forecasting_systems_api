# sheets/start_parameters.py
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def create_sheet_start_parameters(workbook, params):
    sheet_name = "Початкові налаштування"
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])

    ws = workbook.create_sheet(title=sheet_name)

    # Формуємо список регіонів/товарів
    headers = params.get("input_headers", [])
    headers_str = ", ".join(headers) if headers else "—"

    # Дані
    data = [
        ["Параметр", "Значення"],
        ["Файл", params["filename"]],
        ["Аркуш зі статистикою", params["sheet_stat"]],
        ["Аркуш з факторами впливу", params["sheet_factor"]],
        ["Рік прогнозу", params["model_year"]],
        ["", ""],
        ["Налаштування статистичних даних", ""],
        ["Колонка року", params["column_year"]],
        ["Колонка місяця", params["column_month"]],
        ["Діапазон даних", params["range_data"]],
        ["Рядок заголовків", params["row_title"]],
        ["Перший рядок даних", params["row_first_data"]],
        ["Останній рядок даних", params["row_last_data"]],
        ["Коефіцієнт згладжування (k)", params["k"]],
        ["Набори даних", headers_str],
        ["", ""],
        ["Налаштування факторів впливу", ""],
        ["Колонка року (фактори)", params["factor_column_year"]],
        ["Колонка місяця (фактори)", params["factor_column_month"]],
        ["Діапазон значень факторів", params["factor_row_range_data"]],
        ["Рядок опису", params["factor_row_description"]],
        ["Рядок типу", params["factor_row_type"]],
        ["Рядок назв факторів", params["factor_row_title"]],
        ["Перший рядок даних", params["factor_row_first_data"]],
        ["Останній рядок даних", params["factor_row_last_data"]],
    ]

    for row in data:
        ws.append(row)

    # === СТИЛІ ===
    bold = Font(bold=True, size=12)
    bold_white = Font(bold=True, color="FFFFFF", size=13)
    section_font = Font(bold=True, color="1F4E79", size=12)

    header_fill = PatternFill("solid", fgColor="1F4E79")        # темно-синій
    section_fill = PatternFill("solid", fgColor="DDEBF7")      # світло-блакитний
    border = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

    # Заголовок таблиці
    ws["A1"].value = "Параметр"
    ws["B1"].value = "Значення"
    ws["A1"].font = bold_white
    ws["B1"].font = bold_white
    ws["A1"].fill = header_fill
    ws["B1"].fill = header_fill
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B1"].alignment = Alignment(horizontal="center", vertical="center")

    # Групові заголовки
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
        cell = row[0]
        if cell.value and "Налаштування" in str(cell.value):
            cell.font = section_font
            cell.fill = section_fill
            ws.merge_cells(start_row=cell.row, start_column=1, end_row=cell.row, end_column=2)
            cell.alignment = Alignment(horizontal="left", vertical="center")

    # Основні рядки
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        a_cell, b_cell = row[0], row[1]

        # Вирівнювання
        a_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        b_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Шрифти
        a_cell.font = Font(size=11)
        b_cell.font = Font(size=11, bold=True, color="1F4E79")

        # Рамки
        a_cell.border = border
        b_cell.border = border

    # Висота рядків
    for i in range(1, ws.max_row + 1):
        if i == 1:
            ws.row_dimensions[i].height = 30
        elif "Налаштування" in str(ws.cell(i, 1).value or ""):
            ws.row_dimensions[i].height = 24
        else:
            ws.row_dimensions[i].height = 21

    # Особлива висота для довгого списку регіонів
    for row in ws.iter_rows():
        if row[0].value == "Набори даних":
            ws.row_dimensions[row[0].row].height = 36 if len(headers_str) > 80 else 24
            break

    # === АВТОШИРИНА КОЛОНОК (по реальному вмісту, включно із заголовками) ===
    for col_letter in ["A", "B"]:
        max_length = 0
        column = ws[col_letter]
        for cell in column:
            if cell.value:
                # Для колонки B враховуємо можливий перенос
                length = len(str(cell.value))
                if col_letter == "B" and "\n" not in str(cell.value):
                    length = max(length, 20)  # мінімальна ширина для значень
                max_length = max(max_length, length)
        adjusted_width = min(max_length + 3, 70)
        ws.column_dimensions[col_letter].width = adjusted_width

    # Заморозка заголовка
    ws.freeze_panes = "A2"

    #  відступ зверху
    ws.sheet_view.showGridLines = False