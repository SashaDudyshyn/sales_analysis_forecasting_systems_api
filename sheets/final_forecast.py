# sheets/final_forecast.py
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

MONTH_NAMES = ["", "січень", "лютий", "березень", "квітень", "травень", "червень",
               "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"]


def create_sheet_final_forecast(workbook, params):
    sheet_name = "Фінальний прогноз"
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])
    ws = workbook.create_sheet(title=sheet_name)

    model_year = params["model_year"]
    headers = params["input_headers"]                      # ["Схід", "Захід", ...]
    trend_forecasts = params["trend_forecasts"]            # {col_idx: [12 значень]}
    seasonal_coeffs = params["seasonal_coeffs"]            # {(month, col): coeff}
    factors_data = params.get("factors_data", [])           # список з факторами

    # Нормалізуємо назви заголовків для порівняння (без регістру і пробілів)
    header_normalized = {h.strip().lower().replace(" ", ""): h for h in headers}

    # Групуємо фактори по набору даних (header)
    factors_by_header = {}  # "схід" → список факторів, які до нього відносяться
    for f in factors_data:
        header_key = f["header"].strip().lower().replace(" ", "")
        if header_key in header_normalized:
            original_header = header_normalized[header_key]
            if original_header not in factors_by_header:
                factors_by_header[original_header] = []
            factors_by_header[original_header].append({
                "desc": f["description"],
                "type": f["type"],
                "values": f["data"]  # список 12 значень
            })

    # Головний заголовок
    # Рахуємо кількість колонок динамічно
    total_cols = 5
    for header in headers:
        factor_count = len(factors_by_header.get(header, []))
        total_cols += 2 + factor_count + 1  # Тренд + Сезонний + фактори + Фінальний

    ws.cell(1, 1, f"Фінальний прогноз на {model_year} рік")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.append([])  # рядок 2

    # Заголовки колонок
    header_row = ["Рік", "Місяць", "Назва місяця", "Номер місяця", ""]
    for header in headers:
        header_row += ["Тренд", "З урахуванням сезонності"]

        # Додаємо тільки ті фактори, що стосуються цього набору
        for f in factors_by_header.get(header, []):
            header_row += [f"{f['desc']} ({f['type']})"]

        header_row += ["Фінальний прогноз"]

        if header != headers[-1]:
            header_row += [""]  # порожній між групами

    ws.append(header_row)

    # Стилі заголовків
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    orange = PatternFill("solid", fgColor="FF8C00")
    gray = PatternFill("solid", fgColor="D3D3D3")
    light = PatternFill("solid", fgColor="F0F0F0")
    blue = PatternFill("solid", fgColor="DDEBF7")

    for cell in ws[3]:
        cell.font = bold
        cell.alignment = center
        col = cell.column
        if col <= 5:
            cell.fill = orange
        elif "Тренд" in str(cell.value):
            cell.fill = gray
        elif "сезонност" in str(cell.value):
            cell.fill = light
        elif "Фінальний" in str(cell.value):
            cell.fill = blue

    # Дані — 12 місяців
    for month_num in range(1, 13):
        row = [model_year, month_num, MONTH_NAMES[month_num], month_num, ""]

        for idx, header in enumerate(headers):
            col_idx = params["range_start_col"] + idx

            # 1. Тренд
            trend_val = trend_forecasts.get(col_idx, [0]*12)[month_num-1] or 0.0

            # 2. З урахуванням сезонності
            coeff = seasonal_coeffs.get((month_num, col_idx), 1.0)
            seasonal_val = round(trend_val * coeff, 2)

            row += [trend_val, seasonal_val]

            # 3. Фактори — тільки ті, що є для цього header
            final_val = seasonal_val
            factors_for_header = factors_by_header.get(header, [])

            for f in factors_for_header:
                val = f["values"][month_num-1]
                if val is not None:
                    if f["type"] == "коефіцієнт":
                        final_val = round(final_val * val, 2)
                    else:  # "одиниці"
                        final_val = round(final_val + val, 2)
                row += [val if val is not None else ""]

            # 4. Фінальний прогноз
            row += [final_val]

            if idx < len(headers) - 1:
                row += [""]

        ws.append(row)

        # -- Форматування
        # ===  РОЗРАХУНОК КІЛЬКОСТІ КОЛОНОК ДЛЯ ОБ’ЄДНАННЯ ЗАГОЛОВКА ===
        total_cols = len(header_row)


        # Головний заголовок
        ws.cell(1, 1, f"Фінальний прогноз на {model_year} рік")
        if total_cols > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        # === ПЕРЕНЕСЕННЯ + ЦЕНТРУВАННЯ УСІХ ЗАГОЛОВКІВ (рядок 3) ===
        for cell in ws[3]:
            if cell.value:
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True
                )
        ws.row_dimensions[3].height = 70

        # === ЦЕНТРУВАННЯ ВСІХ ДАНИХ (починаючи з рядка 4) ===
        for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.value is not None:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'

        # === АВТОШИРИНА — тільки по даним, без урахування заголовків ===
        for col_idx, column_cells in enumerate(ws.columns, start=1):
            col_letter = get_column_letter(col_idx)
            max_len = 8

            for cell in column_cells:
                if cell.row >= 4 and cell.value is not None:
                    if cell.coordinate in ws.merged_cells:
                        continue
                    max_len = max(max_len, len(str(cell.value)))

            # Додатково для колонок факторів і фінального прогнозу
            header_text = ws.cell(3, col_idx).value or ""
            if any(kw in str(header_text).lower() for kw in ["фактор", "коефіцієнт", "одиниці", "фінальний"]):
                max_len = max(max_len, 18)

            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

        # === РАМКИ ===
        thin = Side(border_style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=total_cols):
            for cell in row:
                cell.border = border