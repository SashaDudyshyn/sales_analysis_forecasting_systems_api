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
    headers = params["input_headers"]
    trend_forecasts = params["trend_forecasts"]
    seasonal_coeffs = params["seasonal_coeffs"]
    factors_data = params.get("factors_data", [])

    # Нормалізація заголовків
    header_normalized = {h.strip().lower().replace(" ", ""): h for h in headers}
    factors_by_header = {}

    for f in factors_data:
        key = f["header"].strip().lower().replace(" ", "")
        if key in header_normalized:
            original = header_normalized[key]
            factors_by_header.setdefault(original, []).append({
                "desc": f["description"],
                "type": f["type"],
                "values": f["data"]
            })

    # === Формування заголовків і підрахунок колонок ===
    header_row = ["Рік", "Місяць", "Назва місяця", "Номер місяця", ""]
    for header in headers:
        header_row += ["Тренд", "З урахуванням сезонності"]
        for f in factors_by_header.get(header, []):
            header_row += [f"{f['desc']} ({f['type']})"]
        header_row += ["Фінальний прогноз"]
        if header != headers[-1]:
            header_row += [""]

    total_cols = len(header_row)

    # === Головний заголовок  ===
    ws.cell(1, 1, f"Фінальний прогноз на {model_year} рік")
    if total_cols > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.append([])  # рядок 2 — порожній

    # === Заголовки колонок (рядок 3) ===
    ws.append(header_row)

    # === Заповнення даних (12 місяців) ===
    for month_num in range(1, 13):
        row = [model_year, month_num, MONTH_NAMES[month_num], month_num, ""]

        for idx, header in enumerate(headers):
            col_idx = params["range_start_col"] + idx
            trend_val = trend_forecasts.get(col_idx, [0]*12)[month_num-1] or 0.0
            coeff = seasonal_coeffs.get((month_num, col_idx), 1.0)
            seasonal_val = round(trend_val * coeff, 2)

            row += [trend_val, seasonal_val]

            final_val = seasonal_val
            for f in factors_by_header.get(header, []):
                val = f["values"][month_num-1]
                if val is not None:
                    if f["type"] == "коефіцієнт":
                        final_val = round(final_val * val, 2)
                    else:
                        final_val = round(final_val + val, 2)
                row += [val if val is not None else ""]

            row += [final_val]
            if idx < len(headers) - 1:
                row += [""]

        ws.append(row)

    # === СТИЛІЗАЦІЯ ===
    bold = Font(bold=True)
    center = Alignment(horizontal="center", vertical="center")
    wrap_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    orange = PatternFill("solid", fgColor="FF8C00")
    gray   = PatternFill("solid", fgColor="D3D3D3")
    light  = PatternFill("solid", fgColor="F0F0F0")
    blue   = PatternFill("solid", fgColor="DDEBF7")

    # Заголовки (рядок 3)
    for cell in ws[3]:
        if cell.value:
            cell.font = bold
            cell.alignment = wrap_center
            if cell.column <= 5:
                cell.fill = orange
            elif "Тренд" in str(cell.value):
                cell.fill = gray
            elif "сезонност" in str(cell.value):
                cell.fill = light
            elif "Фінальний" in str(cell.value):
                cell.fill = blue

    ws.row_dimensions[3].height = 70

    # Дані — центрування + формат чисел
    for row in ws.iter_rows(min_row=4):
        for cell in row:
            if cell.value is not None:
                cell.alignment = center
            if isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00'

    # Автоширина (тільки по даним)
    for col_idx, col in enumerate(ws.columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = 8
        for cell in col:
            if cell.row >= 4 and cell.value is not None and cell.coordinate not in ws.merged_cells:
                max_len = max(max_len, len(str(cell.value)))
        # Додатково для колонок факторів
        header_text = ws.cell(3, col_idx).value or ""
        if any(kw in str(header_text).lower() for kw in ["фактор", "коефіцієнт", "одиниці", "фінальний"]):
            max_len = max(max_len, 18)
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)

    # Рамки
    thin = Side(border_style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=total_cols):
        for cell in row:
            cell.border = border

    # === ЗБИРАЄМО ФІНАЛЬНИЙ ПРОГНОЗ ДЛЯ КОЖНОГО НАБОРУ ДАНИХ ===
    final_forecast_by_col = {}  # ← потрібно для візуалізації

    # Пробігаємо ще раз по місяцях, щоб зібрати фінальні значення в словник
    for month_num in range(1, 13):
        row_values_for_month = []  # тимчасовий список для поточного місяця

        for idx, header in enumerate(headers):
            col_idx = params["range_start_col"] + idx

            trend_val = trend_forecasts.get(col_idx, [0]*12)[month_num-1] or 0.0
            coeff = seasonal_coeffs.get((month_num, col_idx), 1.0)
            seasonal_val = round(trend_val * coeff, 2)

            final_val = seasonal_val
            for f in factors_by_header.get(header, []):
                val = f["values"][month_num-1]
                if val is not None:
                    if f["type"] == "коефіцієнт":
                        final_val = round(final_val * val, 2)
                    else:
                        final_val = round(final_val + val, 2)

            row_values_for_month.append(final_val)

        # Записуємо в словник по місяцях
        for idx, val in enumerate(row_values_for_month):
            col_idx = params["range_start_col"] + idx
            if col_idx not in final_forecast_by_col:
                final_forecast_by_col[col_idx] = []
            final_forecast_by_col[col_idx].append(val)

    # === ПОВЕРТАЄМО РЕЗУЛЬТАТ ===
    return {
        "final_forecast_by_col": final_forecast_by_col   # {col_idx: [12 значень]}
    }