# sheets/smoothed_data.py
from openpyxl.utils import column_index_from_string
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# Українські назви місяців
MONTH_NAMES = [
    "", "січень", "лютий", "березень", "квітень", "травень", "червень",
    "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"
]

def _moving_average(data, k):
    """Ковзне середнє: перше значення не змінюється"""
    if len(data) == 0:
        return []
    if k <= 1:
        return data[:]

    result = []
    for i in range(len(data)):
        start = max(0, i - k + 1)
        window = data[start:i+1]
        avg = sum(window) / len(window)
        result.append(avg)
    return result


def create_sheet_smoothed_data(workbook, params):
    sheet_name = "Згладжені дані"
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])

    ws_input = params["active_sheet"]
    ws = workbook.create_sheet(title=sheet_name)

    # === Параметри ===
    col_year = column_index_from_string(params["column_year"])
    col_month = column_index_from_string(params["column_month"])
    col_start = params["range_start_col"]
    col_end = params["range_end_col"]
    row_title = params["row_title"]
    row_first = params["row_first_data"]
    row_last = params["row_last_data"]
    k = params["k"]

    # === Читання заголовків ===
    input_headers = []
    for c in range(col_start, col_end + 1):
        cell = ws_input.cell(row_title, c)
        input_headers.append(cell.value or f"Колонка {get_column_letter(c)}")

    # === Читання даних ===
    years, months = [], []
    data_matrix = {c: [] for c in range(col_start, col_end + 1)}

    for r in range(row_first, row_last + 1):
        year = ws_input.cell(r, col_year).value
        month = ws_input.cell(r, col_month).value
        if year is None or month is None:
            continue
        years.append(year)
        months.append(int(month))
        for c in range(col_start, col_end + 1):
            val = ws_input.cell(r, c).value
            data_matrix[c].append(float(val) if val not in (None, '', '-') else 0.0)

    # === Згладжування ===
    smoothed = {}
    for c in data_matrix:
        smoothed[c] = _moving_average(data_matrix[c], k)

    # === Формування рядків ===
    total_months = len(years)
    rows = []

    # Рядок 1: Підпис
    input_start_col = 5  # після 4 мета-колонок
    input_end_col = input_start_col + len(input_headers) - 1
    smoothed_start_col = input_end_col + 3  # +2 порожні

    row1 = [""] * (smoothed_start_col + len(input_headers) - 1)
    # ВХІДНІ ДАНІ
    ws.cell(1, 1).value = "ВХІДНІ ДАНІ"
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1, end_column=input_end_col)
    # ЗГЛАДЖЕНІ ДАНІ
    ws.cell(1, smoothed_start_col).value = "ЗГЛАДЖЕНІ ДАНІ"
    ws.merge_cells(start_row=1, start_column=smoothed_start_col,
                   end_row=1, end_column=smoothed_start_col + 4+ len(input_headers) - 1)

    # Рядок 2: порожній
    ws.append([])

    # Рядок 3: Заголовки
    header_row = [
        "Рік", "Місяць", "Назва місяця", "Номер місяця"
    ] + input_headers + ["", ""] + [
        "Рік", "Місяць", "Назва місяця", "Номер місяця"
    ] + input_headers

    ws.append(header_row)

    # === Дані (починаючи з рядка 4) ===
    for i in range(total_months):
        month_name = MONTH_NAMES[months[i]]
        month_num = i + 1

        input_vals = []
        for c in range(col_start, col_end + 1):
            orig_val = data_matrix[c][i]
            input_vals.append(round(orig_val, 2))

        smoothed_vals = []
        for c in range(col_start, col_end + 1):
            sm_val = smoothed[c][i]
            smoothed_vals.append(round(sm_val, 2) if sm_val is not None else None)

        row = [
            years[i], months[i], month_name, month_num
        ] + input_vals + ["", ""] + [
            years[i], months[i], month_name, month_num
        ] + smoothed_vals

        ws.append(row)

        # === СТИЛІЗАЦІЯ ===
        from openpyxl.styles import Font, Alignment, PatternFill

        DARK_ORANGE = "FF8C00"
        LIGHT_GRAY = "D3D3D3"

        fill_orange = PatternFill(start_color=DARK_ORANGE, end_color=DARK_ORANGE, fill_type="solid")
        fill_gray = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")

        bold = Font(bold=True)
        bold_large = Font(bold=True, size=14)
        center = Alignment(horizontal="center", vertical="center")

        # 1. Підписи
        for cell in ws[1]:
            if cell.value in ("ВХІДНІ ДАНІ", "ЗГЛАДЖЕНІ ДАНІ"):
                cell.font = bold_large
                cell.alignment = center

        # 2. Мета-заголовки
        meta_positions = [1, 2, 3, 4]  # ліва частина
        meta_positions += [input_end_col + 3, input_end_col + 4, input_end_col + 5, input_end_col + 6]  # права
        for col_idx in meta_positions:
            cell = ws.cell(3, col_idx)
            cell.fill = fill_orange
            cell.font = bold
            cell.alignment = center

        # 3. Дані-заголовки
        data_cols = list(range(5, input_end_col + 1)) + \
                    list(range(smoothed_start_col + 4, smoothed_start_col + 4 + len(input_headers)))
        for col_idx in data_cols:
            cell = ws.cell(3, col_idx)
            cell.fill = fill_gray
            cell.font = bold
            cell.alignment = center

        # Числа — вирівнювання
        for row in ws.iter_rows(min_row=4):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal="right")

        # Автоширина
        for col_cells in ws.columns:
            column_letter = None
            max_length = 0
            for cell in col_cells:
                if cell.coordinate in ws.merged_cells:
                    continue
                if column_letter is None:
                    column_letter = cell.column_letter
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            if column_letter:
                ws.column_dimensions[column_letter].width = min(max_length + 2, 50)