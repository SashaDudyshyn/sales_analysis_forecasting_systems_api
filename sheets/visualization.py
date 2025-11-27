# sheets/visualization.py
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill


def create_sheet_visualization(workbook, params):
    sheet_name = "Візуалізація"
    if sheet_name in workbook.sheetnames:
        workbook.remove(workbook[sheet_name])
    ws = workbook.create_sheet(title=sheet_name)

    model_year = params["model_year"]
    first_header = params["input_headers"][0] if params["input_headers"] else "Дані"

    years = params["years"]
    months = params["months"]
    raw = params["raw_data"].get(params["range_start_col"], [])
    smooth = params["smoothed_data"].get(params["range_start_col"], [])
    deseas = params["deseasoned_data"].get(params["range_start_col"], [])
    final_fc = params["final_forecast"].get(params["range_start_col"], [])

    n_hist = max(len(years), len(months), len(raw), len(smooth), len(deseas), 1)

    def g(lst, i, default=None):
        return lst[i] if i < len(lst) else default

    # Заголовок
    ws.merge_cells("A1:E1")
    ws["A1"] = f"Динаміка та прогноз: {first_header}"
    ws["A1"].font = Font(bold=True, size=16, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 45

    # Заголовки таблиці
    headers = ["Період", "Сирі дані", "Згладжені", "Тренд", "Фінальний прогноз"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(3, c, h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Дані
    for i in range(n_hist):
        y = int(g(years, i) or 2000)
        m = int(g(months, i) or 1)
        ws.append([f"{y}-{m:02d}", g(raw, i), g(smooth, i), g(deseas, i), None])

    for m in range(1, 13):
        ws.append([f"{model_year}-{m:02d}", None, None, None, final_fc[m-1] if m-1 < len(final_fc) else None])

    # === АВТОШИРИНА: розумна, тільки по вмісту даних ===
    for col_idx in range(1, 6):
        column = get_column_letter(col_idx)
        max_length = 0

        for cell in ws[column]:
            if cell.row < 4:  # пропускаємо заголовок і рядок з назвами колонок
                continue
            if cell.value is not None:
                # Для чисел — рахуємо кількість символів у відформатованому вигляді
                if isinstance(cell.value, (int, float)):
                    val_str = f"{cell.value:,.0f}".replace(",", " ")  # 1234567 → "1 234 567"
                else:
                    val_str = str(cell.value)
                max_length = max(max_length, len(val_str))

        # Додаємо невеликий відступ
        adjusted_width = max_length + 2
        # Але не менше мінімальної ширини для зручності
        adjusted_width = max(adjusted_width, 10)
        # І не більше розумного максимуму
        adjusted_width = min(adjusted_width, 25)

        ws.column_dimensions[column].width = adjusted_width

    # Висота рядка з заголовками колонок — автоматична (бо wrap_text=True)
    ws.row_dimensions[3].height = None

    # Графік
    chart = LineChart()
    chart.title = f"Прогноз продажів: {first_header}"
    chart.style = 27
    chart.height = 20
    chart.width = 34
    chart.x_axis.title = "Період"
    chart.y_axis.title = "Обсяг"
    chart.legend.position = "b"

    data = Reference(ws, min_col=2, max_col=5, min_row=3, max_row=ws.max_row)
    cats = Reference(ws, min_col=1, min_row=4, max_row=ws.max_row)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    colors = ["1F4E79", "ED7D31", "A5A5A5", "70AD47"]
    for i, s in enumerate(chart.series):
        s.graphicalProperties.line.solidFill = colors[i]
        s.graphicalProperties.line.width = 28000
        if i == 3:
            s.graphicalProperties.line.dashStyle = "dash"
            s.graphicalProperties.line.width = 35000
        s.marker.symbol = "circle"
        s.marker.size = 7

    ws.add_chart(chart, "G4")

    return ws