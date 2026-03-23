# ... (все импорты и другие функции остаются без изменений)

def _apply_trend_colors(worksheet, trend_column_number: int, max_row: int):
    """Применяет цветовую заливку для ячеек с трендом"""
    # Цвета
    up_fill = PatternFill("solid", fgColor="FFC6EFCE")      # зеленый для восходящего
    down_fill = PatternFill("solid", fgColor="FFFFC7CE")    # красный для нисходящего
    stable_fill = PatternFill("solid", fgColor="FFFFEB9C")  # желтый для стабильного
    na_fill = PatternFill("solid", fgColor="FFE0E0E0")      # серый для недостаточно данных
    
    for row in range(4, max_row + 1):
        cell = worksheet.cell(row=row, column=trend_column_number)
        value = str(cell.value).lower() if cell.value else ""
        
        if value in ["вверх", "восходящий"]:
            cell.fill = up_fill
            cell.value = "↑ вверх"
        elif value in ["вниз", "нисходящий"]:
            cell.fill = down_fill
            cell.value = "↓ вниз"
        elif value in ["ровно", "стабильный"]:
            cell.fill = stable_fill
            cell.value = "→ ровно"
        else:
            cell.fill = na_fill
            cell.value = "? нет данных"


# ... (функция _apply_fixed_column_widths_like_example остается как в последней версии)

def create_excel_report(results: List[Dict]) -> io.BytesIO:
    """Создает Excel отчет с результатами анализа"""
    # ... (все до строки с _apply_trend_colors без изменений)
    
        # Применяем цвета для тренда
        if c_trend is not None:
            _apply_trend_colors(worksheet, c_trend, max_row)  # передаем номер колонки, а не букву

    # ... (остальной код без изменений)