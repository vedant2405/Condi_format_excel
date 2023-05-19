format_condition = pivot_table.FormatConditions.Add(win32.constants.xlCellValue, win32.constants.xlGreater, "10")
format_condition.Interior.Color = win32.constants.RGB(255, 165, 0)  # Orange color
