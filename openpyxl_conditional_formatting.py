from openpyxl import Workbook
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule, Rule

wb = Workbook()
ws = wb.active

redFill = PatternFill(start_color='EE1111', end_color='EE1111', fill_type='solid')

ws.conditional_formatting.add('D1:D10', FormulaRule(formula=['NOT(ISBLANK(D1))'], stopIfTrue=True, fill=redFill))
ws.conditional_formatting.add('E1:E10', FormulaRule(formula=['ISBLANK(E1)'], stopIfTrue=True, fill=redFill))
ws.conditional_formatting.add('F1:F10', FormulaRule(formula=['F1=1'], stopIfTrue=True, fill=redFill))

ws['D1'] = "A"
ws['D2'] = "B"

ws['E1'] = "A"
ws['E2'] = "B"

ws['F1'] = 0
ws['F2'] = 1

wb.save("test.xlsx")
