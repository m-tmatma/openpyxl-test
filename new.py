import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
wb.save('Sample1.xlsx')
