import openpyxl

wb = openpyxl.Workbook()
ws = wb.worksheets[0]
ws.title = 'testsheet'
wb.save('Sample1.xlsx')
