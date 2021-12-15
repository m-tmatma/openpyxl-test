import os
import sys
import openpyxl

output_xlsx = "output_ " + os.path.basename(sys.argv[0]) + '.xlsx'

wb = openpyxl.Workbook()
ws = wb.worksheets[0]
ws.title = 'testsheet'
wb.save(output_xlsx)
