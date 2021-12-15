import os
import sys
import openpyxl

output_xlsx = "output_ " + os.path.basename(sys.argv[0]) + '.xlsx'

wb = openpyxl.Workbook()
ws = wb.active
wb.save(output_xlsx)
