'''
    https://qiita.com/github-nakasho/items/3f861395227e5645cce7
    https://qiita.com/github-nakasho/items/358e5602aeda81c58c81
    https://stackoverflow.com/questions/51566349/openpyxl-how-to-add-filters-to-all-columns
'''

import os
import sys
import openpyxl

output_xlsx = "output_ " + os.path.basename(sys.argv[0]) + '.xlsx'

wb = openpyxl.Workbook()
ws = wb.active
wb.save(output_xlsx)
