'''
    https://qiita.com/github-nakasho/items/3f861395227e5645cce7
    https://qiita.com/github-nakasho/items/358e5602aeda81c58c81
    https://stackoverflow.com/questions/51566349/openpyxl-how-to-add-filters-to-all-columns
'''

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font

# set input file name
inputfile = 'test2.xlsx'

# read input xlsx
wb1 = openpyxl.Workbook()
ws1 = wb1.worksheets[0]


green = '92d050' # R = 146, G = 208, B = 80
yello = 'ffff00' # R = 255, G = 255, B = 0

fill      = PatternFill(patternType='solid', fgColor=yello)
fillTitle = PatternFill(patternType='solid', fgColor=green)

side = Side(style='thin', color='000000')
border = Border(top=side, bottom=side, left=side, right=side)
font = Font(name='メイリオ')

# 最初のブロック
# 絶対位置指定
for row in range(2, 10):
    for column in range(5, 10):
        cell = ws1.cell(row=row,column=column)
        if row == 2:
            cell.value = "text " + str(column)
            cell.fill = fillTitle
        else:
            cell.value = row * column
            cell.fill = fill

        cell.border = border
        cell.font = font


# 二番目のブロック
# 起点からの相対位置指定 (スクリプトの制御側の話)
offset_x = 3
offset_y = 15

for y in range(8):
    row = y + offset_y
    for x in range(5):
        column = x + offset_x
        cell = ws1.cell(row=row,column=column)
        if y == 0:
            cell.value = "text " + str(x)
            cell.fill = fillTitle
        else:
            cell.value = x * y
            cell.fill = fill

        cell.border = border
        cell.font = font

ws1.auto_filter.ref = ws1.dimensions

# save target xlsx file
wb1.save(inputfile)
