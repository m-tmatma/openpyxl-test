'''
    https://stackoverflow.com/questions/55309671/more-precise-image-placement-possible-with-openpyxl-pixel-coordinates-instead
'''

import os
import sys
import io
import openpyxl
import numpy as np
import matplotlib.pyplot as plt

from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU, cm_to_EMU, pixels_to_points
#from openpyxl.utils.units import EMU_to_pixels
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker

def plot(multiple):
    '''
    Make a plot to memory object.
    '''
    fig, ax = plt.subplots(figsize=(4, 3))
    x=np.array([1,2,3,4,5])
    y=np.array([2,4,5,1,2]) * multiple
    ax.plot(x, y)
    img_data = io.BytesIO()
    fig.savefig(img_data, format='png')

    img = openpyxl.drawing.image.Image(img_data)

    # https://matplotlib.org/stable/api/figure_api.html#matplotlib.figure.Figure.figimage
    # https://www.unitconverters.net/typography/inch-to-pixel-x.htm
    img.width  = fig.get_figwidth()  * fig.dpi
    img.height = fig.get_figheight() * fig.dpi
    print(img.width, img.height)
    return img

c2e = cm_to_EMU
p2e = pixels_to_EMU

def cellh(x):
    return c2e((x * 49.77)/99)

def cellw(x):
    return c2e((x * (18.65-1.71))/10)

def insert_plt(worksheet, num=10, start_row=1, start_column=1):
    '''
    Insert matplotlit plot to EXCEL sheet.
    '''
    #green = '92d050' # R = 146, G = 208, B = 80
    #yello = 'ffff00' # R = 255, G = 255, B = 0

    #fill      = openpyxl.styles.PatternFill(patternType='solid', fgColor=yello)
    #fillTitle = openpyxl.styles.PatternFill(patternType='solid', fgColor=green)

    side   = openpyxl.styles.borders.Side(style='thin', color='000000')
    border = openpyxl.styles.borders.Border(top=side, bottom=side, left=side, right=side)

    #e2p = EMU_to_pixels

    for i in range(1, num+1):
        multiple = i + 1

        img = plot(multiple)

        height = img.height
        width  = img.width
        size = XDRPositiveSize2D(p2e(width), p2e(height))
        row = i+1

        cell = worksheet.cell(row=i,column=1)
        # width unit is character counts.
        width_in_font = 2.2 * int( pixels_to_points(width) / cell.font.size + 1)
        worksheet.column_dimensions["B"].width = width_in_font
        worksheet.row_dimensions[row].height   = height

        row    = start_row + i - 1
        column = start_column


        # Calculated number of cells width or height from cm into EMUs
        # Want to place image in row 5 (6 in excel), column 2 (C in excel)
        # Also offset by half a column.

        coloffset= cellw(0.2)
        rowoffset= cellh(0.2)

        marker= AnchorMarker(col=column, row=row, colOff=coloffset, rowOff=rowoffset)
        img.anchor= OneCellAnchor(_from=marker, ext=size)
        worksheet.add_image(img)

        cell = worksheet.cell(row=i+1,column=1)
        cell.border = border

        cell = worksheet.cell(row=i+1,column=2)
        cell.border = border


output_xlsx = "output_ " + os.path.basename(sys.argv[0]) + '.xlsx'

wb1 = openpyxl.Workbook()
ws1 = wb1.worksheets[0]
insert_plt(ws1)

wb1.save(output_xlsx)
