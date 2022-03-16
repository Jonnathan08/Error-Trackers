from openpyxl import Workbook, load_workbook
from openpyxl.styles import colors , fills
from openpyxl.utils import get_column_letter
from datetime import datetime

def format_excel():

    today = datetime.strftime(datetime.today(), '%m-%d-%Y')

    wb = load_workbook('error_report_%s.xlsx' % today)

    ws = wb.active

    blue = colors.Color(rgb='008698')
    filling = fills.PatternFill(patternType='solid', fgColor=blue)
    green = colors.Color(rgb='00B4CC')
    filling2 = fills.PatternFill(patternType='solid', fgColor=green)
    red = colors.Color(rgb='6A6A6A')
    filling_title = fills.PatternFill(patternType='solid', fgColor=red)

    #add colors in columns
    for col in range (4, ws.max_column+1):
        for cel in range (1, len(ws['A']) + 1):        
            if get_column_letter(col) in ['D','E','F','G','H','I','O','P','Q','R''AB','AC','AD','AE']:
                ws[get_column_letter(col) + str(cel)].fill = filling
            else:
                ws[get_column_letter(col) + str(cel)].fill = filling2

            if cel == 1:
                ws[get_column_letter(col) + str(cel)].fill = filling_title

    #adjust columns width
    for i in range(2, ws.max_column+1):
        if i < 4:
            ws.column_dimensions[get_column_letter(i)].width = 20
        else:
            ws.column_dimensions[get_column_letter(i)].width = 30

    #add filter in every column
    ws.auto_filter.ref = ws.dimensions


    wb.save('error_report_%s.xlsx' % today)

