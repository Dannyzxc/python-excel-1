import openpyxl
from openpyxl.styles import *

#creating new excel sheet
cs = 'myexel.xlsx'
wb = openpyxl.load_workbook(cs)


load_sheet_1 = wb['Sheet_1']
load_sheet_2  = wb['Sheet2']


cell_a2 = load_sheet_1['A4']

print(dir(openpyxl.styles))

font = Font(name='Tahoma',size=16, bold=True, strike=False)

cell_a2.font = font


#fill = PatternFill(fill_type=('solid'),fgColor = colors.YELLOW )

#cell_a3.fill = fill

#border_green = Side(border_style= ' double',color = 'FF0000')

#thin_border = Side(border_style = 'thin', color='FF0000')

#cell_a2.border = Border(left=border_green) 

alignment = Alignment(horizontal='right',vertical= 'center')
cell_a2.alignment = alignment

wb.save(cs)