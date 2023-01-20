import openpyxl

#creating new excel sheet
cs = 'myexcel.xlsx'
wb = openpyxl.load_workbook(cs)

print(wb)
sheet = wb.active


