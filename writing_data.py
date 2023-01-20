import openpyxl

sheets = 'myexel.xlsx'

current_sheet = openpyxl.load_workbook(sheets)

load_sheet_1 = current_sheet['Sheet1']
print(load_sheet_1)



print(load_sheet_1)







