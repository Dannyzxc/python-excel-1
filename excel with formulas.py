import openpyxl

#creating new excel sheet
cs = 'sales.xlsx'
wb = openpyxl.load_workbook(cs)

sheet = wb.active

for c, d, e in sheet['C2:E4']:
    e.value = f'={c.coordinate}*{d.coordinate}'
    
wb.save(cs)
