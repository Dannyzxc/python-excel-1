import openpyxl

#creating new excel sheet
wb = openpyxl.Workbook()

sheet = wb.active
sales = {2017:5000,2018:6500,2019:4900}

sheet['A1'] = 'Year'
sheet['B1'] = 'sales'

for k,v in sales.items():
    sheet.append((k,v))
    
    
wb.save('sales.xlsx')
