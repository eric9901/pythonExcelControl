from openpyxl import load_workbook       #load the packect for I/O excelfile

wb = load_workbook('testcase1.xlsx')     #load the excel file
print(wb.sheetnames)              #get the sheet name that are reading

a_sheet= wb['工作表1']#get the sheet that by the sheet name
print(a_sheet)

sheet=wb.active                          #get the sheet that are reading
b4=sheet['B4']

print(f'({b4.column}, {b4.row}) is {b4.value}') # f係f-string 即Literal String Interpolatio, 可以用大括號係同一''內用var name 表示var value
b4_too = sheet.cell(row=4, column=2)            #直接用row,column 數揾個位

print(b4_too.value)
print(sheet['A2'].value)
print('----------------------------')   # print cell value by row

for row in  sheet.rows:
   for cell in row:
       print(cell.value)
print('--------------------')          # print cell value by column

for column in sheet.columns:
    for cell in column:
       print(cell.value)





