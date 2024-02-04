```
import openpyxl
wb = openpyxl.load_workbook("Table.xlsx")
print(type(wb))
print(wb.get_sheet_names())

sheet = wb.get_sheet_by_name('Sheet3')
print(sheet)
print(sheet.title)



sheet1 = wb.get_sheet_by_name('Sheet1')
print(sheet1['A2'])
print(sheet1['A2'].value)



x2 = sheet1.cell(row=1, column=2)
print(x2)
x3 = sheet1.cell(row=1, column=2).value
print(x3)


x4 = sheet1.get_highest_row()
print(x4)
x5 = sheet1.get_highest_column()
print(x5)



x6 = sheet1.columns[1]
print(x6)
for objects in x6:
    print(objects.value)



# Creating a Workbook

wb = openpyxl.Workbook()
print(wb.get_sheet_names())

sheet2 = wb.get_active_sheet()
sheet2.title = 'New Title'
print(sheet2.title)
print(wb.get_sheet_names())

print(wb.create_sheet(index=0, title='First Sheet'))
print(wb.remove_sheet(wb.get_sheet_by_name('FirstSheet')))

sheet2['A1'] = 'VALUE'
print(sheet2['A1'])
sheet2.row_dimensions[1].height = 70
sheet2.column_dimensions['B'].width = 20

wb.save('example_copy.xlsx')
```
