# Automation
## Overview
As my automation project, I am going to talk about a problem I had with modeling tables, the table in my third post, which you can see below as well:

![md-table](https://github.com/yashar2028/yashar/assets/148863523/81405fc9-433c-486c-af4c-8780a25f934d)

The problem is that this table is not suitable enough. Since it is just an image you can’t edit the items, add rows and columns and personally adjust it.

The possible solution for me was to use Excel to create my table and use python to rewrite the very same table in excel and save it, so I can adjust it where needed just by editing my code. In addition I can easily track a specific cell if there are hundreds of columns and rows.

This is how my Excel file looks like (located at first sheet):

![IMG_20240204_222419](https://github.com/yashar2028/yashar/assets/148863523/5394a66a-f2b5-4a14-97c2-fc21ca41b12a)


### First Part:
1-	Starting by opening an Excel file with extension “.xlsx” at our programming environment.  Having a python program that is able to read and modify Excel documents requires a module called “openpyxl”.
```
Import openpyxl
```
2-	After importing module we can load our Excel file using “.load_workbook()” function. In this case my file’s name is “Table”:
```
openpyxl.load_workbook("Table.xlsx")
```
3-	Now I can get worksheets of my workbook simply by using “.get_sheets_name()” function or call them by name using  “.get_sheet_by_name()” function. We can use “.title” to see the title of a sheet.
```
wb = openpyxl.load_workbook("Table.xlsx")
print(wb.get_sheet_names())  #This will output all the sheets: ['Sheet1', 'Sheet2', 'Sheet3']

sheet = wb.get_sheet_by_name('Sheet3')
print(sheet)                       #This will output:  <Worksheet "Sheet3">
print(sheet.title)                                                   Sheet3

```

### Second Part:
1-	I can easily track a cell in my table by using its name. Using  ‘’.value” will show the content of the cell.
```
sheet1 = wb.get_sheet_by_name('Sheet1')
print(sheet1['A2'].value)           # This will output “Intensity” which is at first column and
                                    second row as defined.
```
We can also use a function called “cell()” which has two argument row and column:
```
sheet1 = wb.get_sheet_by_name('Sheet1')
x3 = sheet1.cell(row=1, column=2).value
print(x3)      #This will output “Migraine” which is at the defined coordination.
```
2-	I can determine the size of the sheet with the Worksheet object’s “.get_highest_row()” and “.get_highest_column()” functions:
```
sheet1 = wb.get_sheet_by_name('Sheet1')
x4 = sheet1.get_highest_row()
print(x4)        #This will output “7”
x5 = sheet1.get_highest_column()
print(x5)        #This will outout “3”
```
3-	I can also easily access to all values of a specific column or row using  a worksheet object’s rows and columns attribute:
```
sheet1 = wb.get_sheet_by_name('Sheet1')
x6 = sheet1.columns[1]
print(x6)
for objects in x6:
    print(objects.value)                #This will output all the values of first column
```

### Third Part:
1-	Here I will show how to write Excel documents using python. Calling the “openpyxl.Workbook()” function to create a new, blank Workbook object with a single sheet named “Sheet”. Using “.title” we can change the default name:
```
wb = openpyxl.Workbook()
print(wb.get_sheet_names())              #This will output “['Sheet']”

sheet2 = wb.get_active_sheet()
sheet2.title = 'New Title'
print(sheet2.title)
print(wb.get_sheet_names())               #This will output “['New Title']”
```
But remember we should use “.save()” to save our changes in the document.
```
wb.save('example_copy.xlsx')              #Saved in a new Excel file called “example_copy”.
```
2-	I can create or remove sheets just by using functions “.create_sheet()”
and “.remove_sheet()”:
```
print(wb.create_sheet(index=0, title='First Sheet'))
print(wb.remove_sheet(wb.get_sheet_by_name('FirstSheet')))
```
3-	Using python I can also write values to cells and adjust the cells’ width and height:
```
import openpyxl
wb = openpyxl.Workbook()
sheet2 = wb.get_active_sheet()
sheet2['A1'] = 'VALUE'              
print(sheet2['A1'])           #After saving, A1 cell will contain a value called “VALUE”
```
Worksheet objects have row_dimensions and column_dimensions attributes that
help us to adjust widths and heights:
```
import openpyxl
wb = openpyxl.Workbook()
sheet2 = wb.get_sheet_by_name('Sheet')
sheet2.row_dimensions[1].height = 70
sheet2.column_dimensions['B'].width = 20
```

### Fourth Part:
As the last part I am going to talk about its efficiency. This type of automation is enormously useful when dealing with big data. For example you can change some values by looping through the rows or columns. However, writing a new table and assigning values to each rows and columns along with editing the font, size, etc; is much more time consuming than writing the same thing in Excel (so I wrote my table directly in Excel). As I said, it is most beneficial for quick and easy editing. By this I have my own table and I can edit it (e.g. add extra columns) so easily. In the future, I think this will help me a lot to be efficient in doing tasks by figuring out different features of this type of automating.

Note: I wrote this project with the help of book called “AUTOMATE THE BORING STUFF WITH PYTHON”.

All of the codes can be found in “Script.py”.  

