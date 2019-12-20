# Machine-Learning

## Python
[100 days of ML](https://github.com/MLEveryday/100-Days-Of-ML-Code)  
[Numpy & Pandas](https://morvanzhou.github.io/tutorials/data-manipulation/np-pd/)

## xlrd, xlwt and xlutils
reference:  
[利用Python读取和修改Excel文件（包括xls文件和xlsx文件）——基于xlrd、xlwt和openpyxl模块](https://blog.csdn.net/sinat_28576553/article/details/81275650)  
[Ptyhon xlrd常用函数用法介绍](https://blog.csdn.net/BearStarX/article/details/81583415)  
[Edit existing excel workbooks and sheets with xlrd and xlwt](https://stackoverflow.com/questions/26957831/edit-existing-excel-workbooks-and-sheets-with-xlrd-and-xlwt)

Open an excel file and get one of the sheets
```python
import xlrd

wb = xlrd.open_workbook(filename="test1.xlsx")
sheet_names = wb.sheet_names()
sheet = wb.sheet_by_name(sheet_name=sheet_names[1])   # returns a xlrd.sheet.Sheet() object
# sheet = wb.sheet_by_index(1)
# sheet = wb.sheets()[1]
wb.sheet_loaded(sheet_names[1])                       # check if a sheet is fully loaded

# sheet.name
# sheet.nrows
# sheet.ncols
```

Get rows (or colons, or cells) from a sheet
```python
sheet_rows = []
for i in range(sheet.nrows):
    sheet_rows.append(sheet.row_values(i))
    
cell_value = sheet.cell_value(0,0)
cell_value = sheet.cell(0,0).value
```

Write data to excel
```python
from xlutils.copy import copy

rb = xlrd.open_workbook(filename="test2.xlsx")
wb2 = copy(rb)
sheet2 = wb2.get_sheet(0)

nrows = len(sheet_rows)
ncols = len(sheet_rows[0])
for i in range(nrows):
    for j in range(ncols):
        sheet2.write(i, j, sheet_rows[i][j])

wb2.save("test2.xls")
```

## openpyxl
Referece:
[python用openpyxl操作excel](https://blog.csdn.net/longshenlmj/article/details/51706010)  
[OpenPyXL official tutorial](https://openpyxl.readthedocs.io/en/stable/tutorial.html#playing-with-data)  

Open an excel file and get one of the sheets
```python
import openpyxl

wb1 = openpyxl.load_workbook("test1.xlsx")
wb1_sheet_names = wb1.sheetnames
wb1_sheet = wb1[wb1_sheet_names[1]]
# sheet = wb1.worksheets[0]

# wb1_sheet.title
# nrows = wb1_sheet.max_row
# ncolumns = wb1_sheet.max_column
```

Access data
```python
wb1_sheet_data = []             # access all data
for row in wb1_sheet.rows:
    row_data = []
    for cell in row:
        row_data.append(cell.value)
    wb1_sheet_data.append(row_data)
wb1_sheet_data = np.array(wb1_sheet_data)       # converting to a np array becayse 2D list cann't do index slicing

a = wb1_sheet['A2']             # access data by cell
a = wb1_sheet['A2':'D4']
a = ws.cell(row=4, column=2, value=10)
wb1_sheet['A2'] = 3

b = wb1_sheet['A']              # access data by column
b = wb1_sheet['A':'D']

c = wb1_sheet[3]                # access data by row
c = wb1_sheet[3:5]
```

Write and save data
```python
wb1_sheet_data_extracted = wb1_sheet_data[0:5, 0:3]     # extract useful data

nrows = len(wb1_sheet_data_extracted)
ncols = len(wb1_sheet_data_extracted[0])
for i in range(nrows):
    for j in range(ncols):
        wb2_sheet.cell(row=i+1,column=j+1).value = wb1_sheet_data_extracted[i,j]    # write data to each cell

wb2.save("test2.xlsx")
```
