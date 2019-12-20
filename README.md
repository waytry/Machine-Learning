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

