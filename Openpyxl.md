#### create a file
```
##### from openpyxl import Workbook
wb = Workbook
```
A Workbook is always created with at least one worksheet. you can get it by using the property
```
ws = ws.active
```

##### Can also create new worksheet by using the method
```
ws1 = wb.create_sheet("MySheet") # insert at the end by default
ws2 = wb.create_sheet("MySheet2") #insert at the end by default
```
sheets are given a name (sheet1,sheet2,sheet3....
##### change the title with the title property
```
wb.title = "NewTitle"
```
##### Change the backgroup color of tab
```
ws.sheet_properties.tabColor = "RRGGBB"
```
the default color is white

#### get the title as a key of the Workbook
```
wb2 = wb['NewTitle']
```
#### 只能设置Title为key of the Workbook
```
try:
 
    ws = wb["NewTitle2"]
except:
    print('error1')
try:
    ws.title = "NewTitle2"
    ws = wb["NewTitle2"]
except:
    print("error")
    
```

##### review the names of all worksheets of the workbook
```
print(wb.sheetnames)# sheetnames is a property not method
```
#### wb is iterable
```
for sheet in wb:
  print(sheet.title)
```
####you can create copies of worksheet within a single workbook
```
source = wb.active
target = wb.copy_worksheet(source)
```
> You cannot copy worksheets **between workbooks**. You also cannot copy a worksheet if the workbook is open in read-only or write-only mode.

## Playing with data

#### Accessing one cell
Cell can be accessed directly as keys of the worksheet
```
c = ws['A4'] # this will return the cell at A4 or create one if it does not exist yet.
# print(c) >>> <Cell 'NewTitle2'.A4>
ws['A4'] = 4 # value can be direcly assigned
```
##### worksheel.cell(row,colum,value)
```
provides access to cells using row and colum notation
d = ws.cell(row = 4,colum = 2,value = 10
```
#When a worksheet is created in memory, it contains no cells. They are created when first accessed.

Because of this feature, scrolling through cells instead of accessing them directly will create them all in memory, even if you don’t assign them a value.
```
for i in range(1,101):
  for j in range(1,101):
    ws.cell(row = i,colume = j)
# creat 100*100cells in memory
```
#### Access many cells
ranges of cells can be access using slicing
```
cell_range = ws['A1':'C2']
# creat 100*100 cell_range_100_100 = ws['A1':'HD100']
```
ranges of rows or colume can be obtained using slicing
```
colC = ws['C']
col_range_C_D = ws[C:D]
row10 = ws[10]
row5_10 = ws[5:10]
```
#### 创造一块固定区域的cell，用从不同方向遍历， 一排遍历(iter_rows) 还是一列列遍历(iter_cols)
```
>>> for row in ws.iter_rows(min_row=1, max_col=3, max_row=2):
...    for cell in row:
...        print(cell)
<Cell Sheet1.A1>
<Cell Sheet1.B1>
<Cell Sheet1.C1>
<Cell Sheet1.A2>
<Cell Sheet1.B2>
<Cell Sheet1.C2>
>>> for col in ws.iter_cols(min_row=1, max_col=3, max_row=2):
...     for cell in col:
...         print(cell)
<Cell Sheet1.A1>
<Cell Sheet1.A2>
<Cell Sheet1.B1>
<Cell Sheet1.B2>
<Cell Sheet1.C1>
<Cell Sheet1.C2>
#实际上相当于 ws['A1':'C2'] 就是遍历方式不一样
```
可以iterate all through row 或者colume. （可以把所有cells装进一个tuple里）
```
tuple(ws.rows)
tuple(ws.columes)
```
## DATA storage
```
c = ws['A4']
c.value = 'hello,world'
d = ws['D5']
d.value = 3.14
print(c,c.value)
>>> <Cell 'NewTitle2'.A4>,'hello,world'
```
## Saving to a file 
```
wb = Workbook()
wb.save('balance.xlsx')
#This operation will overwrite existing files without warning.
```
#### 打开文件
wb = load_workbook('Document.xlsx')

# to save a workbook as a template
wb.template = True
wb.save('Document_tamplate.xlsx')
#模板文件是你可以在新建的时候选择以模板文件为模板创建doc.
```
#### Loading from a file
```
from openpyxl import load_workbook
wb = load_workbook('XXXX.xlsx')


