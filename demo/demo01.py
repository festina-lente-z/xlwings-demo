# 获取有数据的所有行数和列数
import xlwings as xw
 
wb = xw.Book(r'F:\xlwings\OriginalData.xlsx')
 
sht = wb.sheets[0]
 
info = sht.used_range
 
nrows = info.last_cell.row
print(nrows)
 
ncolumns = info.last_cell.column
print(ncolumns)