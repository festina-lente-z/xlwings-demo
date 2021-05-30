# 读取整个sheet到pandas.DataFrame
import xlwings as xw
import pandas as pd
from pandas import Series,DataFrame
 
wb = app.books.add()
sht_All = wb.sheets[0]
 
info = sht_All.used_range
nrows = info.last_cell.row
 
def GetDataFrame(Sheets,N,M):
    index1 = Sheets.range((1,1),(1,15)).value
    index2 = Series(index1)
    Data = Sheets.range((2,1),(N,M)).value
    Data = pd.DataFrame(Data,columns=index2)
    return Data
m = GetDataFrame(sht_All,nrows,15)
 