import xlwings as xw
 
wb = xw.Book()
#新建一个工作表
sht = wb.sheets[0]
#shee1
sheet_name = 'NEWSHEET'
sht.name = sheet_name
#更改第一个sheet名字
col_a = [1,2,3,4,5,6,7]
sht.range('A1:A7').options(transpose=True).value = col_a
#整列赋值
sht.api.Columns(1).Insert()
#在第一列前插入一列
sht.api.Rows(1).Insert()
#在第一行前插入一行
sht.range('A3:A4').api.EntireRow.Delete()
#删除3,4行
sht.api.Columns(2).Copy(sht.api.Columns(1))
#复制第二列到第一列，可以带格式复制
sht.range('B1').api.EntireColumn.Delete()
#删除第二列B列
wb.save(r'F:\PythonData\xlwings\NewData.xlsx')
xw.App().quit()
#退出整个excel，不写的话打开excel会显示被其他人使用