import xlwings as xw
 
wb = xw.Book(r'F:\xlwings\OriginalData.xlsx')
#or
wb1 = xw.books.open(r'F:\xlwings\OriginalData01.xlsx')
#打开文件
 
wb.save()
#保存原文件
wb1.save(r'F:\xlwings\PresentData01.xlsx')
#另存为PresentData01.xlsx