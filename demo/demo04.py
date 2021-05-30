import xlwings as xw
 
wb = xw.Book(r'F:\PythonData\xlwings\Style.xlsx')
 
sht = wb.sheets[0]
 
sht_color = sht.range((1,1)).color
print(sht_color)
#(255, 153, 255)
sht.range((3,1)).color = (255, 153, 255)
#A3背景颜色为粉色
sht_BoldA = sht.range((1,1)).api.Font.Bold
print(sht_BoldA)
#True
sht_BoldB = sht.range((1,2)).api.Font.Bold
print(sht_BoldB)
#False
sht.range((3,1)).value = 'A3'
sht.range((3,1)).api.Font.Bold = True
#加粗
sht_Fontstyle = sht.range((1,2)).api.Font.FontStyle
print(sht_Fontstyle)
#倾斜
sht.range((3,2)).value = 'B3'
sht.range((3,2)).api.Font.FontStyle = "倾斜"
#设置为斜体
sht_Underline = sht.range((1,3)).api.Font.Underline
print(sht_Underline)
#2,为下划线
sht.range((3,3)).value = 'C3'
sht.range((3,3)).api.Font.Underline = 2
#设置下划线
sht_style = sht.range((1,1),(1,5)).api.Borders.LineStyle
print(sht_style)
#1
#全框线
sht.range((3,1),(3,3)).api.Borders.LineStyle = 1
#设置全框线
sht_HA_A1 = sht.range((1,1)).api.HorizontalAlignment
print(sht_HA_A1)
#水平左对齐
#1
sht_HA_A2 = sht.range((1,2)).api.HorizontalAlignment
print(sht_HA_A2)
#水平居中
#-4108
sht_HA_A5 = sht.range((1,5)).api.HorizontalAlignment
print(sht_HA_A5)
#水平右对齐
#-4152
sht_VA_A3 = sht.range((1,3)).api.VerticalAlignment
print(sht_VA_A3)
#垂直靠上
#-4160
sht_VA_A4 = sht.range((1,4)).api.VerticalAlignment
print(sht_VA_A4)
#垂直居中
#-4108
sht_VA_A5 = sht.range((1,5)).api.VerticalAlignment
print(sht_VA_A5)
#垂直靠下
#-4107
 
wb.save()
xw.App().quit()