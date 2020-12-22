import xlrd,xlwt
from openpyxl import load_workbook
from openpyxl import Workbook

#workbook = xlrd.open_workbook('数据.xls')
workbook= load_workbook("1.xlsx")
sheet1 = workbook["Sheet"]
cols = sheet1["A"]


#u第一列数据   
u=list()
for i in cols:
    if(i.value not in ['','------------------------------'] and i.value != None):
        u.append(i.value)


#y第二列数据
cols2 = sheet1["B"]
y=list()
for i in cols2:
    y.append(i.value)

#t第三列数据
cols3 = sheet1["C"]
t=list()
for i in cols3:
    t.append(i.value)


r=list()

