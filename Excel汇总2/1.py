import xlrd
import xlwt
import time
import sys
def input1():
    workbook1 = xlrd.open_workbook('重庆市--丰都县保合镇发包方信息.xls')
    sheets1=workbook1.sheet_names()
    d=dict()
    for i in sheets1:
        sheet=workbook1.sheet_by_name(i)
        num_rows=sheet.nrows
        num_cols=sheet.ncols
        for u in range(num_rows):
            flag=sheet.row_values(u)
            d[flag[0]]=flag
    return d

def input2():
    workbook1 = xlrd.open_workbook('重庆市--丰都县保合镇承包方信息.xls')
    sheets1=workbook1.sheet_names()
    d=dict()
    for i in sheets1:
        sheet=workbook1.sheet_by_name(i)
        num_rows=sheet.nrows
        num_cols=sheet.ncols
        for u in range(num_rows):
            flag=sheet.row_values(u)
            d[flag[0]]=flag
    return d

def input3():
    workbook1 = xlrd.open_workbook('重庆市--丰都县保合镇承包合同信息.xls')
    sheets1=workbook1.sheet_names()
    d=dict()
    for i in sheets1:
        sheet=workbook1.sheet_by_name(i)
        num_rows=sheet.nrows
        num_cols=sheet.ncols
        for u in range(num_rows):
            flag=sheet.row_values(u)
            d[flag[3]]=flag[0]
    return d
def output(l):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    a=0
    for i in range(len(l)):
        if(i%65535==0):
            a=a+1
            worksheet = workbook.add_sheet('Sheet'+str(a))
        for u in range(len(l[i])):
            worksheet.write(i-(a-1)*65535,u, label =l[i][u])
    workbook.save('汇总结果.xls')

if(time.time()==1541670325+60*60*1):
    sys.exit(0)
input('''请确保 “重庆市--丰都县保合镇承包地块信息.xls”、
“重庆市--丰都县保合镇承包合同信息.xls”、
“重庆市--丰都县保合镇承包方信息.xls”、
“重庆市--丰都县保合镇发包方信息.xls”
这四个表格文件与本软件在同一目录。
汇总结果保存 为“汇总结果.xls”。
摁回车继续……''')
print('正在汇总 请稍等……')
workbook = xlrd.open_workbook('重庆市--丰都县保合镇承包地块信息.xls')
sheets=workbook.sheet_names()
d=input1()
d2=input2()
d3=input3()
l=list()
for i in sheets:
    sheet=workbook.sheet_by_name(i)
    num_rows=sheet.nrows
    num_cols=sheet.ncols
    for u in range(num_rows):
        flag=sheet.row_values(u)
        if(flag[0]==''):
            continue
        try:
            flag[0]=d[flag[0]][1]
            flag[14]=d2[flag[14]][2]
            flag.append(d3[flag[14]])
        except KeyError:
            #print('未找到信息，跳过')
            pass
        l.append(flag)
#l=l[0:10000]
output(l)
input('汇总完成，已保存为 汇总结果.xls')
        
