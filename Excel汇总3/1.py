import xlrd
import xlwt
import time
import sys
#import win32com.client as win32
import openpyxl
#import warnings

def input1(path1):
    workbook=xlrd.open_workbook(path1)
    sheets=workbook.sheet_names()
    d=dict()
    for i in sheets:
        sheet=workbook.sheet_by_name(i)
        num_rows=sheet.nrows
        num_cols=sheet.ncols
        for u in range(num_rows):
            data=dict()
            row_v=sheet.row_values(u)
            name=row_v[14]
            if(name in d.keys()):
                data=d[name]
            area_name=row_v[2]
            data[area_name]=[row_v[5],row_v[13]]
            data['cq']=row_v[16]
            d[name]=data
    return d

def xls_to_xlsx(path):
    fname = path
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat = 51)
    wb.Close()
    excel.Application.Quit()
    
def output1(d1,path2):
    d=dict()
    wb = openpyxl.load_workbook(path2)
    for i in wb.worksheets:
        sheet=i
        rows=list(sheet.rows)
        for i in range(5,len(list(sheet.rows))):
            try:
                if(rows[i][1].value!=None):
                    name=rows[i][1].value
                    rows[i][44].value=d1[name]['cq']
                area_name=rows[i][20].value
                rows[i][23].value=d1[name][area_name][1]
                rows[i][42].value=d1[name][area_name][0]
            except:
                continue
    wb.save(r'汇总完成.xlsx')

input(''''
请将 汇总结果.xls ，与本软件放在同一目录内。
请将 调查表.xls
另存为 .xlsx 后缀文件，然后放在与本软件同一目录内。
摁回车继续……

''')
path2=input('请输入调查表的绝对路径:')
path1=input('请输入汇总结果绝对路径:')
d=dict()
d=input1(path1)
#xls_to_xlsx(path)
print('正在汇总，请稍等……')
output1(d,path2)
input('汇总完成！')
input()
