#-*- coding:utf-8 -*-
import xlrd,xlwt,openpyxl
import os.path

def save(excel,path):
    excel.save(path)
def find_re(l):
    d=dict()
    for i in l:        
        if l.count(i) >1:
            l_1=list()
            for u in range(l.count(i)):
                l_1.append(l.index(i)+1)
                l[l.index(i)]=None
            d[i]=l_1
    return d
        
def old_excel(path):
    print('暂不支持.xls')
def new_excel(path):
    excel=openpyxl.load_workbook(path)
    sheets=excel.sheetnames
    print('所有sheet:'+str(sheets))
    while(1):
        sheet=input('请输入sheet名称：')
        if sheet not in sheets:
            print('此sheet未找到')
        else:
            break
    Work=excel[sheet]
    l=list()
    for i in Work.values:
        l.append(i[0])
    d=find_re(l)
    c_n=Work.max_column
    #Work.cell(row=1,column=c_n+2,value='重复')
    #Work.cell(row=1,column=c_n+1,value='序号')
    a=2
    for i in d:
        if i==None:
            continue
        for u in range(len(d[i])):
            Work.cell(row=d[i][u],column=c_n+2,value=i)
            Work.cell(row=d[i][u],column=c_n+1,value=u+1)
            a+=1     
    d_1=dict()
    a=1
    for i in Work.values:
        if(i[0] not in d_1.keys()):
            d_1[i[0]]=a
            a+=1
    l=list()
    for i in Work.values:
        l.append(i[0])
    a=0
    for u in l:
        if u not in d.keys():
            Work.cell(row=1+a,column=c_n+1,value=d_1[u])
            Work.cell(row=1+a,column=c_n+2,value=u)
        else:
            w=Work.cell(row=1+a,column=c_n+1).value
            if w==None:
                break
            Work.cell(row=1+a,column=c_n+1,value=w/10+d_1[u])
        a+=1
    save(excel,path)
    print('查重完成，已保存')
path=input('请输入Excel位置：')
if os.path.splitext(path)[1] == '.xls':
    old_excel(path)
else:
    new_excel(path)
    
