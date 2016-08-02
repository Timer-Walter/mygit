# -*- coding: utf-8 -*
import xlrd;import xlwt
bottomNumber = 14120776;topNumber = 14120811
tablename = []
for i in range(bottomNumber-1,topNumber):
    tablename.append(i+1)
    try:
        data = xlrd.open_workbook('%d.xls' %(i+1))
    except IOError:
        tablename.remove(i+1)
data0 = xlrd.open_workbook('%d.xls' %tablename[0])
sheet0 = data0.sheets()[0]
col = sheet0.ncols
head =[]
for i in range(col):
    head.append(sheet0.cell(0,i).value)
information = []
for i in range(len(tablename)):
    data = xlrd.open_workbook('%d.xls' %tablename[i])
    sheet = data.sheets()[0]
    pieceinfor = []
    for j in range(col):
        pieceinfor.append(sheet.cell(1,j).value)
    information.append(pieceinfor)
workbook1 = xlwt.Workbook(encoding = 'ascii')
worksheet1 = workbook1.add_sheet('My Worksheet')
for i in range(col):
    worksheet1.write(0,i, label = head[i])
    for j in range(len(tablename)):
        worksheet1.write(j+1,i, label = information[j][i])
workbook1.save('Excel_new.xls')