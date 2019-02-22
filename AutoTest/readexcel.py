import xlrd
data = xlrd.open_workbook('D:/Program Files (x86)/test.xlsx')
table = data.sheets()[0]
row=table.nrows

print(row)
#for n in rangrowe(1,row):
            #tmpdict = {}                                                #把一行记录写进一个{}
            #tmpdict['id'] = n                                           #n是Excel中的第n行
            #tmpdict['name'] = table.cell(n,2).value
print(table.cell(0,0).value)


