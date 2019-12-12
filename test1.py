import xlrd,xlwt,os
import numpy as np
from pylab import *
from mpl_toolkits.mplot3d import Axes3D

class ReadExcel():
    def read(self,path,index=0):
        excelbook=xlrd.open_workbook(path)
        sheet=excelbook.sheet_by_index(index)
        return sheet
class GetDataDict():
    def getorderdict(self,sheet):
        orderlist=[]
        for row in range(14,sheet.nrows):
            orderlist.append([sheet.cell_value(row,0),sheet.cell_value(row,1)])
        orderdict=dict(orderlist)
        return orderdict
    def getloaddict(self,sheet,loadlist):
        for load in loads:
            if load in mark_load:
                orderdict=self.getorderdict(sheet)
                loadlist.append([load,orderdict])

loads=["{num:.0%}".format(num=load/100) for load in range(10,101,10)]
ccwloadlist=[]
cwloadlist=[]
direcdict={}
for root,dir,files in os.walk('D:\\Torque ripple'):
    for file in files:
        path=root+'\\'+file
        sheet1=ReadExcel().read(path)
        mark_direction=sheet1.cell_value(3,1).split('_')[-1].find('CCW')
        mark_load=sheet1.cell_value(3,1).split('_')
        if mark_direction==-1:
            GetDataDict().getloaddict(sheet1,cwloadlist)
        else:
            GetDataDict().getloaddict(sheet1,ccwloadlist)

ccwloaddict=dict(ccwloadlist)
cwloaddict=dict(cwloadlist)
load_value=[]
loadpercent=[]
positive_orders=[]
newbook=xlwt.Workbook(encoding='utf-8')
newsheet=newbook.add_sheet('sheet1')

i=0
j=0
for load in loads:
    ccwlist = []
    cwlist = []
    for value in ccwloaddict[load].values():
        ccwlist.append(value)
        ccwlist.reverse()
    for value in cwloaddict[load].values():
        cwlist.append(value)
    valuelist=ccwlist+cwlist
    load_value.append(valuelist)
    print([load,valuelist])
    for value in valuelist:
        print(value)
        newsheet.write(i, j, value)
        j+=1
    i+=1
newbook.save('D:/report/report_2019.xls')
for load in loads:
    x=float(load.strip('%'))
    num=x/100
    loadpercent.append(num)

for key in ccwloaddict[loads[0]].keys():
    positive_orders.append(key)
negtive_orders=[-order for order in positive_orders]
negtive_orders.reverse()
orders=negtive_orders+positive_orders

Y=loadpercent
X=orders

Z=np.array(load_value)

fig=figure()
ax=Axes3D(fig)
X,Y=meshgrid(X,Y)
ax.plot_surface(X, Y, Z, rstride=1, cstride=1, cmap=plt.cm.hot)
ax.contourf(X, Y, Z, zdir='z', offset=-2, cmap=plt.cm.hot)
ax.set_zlim(-0.01,1)
show()



