import xlrd,xlwt,os
class read_excel():
    def read(self,path,index=0):
        book=xlrd.open_workbook(path)
        sheet=book.sheets()[index]
        return sheet

def get_data_dir():
    if Judge[-3] in desired_fulloads:
        dir2 = {}
        for desired_fulload in desired_fulloads:
            if desired_fulload in Judge:
                dir3 = {}
                current_fulload = desired_fulload
                for Order in Desired_Orders:
                    for i in range(14, a.nrows):
                        if a.cell(i, 0).value == Order:
                            desired_row = i
                            desired_value = a.cell(desired_row, 1).value#读取当前列的数据
                            dir3[Order] = desired_value
                dir2[current_fulload] = dir3
                write_data_to_excel(dir2,delta_col)

def input_orders(Input_Orders,Default_Orders=[1,2,3,4,5]):
    if Input_Orders=='':
        return Default_Orders
    else:
        Desired_Orders=[]
        buffer1=Input_Orders.split(',')
        for value in buffer1:
            Desired_Orders.append(int(value))
        return Desired_Orders

def input_fulloads(Input_fulloads,Default_fulloads=['0%', '20%', '40%', '60%', '80%', '100%']):
    if Input_fulloads=='':
        return Default_fulloads
    else:
        Desired_fulloads=Input_fulloads.split(',')
        return Desired_fulloads


def write_data_to_excel(dir2,delta_col):
    for fulload, vlue in dir2.items():
        for k in range(0, desired_fulloads_Len):
            if desired_fulloads[k] == fulload:
                row = -(k - desired_fulloads_Len)
        for order, value in vlue.items():
            # print(current_direction, fulload, order, value)
            for m in range(0, Desired_Order_Len):
                if Desired_Orders[m] == order:
                    col = m + delta_col
                    sheet1.write(row, col, value)

def check_direction(Judge):
    if 'CCW.hdf' in Judge:
        current_direction = 'CCW'
    elif 'CW.hdf'in Judge or 'CW1.hdf' in Judge:#为了这个我要疯
        current_direction = 'CW'
    else:
        current_direction = 'CCW'
    return current_direction

#输入文件目录地址
file_path=input('请按照X:/XXX/XXX的格式粘贴文件所在目录:')
#获取文件夹下文件数量
for root,dirs,file in os.walk(file_path):

    file_list=file
    global file_number
    file_number=len(file_list)
    print ('file number=',file_number)
#阶次
Desired_Orders=input('输入阶次(以逗号隔开):')
Desired_Orders=input_orders(Desired_Orders)
Desired_Order_Len=int(len(Desired_Orders))
#负载
Desired_fulloads=input('输入负载(以逗号隔开):')
desired_fulloads=input_fulloads(Desired_fulloads)
desired_fulloads_Len=int(len(desired_fulloads))
#创建excel
book = xlwt.Workbook(encoding='utf-8')
sheet1 = book.add_sheet('sheet1')
style = xlwt.XFStyle()
sheet1.write(0,0,'CCW')
sheet1.write(0,1+Desired_Order_Len,'CW')

for i in range(0,desired_fulloads_Len):
    a=desired_fulloads[-i-1]+' load'
    sheet1.write(i+1,0,a)
    sheet1.write(i+1,1+Desired_Order_Len,a)
for j in range(0,Desired_Order_Len):
    sheet1.write(0,j+1,Desired_Orders[j])
    sheet1.write(0,j+2+Desired_Order_Len,Desired_Orders[j])

for i in range(1,int(file_number)+1):#根据文件数量读取文件
    path=file_path+'.'+str(i)+'.xlsx'#读取文件路径
    a=read_excel().read(path)
    title=a.cell(3,1).value
    Judge=title.split("_")
    current_direction=check_direction(Judge)
    if current_direction == 'CCW':
        delta_col=1
        get_data_dir()
    if current_direction == 'CW':
        delta_col=2+Desired_Order_Len
        get_data_dir()
save_name=input('请输入输出报告的文件名:')
for root,dir,file in os.walk('D:/report/'):
    global file_names
    file_names=file
while True:
    if save_name in file_names:
        print('文件名重复')
    else:
        book.save('D:/report/'+save_name+'.xls')
        print('文件保存成功')
        break
