import xlrd,xlwt,os
class read_excel():
    def read(self,path,index=0):
        book=xlrd.open_workbook(path)
        sheet=book.sheets()[index]
        return sheet

def get_order_value():
    order_value = []
    for order in Desired_Orders:
        for i in range(14, a.nrows):
            if a.cell(i, 0) == order:
                desired_value = a.cell(i, 1).value
                order_value.append([order, desired_value])
    return order_value

def get_data_dict():
    list_to_dict=[]
    for fulload in desired_fulloads:
        if fulload in title.split('_'):
            order_value = []
            for order in Desired_Orders:
                print(a.nrows)
                for i in range(14, a.nrows):
                    if a.cell(i, 0).value == order:
                        desired_value = a.cell(i, 1).value
                        print(desired_value)
                        order_value.append([order, desired_value])
            Order_value_dict=dict(order_value)
            list_to_dict.append([fulload,Order_value_dict])
    data_dict=dict(list_to_dict)
    write_to_excel(data_dict,delta_col)

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


def write_to_excel(data_dict,delta_col):
    for fulload, order_value in data_dict.items():
        for k in range(0, desired_fulloads_Len):
            if desired_fulloads[k] == fulload:
                row = -(k - desired_fulloads_Len)
        for order, value in order_value.items():
            for m in range(0, Desired_Order_Len):
                if Desired_Orders[m] == order:
                    col = m + delta_col
                    sheet1.write(row, col, value)



#输入文件目录地址
file_path=input('请按照X:/XXX/XXX的格式粘贴文件所在目录:')
#获取文件夹下文件数量
for root,dirs,file in os.walk(file_path):
    global file_list
    file_list=file

#阶次
Desired_Orders=input('输入阶次(以逗号隔开):')
Desired_Orders=input_orders(Desired_Orders)
Desired_Order_Len=int(len(Desired_Orders))
#负载
Desired_fulloads=input('输入负载(以逗号隔开):')
desired_fulloads=input_fulloads(Desired_fulloads)
desired_fulloads_Len=int(len(desired_fulloads))
print(desired_fulloads)
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

for file_name in file_list:#根据文件数量读取文件
    # print (file_name)
    path=file_path+'\\'+file_name#读取文件路径
    # print(path)
    a=read_excel().read(path)
    title=a.cell(3,1).value#获取当前文件第4行第二列的内容
    check_direction=title.find('CCW')#判断文件所指数据的转向
    # print(check_direction)
    if check_direction==-1:
        current_direction = 'CW'
        delta_col = 2 + Desired_Order_Len
        get_data_dict()
    else:
        current_direction = 'CCW'
        delta_col = 1
        get_data_dict()
save_name=input('请输入输出报告的文件名:')
for root,dir,files in os.walk('D:/report/'):
    global report_files
    report_files=files
while True:
    if save_name in report_files:
        print('文件名重复')
    else:
        book.save('D:/report/'+save_name+'.xls')
        print('文件保存成功')
        break
