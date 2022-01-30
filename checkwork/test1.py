# -*- coding: utf-8 -*-
"""
Created on Fri Jan 21 15:00:52 2022

@author: LENOVO

需求：1、计算
"""



import xlwings as xw


wb= xw.Book('11、12月.xlsx')
sht = wb.sheets['Sheet1']
# print(sht.range('A1').value)
# print(sht.range(0,0).value) 
# sht.range('B10').value=10  # 写入数据
data=sht.range('A1','AD1').value # list,  输出从A1-AD1的数据

info = sht.used_range
nrow= info.last_cell.row
ncoloumns= info.last_cell.column
# print(ncoloumns)
workday= ncoloumns-9 # 工作日
print('工作日总计:',workday,'天')


def per_sum(per_data):
    sum=0
    for i in range(len(per_data)):
        if per_data[i] == None:
            sum += 1
            # ++i   python不适用
    return sum

def org_sum():
    pass

org_list=[]  # 组织列表

org_name_now=''
org_flga=''
begin=True
for _row in range(2,nrow-2):
    """xlwings，对应实际单元格的标识，从A1开始
    for 循环是从0开始，，A0所以报错！！！！！
    """
    
    per_leave_number=0 # 个人本月请假天数
    
    add_orgname = True
    str_row='A' + str(_row)
    org_name= sht.range(str_row).value
    data_be='C' + str(_row)
    data_ed='AC' + str(_row)
    per_data=sht.range(data_be,data_ed).value
    data_stor_coor='AD'+ str(_row)
    if begin:
        
        org_flga=org_name
        begin= False
        # print(1)
        #  统计此人请假天数
        
        per_leave_number= per_sum(per_data)
        
    else:
        if org_name==org_flga:
            # 仍是同一单位
            # 统计此人请假天数 
            add_orgname=False
            per_leave_number= per_sum(per_data)
            # print(2)
        else:
            if org_name==None:
            # 空行
            # 计算此单位的总出勤率
                add_orgname=False
                # print(3)
            else:
            # 下一个单位
                add_orgname=True
                org_flga=org_name
                # print(4)
            
     
    sht.range(data_stor_coor).value= per_leave_number
    
    print(sht.range(data_stor_coor).value)
    
    if add_orgname:
        # print('add')
        org_list.append(org_name)
      

           
    
   

    
    






# wb.save(r'成绩01.xlsx')  # 加名字会另存为。不加名字直接保存当前文件
# wb.close()  # 注意要退出