# -*- coding: utf-8 -*-
"""
Created on Fri Jan 21 15:00:52 2022

@author: LENOVO

不足：1、存放统计结果的列需要手动填写 
        请假（人次）	data_stor_coor
        合计  data_total
        出勤率 atten_coor
    2、字体格式，没统一。需后续改格式
 
"""



import xlwings as xw


wb= xw.Book('11、12月.xlsx')
sht = wb.sheets['12月']
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
        if per_data[i] != None:
            sum += 1
            # ++i   python不适用
    return sum

def org_sum():
    pass

org_list=[]  # 组织列表

org_name_now=''
org_flga=''
begin=True
org_total=0 # 部门总请假次数
per_total=0 # 部门总人数

# 出勤率 分组
org_class=[]
attA = [] # 100%
attB = []
attC = [] # <90%

for _row in range(2,nrow-5):
    # 包头 不包尾  [2,3,,,,,nrow-5) 
    """xlwings，对应实际单元格的标识，从A1开始
    for 循环是从0开始，，A0所以报错！！！！！
    
    表格最后还有 统计出勤率排名，要再减去几行
    """
    
    per_leave_number=0 # 个人本月请假天数
    
    add_orgname = True
    str_row='A' + str(_row)
    org_name= sht.range(str_row).value
    data_be='C' + str(_row)
    
    data_ed='AC' + str(_row)
    
    per_data=sht.range(data_be,data_ed).value # 把这个人这个月的请假情况 存到list 中，再遍历list 
    
    data_stor_coor='AD'+ str(_row) # 个人请假次数
    data_total = 'AB'+ str(_row) # '总计'
    org_attendance = '' # 计算结果，存为文本
    attendance =0
    atten_coor1 = 'AE'+ str(_row)
    atten_coor2 = 'AF'+ str(_row)
    
    if begin:
        
        org_flga=org_name
        org_total=0 # 部门请假次数 归零
        per_total=0
             
        begin= False
        # print(1)
        #  统计此人请假天数
        
        per_leave_number= per_sum(per_data)
        sht.range(data_stor_coor).value= per_leave_number
        
        org_total += per_leave_number
        per_total +=1
        
    else:
        if org_name==org_flga:
            # 仍是同一单位
            # 统计此人请假天数 
            add_orgname=False
            per_leave_number= per_sum(per_data)
            sht.range(data_stor_coor).value= per_leave_number
            
            org_total += per_leave_number
            per_total +=1
            
            # print(2)
        else:
            if org_name==None:
            # 空行"""计算本部门出勤率"""
            # 计算此单位的总出勤率
            
                add_orgname=False
                # print(3)
                
                
                
                total_num= workday*per_total # 总出勤天数
                att_num = (total_num-org_total) / total_num
                
 
                sht.range(data_total).value=('总计：')
                sht.range(data_stor_coor).value=org_total
                sht.range(atten_coor1).value=('出勤率：')
                sht.range(atten_coor2).value=(att_num)
                
                
                """所有部门的出勤率 存到字典中
                    最后遍历字典，分组"""
                   
                org_={
                    'name': org_flga,
                    'attence': att_num
                
                    }
                org_class.append(org_) # 存放字典元素，字典里存 单位：出勤率
                
                
            else:
            # 下一个单位
                add_orgname=True
                org_total =0 # 部门请假次数 归零
                per_total =0
                org_flga=org_name
                
                per_leave_number= per_sum(per_data)
                sht.range(data_stor_coor).value= per_leave_number
                
                org_total += per_leave_number
                per_total +=1
                # print(4)
            
     
    
    
    # print(sht.range(data_stor_coor).value)
    
    # 添加新部门
    if add_orgname:
        # print('add')
        org_list.append(org_name)
        
        
# 考勤率分组
for data in org_class:
    if data['attence'] == 1.0:
        attA.append(data)
    
    else:
        if data['attence'] < 0.9:
            attC.append(data)
        else:
            attB.append(data)

# for data in attA:
#     print(data['name'],'{:.2%}'.format(data['attence']))
    
# 打印分组

# 位置坐标
A_coor= 'B'+str(nrow)
B_coor= 'B'+str(nrow-2)
C_coor= 'B'+str(nrow-3) 
# 整合分组所有数据 存为list，之后再一起打印
attA_list=[]
attB_list=[]
attC_list=[]

  # 存放要打印的数据


def print_class(att,coor):
    ct= len(att)
    count = 0
    str_data=''
    for data in att:
        str_data += data['name']
        count +=1
        
        if count != ct:
            str_data += '、'
        else:
            str_data += '。'
            
    sht.range(coor).value=str_data

print_class(attA,A_coor)
print_class(attB,B_coor)
print_class(attC,C_coor)

# attA 是个字典，，不能索引！！ attA[i] 报错 'list' object is not callable
# for i in range(len(attA)):
#     str_data += attA(i)['name']
#     if i != len(attA):
#         str_data += '、'
#     else:
#         str_data += '。'

    





# wb.save(r'成绩01.xlsx')  # 加名字会另存为。不加名字直接保存当前文件
# wb.close()  # 注意要退出