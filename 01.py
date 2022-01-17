# -*- coding: utf-8 -*-
"""
Created on Mon Jan 10 14:23:05 2022

@author: LENOVO
"""

import xlrd
import xlwt
# from xlutils.copy import copy


# 读取数据 
# xlsx = xlrd.open_workbook('01.xls')
# sheet = xlsx.sheet_by_index(0)
# data = sheet.cell_value(2,5)

# 写入
# new_workbook = xlwt.Workbook()
"""
注意，要为字符型
"""
# worksheet = new_workbook.add_sheet('new') 
# worksheet.write(0,1,'测试')
# worksheet.write(1,1,1)  # 可以直接存int
# new_workbook.save('testwr.xls')



"""计算每个人的平均成绩"""

xlsx = xlrd.open_workbook('成绩.xls')  # xlrd 无法读取xlsx文件，直接文件后缀不管用的，要新建xls文件，把数据拷进去
sheet = xlsx.sheet_by_index(0)

all_data=[]
num_set= set() #集合，函数创建一个无序不重复元素集
# 可以自动排重。新加入的数据如果已有，则不再存

# 读取数发据
for i in range(1,sheet.nrows):
    # index 从1开始，因为第一行是表头
    num= sheet.cell_value(i,0)
    name= sheet.cell_value(i,1)
    grade= sheet.cell_value(i,3) # 读取的默认是float 
    
    # 字典，key:value
    student={
        'num' : num,
        'name' : name,
        'grade' : grade,
        
        }
    # 这一条数据作为字典元素，存入到列表
    all_data.append(student)
    num_set.add(num)  # 一个人有多个成绩数据，会存入all_data。 但是名字不会再存入字典中
    
# 计算总分
sum_list= [] 
for num in num_set:
    name= ''
    sum= 0
    for student in all_data:
        # 找到这个学号num 的所有成绩，求和
        if num == student['num']:
            sum += student['grade']
            name = student['name']
            
            
    sum_stv = {
        'num' : num,
        'name' : name,
        'sum' : sum
        }
    
    sum_list.append(sum_stv)
print(sum_list)
    
    
# 写入表格
new_workbook = xlwt.Workbook()
worksheet = new_workbook.add_sheet('总成绩')
# worksheet = xlsx.sheet_by_index(1) # xlsx 是xlrd 读取出来的。xlrd只读，没有写的功能。所以不要想着用xlrd，在本表新建表单了
worksheet.write(0,0,'学号')
worksheet.write(0,1,'姓名')
worksheet.write(0,2,'总分')

for row in range(0,len(sum_list)):
    worksheet.write(row+1,0,sum_list[row]['num'])
    worksheet.write(row+1,1,sum_list[row]['name'])
    worksheet.write(row+1,2,sum_list[row]['sum'])
    # 储存到表格中后，又变成int类型了
    # 也许，本身就是float，Excel把小数点后的省略了？
    
new_workbook.save('总分.xls')  
if __name__ == '__main__':
    a=120
    # print(data)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    