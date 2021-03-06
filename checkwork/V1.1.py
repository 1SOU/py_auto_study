# -*- coding: utf-8 -*-
"""
Created on Fri Jan 21 15:00:52 2022

@author: LENOVO

不足：1、存放统计结果的列需要手动填写 
        data_be,data-end
        请假（人次）	data_stor_coor
        合计  data_total
        出勤率 atten_coor
        分组位置   A_coor
    1.1 遍历各单位数据 for循环的范围也要调整，按理说是不用的
    2、字体格式，没统一。需后续改格式
    3、流程有点乱，
    4、缺 界面
    
    3.8修改：
    新的表头，日期合并，不再空出节假日
    列数总是多数几列
    工作日，直接数吧
    
    
    V1.1 添加新功能
    打印到word，统计每个人请假分类，每个单位


 新增  公假不计请假, per_sum()
"""



import xlwings as xw



"""初始化，"""
def _initial(work_name,sheet_name):
    
    # 打开exel
    wb= xw.Book(work_name)
    sht = wb.sheets[sheet_name]
    # print(sht.range('A1').value)
    # print(sht.range(0,0).value) 
    # sht.range('B10').value=10  # 写入数据

    # 获取行列
    info = sht.used_range
    nrow= info.last_cell.row
    ncoloumns= info.last_cell.column
    print("COLOUMNS:",ncoloumns)
    # workday= ncoloumns-work_day_adj # 工作日
    # print('工作日总计:',workday,'天')
    
    return nrow,ncoloumns,wb,sht


if __name__ == "__main__":
    
    work_name= '2月份考勤汇总表.xlsx'
    # work_name= '11、12月.xlsx'
    sheet_name= '12月'

    # work_day_adj= 8 # 直接数，不需要调整
    workday= 16
    nrow,ncoloumns,wb,sht= _initial(work_name,sheet_name)

    print('工作日总计:',workday,'天')
    print("行数：",nrow)


    
    org_list=[]  # 组织列表
    org_name_now=''
    org_flga=''
    begin=True
    org_total=0 # 部门总请假次数
    per_total=0 # 部门总人数
    # 出勤率 分组
    org_class=[]
    
    attA = {} # 100%
    attB = {}
    attC = {} 
    attD = {} # <90%
    
    # 个人请假天数求和
    def per_sum(per_data):
        sum=0
        # for i in range(len(per_data)):
        #     if per_data[i] != None :
        #         if per_data[i]!= 0: # 不是公假
        #             sum += 1
        #         # ++i   python不适用
                
                
        for i in range(len(per_data)):
            if per_data[i] != None :
                sum += 1
        
        
        return sum
    
    # 根据出勤率 分组
    def org_classify(data,attA,attB,attC,attD):
        # 考勤率分组
        # print(data)
        attence= list(data.values())
        
        if attence[0] == 1.0:
            attA.update(data) # 更新添加
            # print(attA)
        
        else:
            if attence[0]< 0.9:
                attD.update(data)
            else:
                if attence[0] < 0.95:
                    attC.update(data)
                else:
                    attB.update(data)
        
    # 对分组内 部门 按出勤率降序排序
    def sort_class(att):
        # print(att)
        return sorted(att.items(),key= lambda x:x[1], reverse=False)
    
    
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

        """需要修改 坐标"""
        data_ed='X' + str(_row)  # 改   数据结尾
        data_stor_coor='Y'+ str(_row) # 存放 个人请假次数
        data_total = 'W'+ str(_row) # '总计'
        org_attendance = '' # 计算结果，存为文本
        # attendance =0 # 
        atten_coor1 = 'Z'+ str(_row) # "出勤率："
        atten_coor2 = 'AA'+ str(_row) # "数值，float"
        
        
        per_data=sht.range(data_be,data_ed).value # 把这个人这个月的请假情况 存到list 中，再遍历list 
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
                      
                    name= org_flga
                    attence= att_num
                    org_={
                        name: attence,
    
                        }# {'不动产':0.922}
                    
                    # org_={
                    #     'name':org_flga,
                    #     'attence':att_num
                    #     } # {'name':不动产，'attence':0.922}
                    # org_class.append(org_) # 存放字典元素，字典里存 单位：出勤率
                    
    
                    org_classify(org_,attA,attB,attC,attD) # 输入这个字典元素，将其分类
                    
                    
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
        # 对计算出勤率没用，只是顺便统计所有单位
        if add_orgname:
            # print('add')
            org_list.append(org_name)
    
        
    """打印分组"""
    
    # 位置坐标
    A_coor= 'B'+str(nrow)
    B_coor= 'B'+str(nrow-1)
    C_coor= 'B'+str(nrow-2) 
    D_coor= 'B'+str(nrow-3)
    
    
    attA_sorted= sort_class(attA)
    attB_sorted= sort_class(attB)
    attC_sorted= sort_class(attC)
    attD_sorted= sort_class(attD)
    
    
    def print_class(att,coor):
        
        ct= len(att)
        count = 0
        str_data=''
        for data in att:
            str_data += data[0] 
            str_data += ':'
            str_att= str('{:.2%}'.format(data[1]))
            # str_data += data['attence']
            str_data += str_att
            count +=1
            
            if count != ct:
                str_data += '、'
            else:
                str_data += '。'
                
        sht.range(coor).value=str_data
    
    
    print_class(attA_sorted,A_coor)
    print_class(attB_sorted,B_coor)
    print_class(attC_sorted,C_coor)
    print_class(attD_sorted,D_coor)
    
   
    
        
    
    
    

    
    # wb.save(r'成绩01.xlsx')  # 加名字会另存为。不加名字直接保存当前文件
    # wb.close()  # 注意要退出