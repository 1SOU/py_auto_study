#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2022/4/13 14:00
# @Author : Yisoul
# @Software: PyCharm
# @File : summary.py
# @Tallk is cheep, show me your code !

"""
已实现：1.打印每个人，每个部门的各类请假总计


未实现：1.右侧请假最多统计
        2.整数带小数点

后续：1.每月部门/个人的出勤率折线图
"""

import xlwings as xw

def num2res(num_data):
    """
    :param num_data: 传入数据列表【各类请假的次数】
    :return: 打印输出的文字内容
    """

    result=''
    if num_data[0] != 0:
        result += '事假'+str(num_data[0])+'天，'
    if num_data[1] != 0:
        result += '公假' + str(num_data[1]) + '天，'
    if num_data[2] != 0:
        result+= '病假'+str(num_data[2])+'天，'
    if num_data[3] != 0:
        result+= '婚假'+str(num_data[3])+'天，'
    if num_data[4] != 0:
        result+= '产假'+str(num_data[4])+'天，'
    if num_data[5] != 0:
        result+= '年休'+str(num_data[5])+'天，'
    if num_data[6] != 0:
        result+= '调休'+str(num_data[6])+'天，'
    if num_data[7] != 0:
        result+= '公休'+str(num_data[7])+'天，'
    if num_data[8] != 0:
        result+= '借调'+str(num_data[8])+'天，'
    if num_data[9] != 0:
        result+= '丧假'+str(num_data[9])+'天，'
    if num_data[10] != 0:
        result+= '旷岗'+str(num_data[10])+'天，'

    return result

def sum_org(num_data):
    """
    计算部门整个请假天数
    :param num_data:
    :return:
    """

    res_=0
    for j in range(len(num_data)):
        res_+= num_data[j]
    res1=str(res_)
    return res1

if __name__ == '__main__':

    begin= 9
    end= 295
    data_name= '4月份考勤汇总表.xlsx' #打开 出勤统计结果
    data_sht= '考勤表'
    sum_name= '考勤情况总表1.xlsx' #输出 情况汇总
    sum_sht= 'Sheet1'

    # 打开表格
    wb_data= xw.Book(data_name)
    sht_data= wb_data.sheets[data_sht]

    wb_sum= xw.Book(sum_name)
    sht_sum= wb_sum.sheets[sum_sht]

    print("打开数据表，输出表")

    print_cor=4 # 打印输出那张表的坐标
    per_count=0 # 统计部门人数
    leave_list_org = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  # 部门
    isOrg = False  # 每个部门第一行是部门 名字, 跳过之后 设为Fals
    for row in range(begin,294):
        print(row)

        print_cor+=1
        str_row = 'A' + str(row)


        if sht_data.range(str_row).value is not None: #
            # print(sht_data.range(str_row).value)
            if not isOrg :
                # A列不为空，且不是部门名字
                per_count+=1
                leave_list_per = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  # 存放个人 事假0.5，事假1，......



                # 最右边区域  各类假的次数分类统计
                cla_num_be = 'AH'+str(row)
                cla_num_end = 'AS'+str(row)
                data= sht_data.range(cla_num_be,cla_num_end).value

                print(data)
                for i in range(len(data)): #共有12列请假次数统计，AH-AS


                    # 第一列是事假0.5 ,单独计算
                    if i==0:
                        if data[0] != '':
                            # print(type(data[i]))
                            # print(data[i])
                            leave_list_per[0] += 0.5*data[0]
                            leave_list_org[0] += 0.5*data[0]
                    # 之后正常计算,注意事假0.5，和事假1 都存在leave_list_per[0]这一列！！
                    elif data[i] !='':

                        leave_list_per[i-1] += data[i]
                        leave_list_org[i-1] += data[i]
                        # 每次个人增加，部门同事也增加。但是但是到下一行，个人会清空，部门要到下一次isOrg变为True的时候清空
                print(leave_list_per)
                print(leave_list_org)
                print('')

                # 读取下一个人之前，打印这个人的数据

                print("打印个人：", print_cor)
                print('')

                res = num2res(leave_list_per)


                coord = 'C' + str(print_cor)
                sht_sum[coord].value = res


            isOrg =False  #


        else:
            # 遇到空行，下一行 为部门名

            print('打印部门：',print_cor)
            print('下个部门')
            print('')

            # res = num2res(leave_list_org)

            res = sum_org(leave_list_org)+'天，'  # 总计多少天
            res += num2res(leave_list_org)

            coord = 'C' + str(print_cor)
            sht_sum[coord].value = res



            print_cor += 1  # 因为输出表格 部门之间有一行空格，输出打印之后额外加一行
            isOrg = True
            per_count=0 # 人数归零
            # 读取进行下一行之前，打印本部门情况

            leave_list_org = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0] # 清空部门的统计





        # else:



