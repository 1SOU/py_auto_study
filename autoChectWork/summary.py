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


增删人员需要改动下列参数：
要用到坐标的：
1.工作日出勤数据开始-结束范围 每一行、  好像EXCEL 里自动算出勤率，用不到了
2.请假次数分类统计开始-结束、   cla_num_be
3.出勤率百分比所在坐标         _att_coor
以上基本不变    不变列，只变行（增减人员）
4.全部人员到296还是295，      end=?
5.出勤率排名分组坐标          A_coor

2022.06.02

每月出结果时候，增删人员替换旧空白表，复制一份“考勤情况总表-空白”，在空白表副本上输出防止数据污染

不用了，督查科直接给增删过的空白表
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


def sort_class(att):
    # print(att)
    return sorted(att.items(),key= lambda x:x[1], reverse=False)

def print_class(att, coor, sht):

    ct = len(att)
    count = 0
    str_data = ''
    for data in att:
        str_data += data[0]
        str_data += ':'
        str_att = str('{:.2%}'.format(data[1]))
        # str_data += data['attence']
        str_data += str_att
        count += 1

        if count != ct:
            str_data += '、'
        else:
            str_data += '。'

    sht.range(coor).value = str_data


if __name__ == '__main__':

    begin= 9
    end= 291   # 出勤率排名  这一行
    data_name= '7月份考勤情况.xlsx' #打开 出勤统计结果
    data_sht= '考勤表'
    sum_name= '7月份考勤汇总.xlsx' #输出 情况汇总
    sum_sht= 'Sheet1'

    """打印出勤率所用的坐标"""
    # 存放部门名字，出勤率  的坐标
    org_coor = [8] # 获取部门名字，每次遇到空格，下一行就是部门名字
    # 因为设置的不是从第一个部门名开始，所以默认添加上

    att_coor = [] # 获取部门出勤率，每次遇到空格
    # 分组数据，未排序
    attA = {}
    attB = {}
    attC = {}
    attD = {}
    # 要打印分组的  坐标
    A_coor = 'B295'
    B_coor = 'B294'
    C_coor = 'B293'
    D_coor = 'B292'

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

    for row in range(begin,end):
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
                cla_num_end = 'AT'+str(row)
                data= sht_data.range(cla_num_be,cla_num_end).value

                print(data)
                for i in range(len(data)):
                    if data[i] == '':
                        data[i] = 0
                if data[0] != '' or data[1] != '':
                    leave_list_per[0] += 0.5 * data[0]
                    leave_list_org[0] += 0.5 * data[0]
                    leave_list_per[1] += 0.5 * data[1]
                    leave_list_org[1] += 0.5 * data[1]
                for i in range(2,len(data)): #共有13列请假次数统计，AH-At ,



                    '''错误写法：i=0,1 两次循环，累加两次'''
                    # if i==0 or i==1:
                    #     # print('计算 半天')
                    #     if data[0] != '':
                    #         # print(type(data[i]))
                    #         # print(data[i])
                    #         # leave_list_per[0] += 0.5*data[0]
                    #         leave_list_per[0] += 0.5*data[0]
                    #         leave_list_org[0] += 0.5*data[0]
                    #         print('计算0-')
                    #         print(leave_list_per)
                    #
                    #     if data[1] != '':
                    #
                    #         leave_list_per[1] += 0.5 * data[1]
                    #         leave_list_org[1] += 0.5 * data[1]


                    # 之后正常计算,注意事假0.5，和事假1 都存在leave_list_per[0]这一列！！
                    # 公假都在 leave_list_per[1]




                    if data[i] !='':
                        #     data= [0-,1-,0,1,2,3,4,5,6,7,8,9,10,11,12] 15
                        # leave_list_per= [0,1,2,3,4,5,6,7,8,9,10,11,12] 13
                        # data 从第三列开始，把数据存到 leave_list_per第一列
                        leave_list_per[i-2] += data[i]
                        leave_list_org[i-2] += data[i]
                        # 每次个人增加，部门同时也增加。但是但是到下一行，个人会清空，部门要到下一次isOrg变为True的时候清空
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

            # 获取下个部门名所在坐标，本部门出勤率所在坐标，
            """因为最后一行空格，可能会导致最后排名中多出一个无数据的部门"""
            org_coor.append(row+1)
            att_coor.append(row)

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



    org_coor.pop()
    """出勤率排序"""
    if len(org_coor) == len(att_coor):
        print("开始排序")
        for i in range(len(org_coor)):

            _name_coor= 'A'+str(org_coor[i])
            _att_coor= 'AN'+str(att_coor[i])
            org_name= sht_data.range(_name_coor).value
            org_att= sht_data.range(_att_coor).value
            data={
                org_name:org_att
            }

            if org_att == 1.0:
                attA.update(data)
            else:
                if org_att < 0.9:
                    attD.update(data)
                else:
                    if org_att < 0.95:
                        attC.update(data)
                    else:
                        attB.update(data)

        print("分完组，开始排序")
        # print(attA, attB, attC, attD)

        # 排序后的分组，下一步打印
        attA_sorted = sort_class(attA)
        attB_sorted = sort_class(attB)
        attC_sorted = sort_class(attC)
        attD_sorted = sort_class(attD)

        print_class(attA_sorted, A_coor,sht_data)
        print_class(attB_sorted, B_coor,sht_data)
        print_class(attC_sorted, C_coor,sht_data)
        print_class(attD_sorted, D_coor,sht_data)

        print("over!")
    else:
        print("error")



