#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2022/3/29 13:08
# @Author : Yisoul
# @Software: PyCharm
# @File : sort.py
# @Tallk is cheep, show me your code !


import xlwings as xw

if __name__ == '__main__':
    work_name= "动态考勤表.xlsx"
    sheet_name= "考勤表"

    wb = xw.Book(work_name)
    sht = wb.sheets[sheet_name]

    # 存放部门名字，出勤率  的坐标
    org_coor=[6,24,29,34]
    att_coor=[23,28,33,38]
    # 分组数据，未排序
    attA = {}
    attB = {}
    attC = {}
    attD = {}
    # 要打印分组的  坐标
    A_coor = 'B45'
    B_coor = 'B44'
    C_coor = 'B43'
    D_coor = 'B42'



    # t=sht.range('AM33').value
    # print(t)

    def sort_class(att):
        # print(att)
        return sorted(att.items(),key= lambda x:x[1], reverse=False)


    def print_class(att, coor):

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

    if len(org_coor) == len(att_coor):
        for i in range(len(org_coor)):
            _name_coor= 'A'+str(org_coor[i])
            _att_coor= 'AM'+str(att_coor[i])
            org_name= sht.range(_name_coor).value
            org_att= sht.range(_att_coor).value
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
        print(attA, attB, attC, attD)

        # 排序后的分组，下一步打印
        attA_sorted = sort_class(attA)
        attB_sorted = sort_class(attB)
        attC_sorted = sort_class(attC)
        attD_sorted = sort_class(attD)

        print_class(attA_sorted, A_coor)
        print_class(attB_sorted, B_coor)
        print_class(attC_sorted, C_coor)
        print_class(attD_sorted, D_coor)

        print("over!")

