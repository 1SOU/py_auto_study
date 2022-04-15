#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2022/4/13 14:00
# @Author : Yisoul
# @Software: PyCharm
# @File : summary.py
# @Tallk is cheep, show me your code !

import xlwings as xw

# def _initial(wb_name1,wb_name2,sht1,sht2):
#     wb1 =xw.Book(wb_name1)
#     s

if __name__ == '__main__':

    begin= 9
    end= 295
    data_name= '4月份考勤汇总表.xlsx' # 出勤统计结果
    data_sht= '考勤表'
    sum_name= '考勤情况总表.xlsx' # 情况汇总
    sum_sht= 'Sheet1'

    # 打开表格
    wb_data= xw.Book(data_name)
    sht_data= wb_data.sheets[data_sht]

    wb_sum= xw.Book(sum_name)
    sht_sum= wb_sum.sheets[sum_sht]

    for row in range(begin,end):
        leave_list_per=[0,0,0,0,0,0,0,0,0,0,0] # 存放 事假0.5，事假1，......
        leave_list_org = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]


        str_row= 'A'+str(row)
        if sht_data.range(str_row).value is not None:
            for i in range(12):

        else:







