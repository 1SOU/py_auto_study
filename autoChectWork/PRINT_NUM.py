#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2022/3/15 18:31
# @Author : Yisoul
# @Software: PyCharm
# @File : PRINT_NUM.py
# @Tallk is cheep, show me your code !

"""
    添加新功能
    打印到word，统计每个人请假分类，每个单位
"""


import xlwings as xw
from docx import Document




if __name__ == '__main__':
    excel_name= '2月份考勤汇总表.xlsx'
    sheet_name = '12月'
    wb= xw.Book(excel_name)
    sht = wb.sheets[sheet_name]

    doc = Document()  # 新建文本

    标题= sht.range('A1').value
    print(标题)
    doc.add_heading(标题)
    doc.save('0.docx')
