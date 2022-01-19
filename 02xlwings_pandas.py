# -*- coding: utf-8 -*-
"""
Created on Tue Jan 18 18:39:37 2022

@author: LENOVO
"""

import xlwings as xw

'''
执行 xw.Book()，会自动打开office Excel
'''
# wb= xw.Book() # 新建一个文件
wb= xw.Book('成绩.xls') # 打开已有的文件

# 读写
sht = wb.sheets['Sheet1']
print(sht.range('A1').value)  
sht.range('B10').value=10
sht.range('B11').value='测试'
# 存
wb.save(r'成绩01.xlsx')  # 加名字会另存为。不加名字直接保存当前文件
wb.close()  # 注意要退出




'''pandas'''

'''plt'''