# -*- coding: utf-8 -*-
"""
Created on Wed Jan 19 16:16:36 2022

@author: LENOVO
"""

import os
import xlwt

file_path = 'F:/py_auto_study'  # 直接从地址栏粘贴的地址是反斜杠\,python中是斜杠/
dir_list=os.listdir(file_path) # 取出文件名列表

new_workbook= xlwt.Workbook()
sheet= new_workbook.add_sheet('dir')

n=0
for i in dir_list:
    sheet.write(n,0,i) # 文件名，写入表格
    n +=1
    
new_workbook.save('dir.xlsx')