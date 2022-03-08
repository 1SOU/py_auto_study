# -*- coding: utf-8 -*-
"""
Created on Tue Mar  8 14:54:32 2022

@author: LENOVO
"""


"""可以读取各种符号!!"""
import xlwings as xw

if __name__ == "__main__":
    
    work_name= "t.xlsx"
    sheet_name= "Sheet1"
    
    wb= xw.Book(work_name)
    sht= wb.sheets[sheet_name]
    
    datalist=[]
    # print(sht.range('B2').value) 
    
    for i in range(1,11):
        
        coor= 'B' + str(i)
        datalist.append(sht.range(coor).value)