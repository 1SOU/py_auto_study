# -*- coding: utf-8 -*-
"""
Created on Tue Mar  8 15:25:39 2022

@author: LENOVO
"""



import xlwings as xw



if __name__ == "__main__":
    
    work_name= '2月份考勤汇总表.xlsx'
    sheet_name= '12月'
    
    
    wb= xw.Book(work_name)
    sht = wb.sheets[sheet_name]
    
    info = sht.used_range
    nrow= info.last_cell.row
    ncoloumns= info.last_cell.column
    
    