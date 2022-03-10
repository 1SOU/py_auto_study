# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import xlwings as xw
from docx import Document



if __name__ == '__main__':

    # wb= xw.Book("t.xlsx")
    # sht= wb.sheets['Sheet1']
    #
    #
    # lizi=[]
    # for i in range(1,13):
    #     coor= 'B'+str(i)
    #     lizi.append(sht.range(coor).value)
    #     # print(sht.range(coor).value)
    #     print(lizi)

    # # data= sht.range('C13').value + sht.range('D13').value
    # data = sht.range('c13').value
    # if (data==0.5):
    #     print('ok')

    wb = xw.Book("2月份考勤汇总表.xlsx")
    sht = wb.sheets['12月']

    data = sht.range('c25').value
    data1 = sht.range('d25').value
    data2= sht.range('d32').value
    # if (sht.range('C25').value == '0.5'):
    #     print('ok')
    # else:
    #     print(sht.range('C25').value)


    # wb.save()
    # wb.close()




