#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2022/3/22 17:20
# @Author : Yisoul
# @Software: PyCharm
# @File : switcher.py
# @Tallk is cheep, show me your code !

"""字典，手动实现 switch"""

事假=1
公假=0

def func1():
    global 事假
    print('1', 事假)
    事假 += 1
    return 事假
def func2():
    global 公假
    print('2',公假)
    公假 += 1
    return 公假

def get_default():
    return 'other'
switcher= {
    1:func1(),
    2:func2()
}
# out=2
switcher.get(1,get_default())
print(事假)
switcher.get(1,get_default())
print(事假)
# print('switcher1:',switcher.get(out,get_default()))
# print('switcher2:',switcher.get(out,get_default()))


