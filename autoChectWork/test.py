#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2022/3/23 14:43
# @Author : Yisoul
# @Software: PyCharm
# @File : test.py
# @Tallk is cheep, show me your code !


def per_num(per_data,数据_):
    #     修改 数据[] ，影响所有用到 数据[] 的内容
    #    所以用 浅copy,不影响原数据
    数据=数据_.copy()
    for i in range(len(per_data)):

        if per_data[i] == 1 or per_data[i] ==  '1':
            数据[0] += 1
        elif per_data[i] == 0:
            数据[1] += 1
        elif per_data[i] == 2:
            数据[2] += 1
        else:
            数据[3] += 1


    return 数据

if __name__ == '__main__':
    总数据 = []
    per_data=[0,1,'1',2,3]

    数据_=[0,0,0,0]
    总数据.append(数据_)

    结果= per_num(per_data,数据_)
    总数据.append(结果)

    print('总数据',总数据)
