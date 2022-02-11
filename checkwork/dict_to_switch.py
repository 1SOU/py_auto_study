# -*- coding: utf-8 -*-
"""
Created on Wed Feb  9 17:12:29 2022

@author: LENOVO

用字典实现 switch case

"""


def num_to_str(num):
    numbers={
        0:'zeros',
        1:'one'
        }
    return numbers.get(num,'没有') # 第二个参数是 当输入的num在字典中不存在的时候，输出的代替



if __name__ == "__main__":
    print(num_to_str(1))
    print(num_to_str(5))