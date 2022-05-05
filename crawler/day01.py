#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2022/4/27 16:18
# @Author : Yisoul
# @Software: PyCharm
# @File : day01.py
# @Tallk is cheep, show me your code !

import re


a= '严格落实“13710”工作制度'
r1= re.findall('[0-9]',a)
r2= re.findall('[^0-9]',a)

print(r1)
print(r2)