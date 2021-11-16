#!/usr/bin/env python
# coding=utf-8

import sys
import xlrd
import random

num=int(input('请输入抽奖人数: ')) # 抽几个人

workbook = xlrd.open_workbook('...') # Excel地址

excel_sheet = workbook.sheet_by_index(0)

nrows = excel_sheet.nrows
ncols = excel_sheet.ncols

users = []

for i in range(0, nrows):
    users.append(excel_sheet.row(i)[ncols - 6].value)

result = random.sample(users, num)

print("中奖名单：")
for i in range(0, num):
    print(result[i])