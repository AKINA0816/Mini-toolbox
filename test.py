import xlrd

from datetime import date, datetime

file = 'xlsx/to8664.xlsx'

print(file)

wb = xlrd.open_workbook(filename=file)  # 打开文件
# print(wb.sheet_names())#获取所有表格名字
sheet1 = wb.sheet_by_index(1)  # 通过索引获取表格

print(sheet1.cell(737, 4).value)
print(sheet1.cell(737, 4).value == "")
print(type(sheet1.cell(737, 4).value))
# 想要获取第八行的商品名称
row = 8
col = 1  # B列
print(sheet1.cell(row - 1, col).value)  # 获取表格里的内容，三种方式
print(sheet1.nrows)

import os
allFile = os.listdir("/Users/chen/Desktop/商品图片")
for i in range(len(allFile)):
    allFile[i] = allFile[i].split(".", 1)[0]

for i in range(5):
    print(allFile[i])