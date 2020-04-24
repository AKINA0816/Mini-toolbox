import xlrd
import urllib.request

from datetime import date, datetime

file = 'xlsx/to8664.xlsx'

print(file)

wb = xlrd.open_workbook(filename=file)  # 打开文件
# print(wb.sheet_names())#获取所有表格名字
sheet1 = wb.sheet_by_index(1)  # 通过索引获取表格

print(sheet1.cell(7, 5).value)
print(sheet1.cell(7, 5).value == "")

row = 8  # 想要获取第八行的商品名称
col = 1  # B列
print(sheet1.cell(row - 1, col).value)  # 获取表格里的内容
print(sheet1.nrows)

import os
allFile = os.listdir("/Users/chen/Desktop/商品图片")
for i in range(len(allFile)):
    allFile[i] = allFile[i].split(".", 1)[0]

mylist = [[0] * 2 for i in range(sheet1.nrows - 1)]
for i in range(sheet1.nrows - 1):
    mylist[i][0] = sheet1.cell(i + 1, 2).value
    mylist[i][1] = sheet1.cell(i + 1, 5).value

id = mylist[len(mylist)-1][0]
url =mylist[len(mylist)-1][1]


allIDandUrl = {}
for i in range(sheet1.nrows - 1):
    allIDandUrl[sheet1.cell(i + 1, 2).value] = sheet1.cell(i + 1, 5).value  # ID: URL

ans = allIDandUrl.get("ED831164KC1")
print(ans)
#求差集，在B中但不在A中
# retD = list(set(mylist).difference(set(allFile)))
# print(len(allFile))
# print(len(retD))





