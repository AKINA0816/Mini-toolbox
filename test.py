import xlrd

from datetime import date,datetime



file = 'E:\\商品图片excel\\商品库存列表(1-8095).xlsx'

print(file)

wb = xlrd.open_workbook(filename=file)#打开文件
# print(wb.sheet_names())#获取所有表格名字
sheet1 = wb.sheet_by_index(0)#通过索引获取表格

print(sheet1.cell(1,4).value)
print(sheet1.cell(4292,1).value)#获取表格里的内容，三种方式
print(sheet1.nrows)

