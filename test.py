import xlrd

from datetime import date,datetime

file = 'C:\\Users\\28015\\Desktop\\商品库存列表导出(1-500).xlsx'

def read_excel():

	wb = xlrd.open_workbook(filename=file)#打开文件

	# print(wb.sheet_names())#获取所有表格名字

	sheet1 = wb.sheet_by_index(0)#通过索引获取表格

	print(sheet1.cell(1,0).value)#获取表格里的内容，三种方式



read_excel()
