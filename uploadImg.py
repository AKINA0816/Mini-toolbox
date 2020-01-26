# 上传图片到百度图片搜索图库里
import xlrd
import xlwt
from xlwt import easyxf
from xlutils.copy import copy
from aip import AipImageSearch

""" 你的 APPID AK SK """
APP_ID = '16690686'
API_KEY = 'nrQRHTdphvkcVG9l8V0vkiRt'
SECRET_KEY = 'GRZMxVfpVWRVGRI0cIG0uygWwxn3yahn'

client = AipImageSearch(APP_ID, API_KEY, SECRET_KEY)

""" 读取图片 """
def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()

""" 打开excel """
file = 'E:\\商品图片excel\\商品库存列表(1-8095).xlsx'
# file = 'C:\\Users\\28015\\Desktop\\test.xlsx'
rb = xlrd.open_workbook(filename=file)  # 打开文件
sheet1 = rb.sheet_by_index(0)#通过索引获取表格

wb = copy(rb)
w_sheet1 = wb.get_sheet(0)

""" 带参数调用商品图片搜索—入库, 图片参数为本地图片 """
def uploadpicture_local(image, name, id, location):
    """有可选参数 """
    options = {}
    options["brief"] = "{\"name\":\"%s\", \"id\":\"%s\", \"location\":\"%s\"}"%(name, id, location)
    return client.productAdd(image, options)

totalRows = sheet1.nrows - 1
for i in range (totalRows):
    try:
        name = sheet1.cell(i + 1, 0).value
        id = sheet1.cell(i + 1, 1).value
        imgAddress = get_file_content("E:\商品图片\{}.jpg".format(id))
        location = sheet1.cell(i + 1, 6).value
        getJson = uploadpicture_local(imgAddress, name, id, location)
        w_sheet1.write(i + 1, 18, getJson['cont_sign'])
        wb.save('E:\\商品图片excel\\商品库存列表(1-8095).xlsx')
        print("剩余：{} 个".format(totalRows - i))
    except IOError:
        print("他妈的找不到 {} 这个图片".format(id))
        pass
    continue

