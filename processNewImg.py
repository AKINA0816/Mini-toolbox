# coding: utf8
# 使用条件，本地和vps storage 数据要一致
import xlrd
import urllib.request
import os
from aip import AipImageSearch

def download_img(img_url, imgID, imgdir_path):
    try:
        request = urllib.request.Request(img_url)
        response = urllib.request.urlopen(request)
        # filename = "E:\\商品图片\\" + imgID + ".jpg"
        filePath = imgdir_path + imgID + ".jpg"
        if response.getcode() == 200:
            with open(filePath, "wb") as f:
                f.write(response.read())  # 将内容写入图片
            return filePath
    except:
        return "-1"  # 失败

def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()

def uploadpicture_local(image, name, id, location):
    """有可选参数 """
    options = {}
    options["brief"] = "{\"name\":\"%s\", \"id\":\"%s\", \"location\":\"%s\"}"%(name, id, location)
    return client.productAdd(image, options)


if __name__ == '__main__':
    ### 参数设置 ###
    # xlsx_path = 'E:\\商品图片excel\\商品库存列表(1-8095).xlsx'
    xlsx_path = "xlsx/to8664.xlsx"
    imgdir_path = "/Users/chen/Desktop/商品图片/"
    sheetIndex = 1

    # APPID AK SK
    APP_ID = '16690686'
    API_KEY = 'nrQRHTdphvkcVG9l8V0vkiRt'
    SECRET_KEY = 'GRZMxVfpVWRVGRI0cIG0uygWwxn3yahn'
    client = AipImageSearch(APP_ID, API_KEY, SECRET_KEY)

    ###############
    wb = xlrd.open_workbook(filename=xlsx_path)  # 打开xlsx文件
    sheet1 = wb.sheet_by_index(sheetIndex)  # 通过索引获取表格
    totalRows = sheet1.nrows  # 得到数据的总行数

    # 读取图片目录下所有的文件名
    allFile = os.listdir(imgdir_path)
    for i in range(len(allFile)):  # 删除 ".jpg" 只保留编号
        allFile[i] = allFile[i].split(".", 1)[0]
    # 读取xlsx下所有ID和URL，存为字典
    allIDandUrl = {}
    for i in range(totalRows - 1):
        allIDandUrl[sheet1.cell(i + 1, 2).value] = sheet1.cell(i + 1, 5).value  # ID: URL
    # 读取xlsx下所有ID和Name，存为字典
    allIDandName = {}
    for i in range(totalRows - 1):
        allIDandName[sheet1.cell(i + 1, 2).value] = sheet1.cell(i + 1, 1).value  # ID: Name
    allIDandLocation = {}
    for i in range(totalRows - 1):
        allIDandLocation[sheet1.cell(i + 1, 2).value] = sheet1.cell(i + 1, 7).value  # ID: Location
    # 读取xlsx下所有ID
    allID = [0] * (totalRows - 1)
    for i in range(totalRows - 1):
        allID[i] = sheet1.cell(i + 1, 2).value
    # 求差集，在xlsx中但不在图片目录中
    newID = list(set(allID).difference(set(allFile)))

    for i in range(len(newID)):
        # 下载新增的图片
        imgID = newID[i]
        imgName = allIDandName.get(imgID)
        imgLocation = allIDandLocation.get(imgID)
        img_url = allIDandUrl.get(imgID)
        downloadResult = download_img(img_url, imgID, imgdir_path)
        if downloadResult != "-1":
            print("成功下载: " + imgID)
            getJson = uploadpicture_local(get_file_content(downloadResult), imgName, imgID, imgLocation)
            print("成功上传: " + imgID)
            # sheet1.write
        if downloadResult == "-1":
            print("下载失败: " + imgID)
        print("剩余: {} 张图片".format(len(newID) - (i+1)))
