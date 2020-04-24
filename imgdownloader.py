# coding: utf8
import xlrd
import urllib.request
import os

def download_img(img_url, imgID, imgdir_path):
    request = urllib.request.Request(img_url)
    try:
        response = urllib.request.urlopen(request)
        # filename = "E:\\商品图片\\" + imgID + ".jpg"
        filename = imgdir_path + imgID + ".jpg"
        if response.getcode() == 200:
            with open(filename, "wb") as f:
                f.write(response.read())  # 将内容写入图片
            return "成功下载: " + filename
    except:
        return "失败，未下载到图片: " + imgID


if __name__ == '__main__':
    ### 参数设置 ###
    # xlsx_path = 'E:\\商品图片excel\\商品库存列表(1-8095).xlsx'
    xlsx_path = "xlsx/to8664.xlsx"
    imgdir_path = "/Users/chen/Desktop/商品图片/"
    sheetIndex = 1

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
    # 读取xlsx下所有ID
    allID = [0] * (totalRows - 1)
    for i in range(totalRows - 1):
        allID[i] = sheet1.cell(i + 1, 2).value
    # 求差集，在xlsx中但不在图片目录中
    newID = list(set(allID).difference(set(allFile)))

    for i in range(len(newID)):
        # 下载新增的图片
        imgID = newID[i]
        img_url = allIDandUrl.get(imgID)
        print(download_img(img_url, imgID, imgdir_path))
        print("剩余: {} 张图片".format(len(newID) - (i+1)))
