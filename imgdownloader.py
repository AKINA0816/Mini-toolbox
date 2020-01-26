# coding: utf8
import xlrd
import urllib.request

def download_img(img_url, imgname):
    request = urllib.request.Request(img_url)
    try:
        response = urllib.request.urlopen(request)
        filename = "E:\\商品图片\\"+ imgname + ".jpg"
        if (response.getcode() == 200):
            with open(filename, "wb") as f:
                f.write(response.read()) # 将内容写入图片
            return filename
    except:
        return "failed"

if __name__ == '__main__':
    file = 'E:\\商品图片excel\\商品库存列表(1-8095).xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    rows = sheet1.nrows - 1# 得到数据的总行数
    for i in range(rows):
        # 下载要的图片
        imgname = sheet1.cell(i + 735, 1).value
        img_url = sheet1.cell(i + 735, 4).value
        if(img_url == ""):
            print(imgname + '他妈的没有链接')
            continue
        print(download_img(img_url,imgname))
        print("剩余: {} 张图片".format(rows-735 - i))


