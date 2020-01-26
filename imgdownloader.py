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
    file = 'C:\\Users\\28015\\Desktop\\商品库存列表导出(1-500).xlsx'
    wb = xlrd.open_workbook(filename=file)  # 打开文件
    sheet1 = wb.sheet_by_index(0)  # 通过索引获取表格
    for i in range(266):
        # 下载要的图片
        imgname = sheet1.cell(i + 1,0).value
        img_url = sheet1.cell(i + 1,1).value
        print(download_img(img_url, imgname))


