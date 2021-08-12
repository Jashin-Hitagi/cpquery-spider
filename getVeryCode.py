import json
from aip import AipOcr
import re


def binarizing(img):  # input: gray image
    threshold = 200
    pixdata = img.load()
    w, h = img.size
    for y in range(h):
        for x in range(w):
            if pixdata[x, y] < threshold:
                pixdata[x, y] = 0
            else:
                pixdata[x, y] = 255
    return img


def del_other_dots(img):
    pixdata = img.load()
    w, h = img.size
    for i in range(h):  # 最左列和最右列
        # print(pixdata[0, i]) # 最左边一列的像素点信息
        # print(pixdata[w-1, i]) # 最右边一列的像素点信息
        if pixdata[0, i] == 0 and pixdata[1, i] == 255:
            pixdata[0, i] = 255
        if pixdata[w - 1, i] == 0 and pixdata[w - 2, i] == 255:
            pixdata[w - 1, i] = 255

    for i in range(w):  # 最上行和最下行
        # print(pixdata[i, 0]) # 最上边一行的像素点信息
        # print(pixdata[i, h-1]) # 最下边一行的像素点信息
        if pixdata[i, 0] == 0 and pixdata[i, 1] == 255:
            pixdata[i, 0] = 255
        if pixdata[i, h - 1] == 0 and pixdata[i, h - 2] == 255:
            pixdata[i, h - 1] = 255

    for y in range(1, h - 1):
        for x in range(1, w - 1):
            if pixdata[x, y] == 0:  # 遍历除了四个边界之外的像素黑点
                count = 0  # 统计某个黑色像素点周围九宫格中白块的数量（最多8个）
                if pixdata[x + 1, y + 1] == 255:
                    count = count + 1
                if pixdata[x + 1, y] == 255:
                    count = count + 1
                if pixdata[x + 1, y - 1] == 255:
                    count = count + 1
                if pixdata[x, y + 1] == 255:
                    count = count + 1
                if pixdata[x, y - 1] == 255:
                    count = count + 1
                if pixdata[x - 1, y + 1] == 255:
                    count = count + 1
                if pixdata[x - 1, y] == 255:
                    count = count + 1
                if pixdata[x - 1, y - 1] == 255:
                    count = count + 1

                if count > 4:
                    # print('位置：(' + str(x) + ', ' + str(y) + ')----' + str(count))
                    pixdata[x, y] = 255

    for i in range(h):  # 最左列和最右列
        if pixdata[0, i] == 0 and pixdata[1, i] == 255:
            pixdata[0, i] = 255
        if pixdata[w - 1, i] == 0 and pixdata[w - 2, i] == 255:
            pixdata[w - 1, i] = 255

    for i in range(w):  # 最上行和最下行
        if pixdata[i, 0] == 0 and pixdata[i, 1] == 255:
            pixdata[i, 0] = 255
        if pixdata[i, h - 1] == 0 and pixdata[i, h - 2] == 255:
            pixdata[i, h - 1] = 255
    return img


# 对文字识别后的验证进行格式校验和计算
def getCode(data):
    global verycode
    pattern = '\d[\+\-]\d=?'
    if re.match(pattern, data):
        strs = list(data)
        if strs[1] == '+':
            verycode = int(strs[0]) + int(strs[2])
        if strs[1] == '-':
            verycode = int(strs[0]) - int(strs[2])
        print("验证码计算结果为:", verycode)
        return verycode
    else:
        print("验证码校验格式不匹配，正在重试获取验证码")
        return 404


def getRealCode(image):
    # 百度api
    APP_ID = '23849653'
    API_KEY = '3AH9H5ejMnhFTX1sM4bk9P03'
    SECRET_KEY = 'vytkw76cbWDnzOUvOFWdqVD8dLXGGr66'
    ocr = AipOcr(APP_ID, API_KEY, SECRET_KEY)

    image = binarizing(image)  # 二值化
    image = del_other_dots(image)  # 降噪
    image.save("./file/image.png")  # 图片保存
    # 二进制方式打开图片文件
    f = open(r'./file/image.png', 'rb')
    img = f.read()
    # res = ocr.basicGeneral(img)  # 标准精度文字识别
    res = ocr.basicAccurate(img)   # 高精度文字识别
    # 对结果进行遍历
    if res.get('words_result'):
        result = res.get('words_result').__getitem__(0)
        json_str = json.dumps(result, sort_keys=True)
        params_json = json.loads(json_str)
        items = params_json.items()
        for key, value in items:
            print('文字识别结果：', str(value))
            result = getCode(str(value))  # 验证结果是否符合规则并进行计算
        return result
    else:
        print('文字识别失败，正在重试获取验证码')
        return 404
