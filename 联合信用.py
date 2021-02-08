import urllib.request, urllib.error
import http.cookiejar
import socket
from bs4 import BeautifulSoup  # 网页解析，获取数据
import re   # 正则表达式，进行文字匹配
import xlwt  # 进行excel操作




def main():
    baseurl = "http://www.lianhecreditrating.com.cn/List.aspx?m=20140627094836327647&Page="
# 更改数据
    num_1 = 4742
    page_1 = 238
    # 主程序
    # print(askURL(url))
    datalist, datalist_1, datalist_2 = getData(baseurl, page_1)
    savepath =".\\联合信用评级公告.xls"
    saveData(datalist, datalist_1, datalist_2, savepath, num_1)

# 正则提取
findPDF = re.compile(r'<a href="(.+?)" target="_blank">')

findCompany = re.compile(r'target="_blank">(.*?)</a>', re.DOTALL)

findLevel = re.compile(r'<td style="text-align: center;" width="440">(.*?)</td>', re.DOTALL)

findDate = re.compile(r'<td style="font-family: Verdana; text-align: center;" width="110">(.+?)</td>', re.DOTALL)

def askURL(url):
    # 获取网页源码
    # 获取网页cookie，构造opener
    filename = 'wk_cookie.txt'
    cookie = http.cookiejar.LWPCookieJar(filename)
    handler = urllib.request.HTTPCookieProcessor(cookie)
    opener = urllib.request.build_opener(handler)
    # 获取网页源码
    headers = {
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 84.0.4147.105Safari / 537.36"
    }
    req = urllib.request.Request(url, headers=headers)
    try:
        response = opener.open(req, timeout=5)

        # 中文字符的Unicode编码bai0x0800-0xFFFF之间,(utf-8包含了部分汉字)当你试图将du该“中文字符”转成U码的utf-8时超出了其范筹而GBK 规范收录了 ISO 10646.1 中的全部 CJK 汉字和符号，并有所补充，所以解决方法是将.decode('utf-8')改为.decode('gbk')
        html = response.read().decode('utf-8')

    except urllib.error.HTTPError as e:
        print(e.reason)
    except urllib.error.URLError as e:
        if isinstance(e.reason, socket.timeout):
            print("Time Out")
        else:
            print(e.reason)
    cookie.save(ignore_discard=True, ignore_expires=True)
    return html


def getData(baseurl, page_1):
    datalist = []
    datalist_1 = []
    datalist_2 = []
    for i in range(0, page_1):
        url = baseurl + str(i+1)
        html =askURL(url)

        # 逐一解析数据
        soup = BeautifulSoup(html, "html.parser")
        for item in soup.find_all(name='td'):
            # print(item)
            data = []
            data_1 = []
            data_2 = []
            item = str(item)

            # 获取公司名称
            name = re.findall(findCompany, item)
            # print(name)
            if len(name) == 1:
                name = name[0].strip()
                # print(type(name))
                data.append(name)

            # 获取pdf链接
            pdf = re.findall(findPDF, item)
            if len(pdf) == 1:
                pdf = "http://www.lianhecreditrating.com.cn/"+pdf[0]
                # print(type(pdf))
                data.append(pdf)



            if data != []:
                # print(data)
                datalist.append(data)


            # 获取公司等级/文号/类型
            level = re.findall(findLevel, item)
            # print(level)
            if len(level) == 1 and level[0] != "\r\n                                    等级/文号/类型\r\n                                ":
                if level[0] == '\n':
                    level = level[0].replace('\n', "无")
                else:
                    level = level[0].strip()
                # print(type(level))
                data_1.append(level)


            if data_1 != []:
                datalist_1.append(data_1)

            # 获取日期
            date = re.findall(findDate, item)
            if len(date) == 1:
                date = date[0].strip()
                # print(type(date))
                data_2.append(date)
            if data_2 != []:
                datalist_2.append(data_2)

    # print(datalist_1)
    # print(datalist_2)

    return datalist, datalist_1, datalist_2



def saveData(datalist, datalist_1, datalist_2, savepath, num_1):
    print("save...")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.add_sheet('联合信用评级公告', cell_overwrite_ok=True)
    col = ("名称", "评级报告链接", "等级/文号/类型", "日期")
    for i in range(0, 4):
        sheet.write(0, i, col[i])
    for i in range(0, num_1):
        print("第%d条" % (i+1))
        data = datalist[i]
        for j in range(0, 2):
            sheet.write(i+1, j, data[j])    # 数据
    for i in range(0, num_1):
        data = datalist_1[i]
        sheet.write(i+1, 2, data)
    for i in range(0, num_1):
        data = datalist_2[i]
        sheet.write(i+1, 3, data)

    book.save(savepath)


if __name__ == "__main__":
    main()
    print("爬取完毕")