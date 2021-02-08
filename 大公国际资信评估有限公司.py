# -*- coding = utf-8 -*-
# @Time : 2020/9/10 11:00
# @Author : 陈子翔
# @File : 大公国际资信评估有限公司.py
# @software : PyCharm


import urllib.request, urllib.error
import http.cookiejar
import socket
from lxml import etree
import xlwt



def main():
    # 短期融资券
    base_url_1 = "http://www.dagongcredit.com/index.php?m=content&c=index&a=lists&catid=79"
    # 中期票据
    base_url_2 = "http://www.dagongcredit.com/index.php?m=content&c=index&a=lists&catid=80"
    # 企业债券
    base_url_3 = "http://www.dagongcredit.com/index.php?m=content&c=index&a=lists&catid=81"
    # 公司债券
    base_url_4 = "http://www.dagongcredit.com/index.php?m=content&c=index&a=lists&catid=82"
    # 跟踪评级
    base_url_5 = "http://www.dagongcredit.com/index.php?m=content&c=index&a=lists&catid=158"

    base_url_list = [base_url_1, base_url_2, base_url_3, base_url_4, base_url_5]
    # 各类公告的页数
    page_dict = {
        "url_1": 60,
        "url_2": 45,
        "url_3": 3,
        "url_4": 22,
        "url_5": 222
    }

    # 建立url列表
    url_list = create_URL(base_url_list, page_dict)
    # print(url_list)
    datalist = []
    for urls in url_list:
        datas = []
        for url in urls:
            data = getData(url)
            for dict in data:
                datas.append(dict)
        datalist.append(datas)

    savepath = ".\\大公国际资信评估报告.xls"
    saveData(datalist, savepath)


def create_URL(base_url_list, page_dict):
    url_list = []
    i = 0
    for base_url in base_url_list:
        page = page_dict["url_%d" % (i+1)]
        urls = []
        i = i + 1
        for j in range(0, page):
            url = base_url + "&page=%d" % (j + 1)
            urls.append(url)
        url_list.append(urls)

    return url_list


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
        response = opener.open(req, timeout=500)

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


def getData(url):
    datas = []
    html_str = askURL(url)
    # print(html_str)
    html = etree.HTML(html_str)

    # 获取span标签下的所有信息
    spans = html.xpath('//ul/li')
    for span in spans:
        name_list = span.xpath('./span[not(@class="active")]/a/text()')
        if len(name_list) == 0:
            name = "None"
        else:
            name = name_list[0]
        level_list = span.xpath('./span[not(@class="active")][1]/text()')
        if len(level_list) == 0:
            level = "None"
        else:
            level = level_list[0]
        date_list = span.xpath('./span[not(@class="active")][3]/text()')
        if len(date_list) == 0:
            date = "None"
        else:
            date = date_list[0]
        link_list = span.xpath('./span[not(@class="active")]/a/@href')
        if len(link_list) == 0:
            link = "None"
        else:
            link = link_list[0]


        if name == level:
            pass
        else:
            data = {
                "name": name,
                "level": level,
                "date": date,
                "link": link
            }

            datas.append(data)


    return datas


def saveData(datalist, savepath):
    print("save...")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet_1 = book.add_sheet('短期融资券', cell_overwrite_ok=True)
    sheet_2 = book.add_sheet('中期票据', cell_overwrite_ok=True)
    sheet_3 = book.add_sheet('企业债券', cell_overwrite_ok=True)
    sheet_4 = book.add_sheet('公司债券', cell_overwrite_ok=True)
    sheet_5 = book.add_sheet('跟踪评级', cell_overwrite_ok=True)

    col = ("项目名称", "主体级别", "评级时间", "公告链接")
    for i in range(0, 4):
        sheet_1.write(0, i, col[i])
        sheet_2.write(0, i, col[i])
        sheet_3.write(0, i, col[i])
        sheet_4.write(0, i, col[i])
        sheet_5.write(0, i, col[i])

    # 短期融资券
    i = 1
    for data_dict in datalist[0]:
        j = 0
        for value in data_dict.values():
            sheet_1.write(i, j, value)
            j = j + 1
        i = i + 1
    # 中期票据
    i = 1
    for data_dict in datalist[1]:
        j = 0
        for value in data_dict.values():
            sheet_2.write(i, j, value)
            j = j + 1
        i = i + 1
    # 企业债券
    i = 1
    for data_dict in datalist[2]:
        j = 0
        for value in data_dict.values():
            sheet_3.write(i, j, value)
            j = j + 1
        i = i + 1
    # 公司债券
    i = 1
    for data_dict in datalist[3]:
        j = 0
        for value in data_dict.values():
            sheet_4.write(i, j, value)
            j = j + 1
        i = i + 1
    # 跟踪评级
    i = 1
    for data_dict in datalist[4]:
        j = 0
        for value in data_dict.values():
            sheet_5.write(i, j, value)
            j = j + 1
        i = i + 1

    book.save(savepath)



if __name__ == '__main__':
    main()
    print("爬取完毕")
