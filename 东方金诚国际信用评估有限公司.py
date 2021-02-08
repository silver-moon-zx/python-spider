# -*- coding = utf-8 -*-
# @Time : 2020/9/11 16:33
# @Author : 陈子翔
# @File : 东方金诚国际信用评估有限公司.py
# @software : PyCharm


import urllib.request, urllib.error
import http.cookiejar
import socket
from lxml import etree
import xlwt


def main():
    # 非金融评级
    base_url_1 = "https://www.dfratings.com/index.php?g=Portal&m=Pingj&a=lists&id=45"
    # 定期跟踪
    base_url_2 = "https://www.dfratings.com/index.php?g=Portal&m=Pingj&a=lists&id=50"
    # 不定期跟踪
    base_url_3 = "https://www.dfratings.com/index.php?g=Portal&m=Pingj&a=lists&id=51"

    base_url_list = [base_url_1, base_url_2, base_url_3]

    page_dict = {
        "url_1": 151,
        "url_2": 558,
        "url_3": 27
    }

    # 建立url列表
    url_list = create_URL(base_url_list, page_dict)
    datalist = []
    for urls in url_list:
        datas = []
        for url in urls:
            data = getData(url)
            for dict in data:
                datas.append(dict)
        datalist.append(datas)

    savepath = ".\\东方金诚国际信用评估有限公司。xls"
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

    # 获取li标签下的所有内容
    lis = html.xpath('//div[@class="index-sec1-main-show-con"]/ul/li')
    for li in lis:
        # 1. 项目名称
        name_list = li.xpath('./dl/dd/a/text()')
        if name_list == []:
            name = "None"
        else:
            name = name_list[0]

        # 2. 主体等级
        main_level_list = li.xpath('./dl/dd[2]/text()')
        if main_level_list == []:
            main_level = "None"
        else:
            main_level = main_level_list[0]

        # 3. 债项评级
        debt_level_list = li.xpath('./dl/dd[3]/text()')
        if debt_level_list == []:
            debt_level = "None"
        else:
            debt_level = debt_level_list[0]

        # 4. 评级展望
        rank_level_list = li.xpath('./dl/dd[4]/text()')
        if rank_level_list == []:
            rank_level = "None"
        else:
            rank_level = rank_level_list[0]

        # 5. 公告时间
        date_list = li.xpath('./dl/dd[5]/text()')
        if date_list == []:
            date = "None"
        else:
            date = date_list[0]

        # 6. 公告链接
        link_list = li.xpath('./dl/dd[@class="index-sec-time"]/a/@href')
        if link_list == []:
            link = "None"
        else:
            link = "https://www.dfratings.com/" + link_list[0]


        if name == main_level:
            pass
        else:
            data = {
                "name": name,
                "main_level": main_level,
                "debt_level": debt_level,
                "rank_level": rank_level,
                "date": date,
                "link": link
            }
            datas.append(data)

    return datas


def saveData(datalist, savepath):
    print("save...")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet_1 = book.add_sheet('非金融评级', cell_overwrite_ok=True)
    sheet_2 = book.add_sheet('定期跟踪', cell_overwrite_ok=True)
    sheet_3 = book.add_sheet('不定期跟踪', cell_overwrite_ok=True)


    col = ("项目名称", "主体等级", "债项评级", "评级展望", "公告时间", "公告链接")
    for i in range(0, 6):
        sheet_1.write(0, i, col[i])
        sheet_2.write(0, i, col[i])
        sheet_3.write(0, i, col[i])

    # 1. 非金融评级
    i = 1
    for data_dict in datalist[0]:
        j = 0
        for value in data_dict.values():
            sheet_1.write(i, j, value)
            j = j + 1
        i = i + 1

    # 2. 定期跟踪
    i = 1
    for data_dict in datalist[1]:
        j = 0
        for value in data_dict.values():
            sheet_2.write(i, j, value)
            j = j + 1
        i = i + 1

    # 3. 不定期跟踪
    i = 1
    for data_dict in datalist[2]:
        j = 0
        for value in data_dict.values():
            sheet_3.write(i, j, value)
            j = j + 1
        i = i + 1

    book.save(savepath)

if __name__ == '__main__':
    main()
    print("爬取完毕")