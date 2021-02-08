# -*- coding = utf-8 -*-
# @Time : 2020/9/11 18:17
# @Author : 陈子翔
# @File : 联合资信评估有限公司.py
# @software : PyCharm

import urllib.request, urllib.error
import http.cookiejar
import socket
from lxml import etree
import xlwt


def main():
    base_url = "http://www.lhratings.com/announcement/index.html?type=1&page="
    page = 325
    datas = getData(base_url, page)
    savepath = ".\\联合资信评估有限公司.xls"
    saveData(datas, savepath)


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


def getData(base_url, page):
    datas = []
    for i in range(0, page):
        url = base_url + str(i+1)
        html_str = askURL(url)
        html = etree.HTML(html_str)
        # 获取tr标签下的所有内容
        trs = html.xpath('//table[@class="list-table"]/tbody/tr')
        for tr in trs:
            # 项目名称
            name_list = tr.xpath('./td[2]/a/text()')
            if name_list == []:
                name = "None"
            else:
                name = name_list[0]

            # 主体级别
            main_level_list = tr.xpath('./td[3]/text()')
            if main_level_list == []:
                main_level = "None"
            else:
                main_level = main_level_list[0]

            # 展望
            rank_level_list = tr.xpath('./td[4]/text()')
            if rank_level_list == []:
                rank_level = "None"
            else:
                rank_level = rank_level_list[0]

            # 债项级别
            debt_level_list = tr.xpath('./td[5]/text()')
            if debt_level_list == []:
                debt_level = "None"
            else:
                debt_level = debt_level_list[0]

            # 更新时间
            date_list = tr.xpath('./td[6]/text()')
            if date_list == []:
                date = "None"
            else:
                date = date_list[0]

            # 报告链接
            link_list = tr.xpath('./td[7]/a/@href')
            if link_list == []:
                link = "None"
            else:
                link = link_list[0]


            if name == main_level:
                pass
            else:
                data = {
                    "name": name,
                    "main_level": main_level,
                    "rank_level": rank_level,
                    "debt_level": debt_level,
                    "date": date,
                    "link": link
                }
                datas.append(data)
    return datas


def saveData(datas, savepath):
    print("save...")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.add_sheet('非金融企业', cell_overwrite_ok=True)
    col = ("项目名称", "主体级别", "展望", "债项级别", "更新时间", "报告链接")
    for i in range(0, 6):
        sheet.write(0, i, col[i])
    j = 1
    for data_dict in datas:
        k = 0
        for data in data_dict.values():
            sheet.write(j, k, data)
            k = k + 1
        j = j + 1

    book.save(savepath)

if __name__ == '__main__':
    main()
    print("爬取完毕")
