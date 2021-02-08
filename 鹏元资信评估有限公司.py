# -*- coding = utf-8 -*-
# @Time : 2020/9/8 15:09
# @Author : 陈子翔
# @File : 鹏元资信评估有限公司.py
# @software : PyCharm

import urllib.request,urllib.error
import socket
import http.cookiejar
from lxml import etree
import xlwt

def main():
    # 企业债券(32)
    baseurl_1 = "http://www.pyrating.cn/statement/qiyezhaiquanpingji"
    url_1 = Create_url(baseurl_1, 32)
    # print(url_1)

    # 公司债券(10)
    baseurl_2 = "http://www.pyrating.cn/zh-cn/statement/gongsizhaiquanpingji"
    url_2 = Create_url(baseurl_2, 10)

    # 集合债券(1)
    baseurl_3 = "http://www.pyrating.cn/zh-cn/statement/jihezhaiquanpingji"
    url_3 = Create_url(baseurl_3, 1)

    # 主体评级(1)
    baseurl_4 = "http://www.pyrating.cn/zh-cn/statement/zhutipingji"
    url_4 = Create_url(baseurl_4, 1)


    # 地方政府债券(4)
    baseurl_5 = "http://www.pyrating.cn/zh-cn/statement/difangzhengfuzhaiquan"
    url_5 = Create_url(baseurl_5, 4)


    # 短期融资券(1)
    baseurl_6 = "http://www.pyrating.cn/zh-cn/statement/duanqirongziquan"
    url_6 = Create_url(baseurl_6, 1)


    # 中期票据(1)
    baseurl_7 = "http://www.pyrating.cn/zh-cn/statement/zhongqipiaoju"
    url_7 = Create_url(baseurl_7, 1)

    # 跟踪评级

    # 定期跟踪评级(181)
    baseurl_8 = "http://www.pyrating.cn/statement/track"
    url_8 = Create_url(baseurl_8, 181)
    # 延迟披露公告(17)
    baseurl_8_2 = "http://www.pyrating.cn/zh-cn/statement/yanshipilugonggao"
    url_8_2 = Create_url(baseurl_8_2, 17)

    # 不定期跟踪评级
    # 级别调整公告(5)
    baseurl_9 = "http://www.pyrating.cn/zh-cn/statement/genzongpingji/budingqigenzongpingji/jibiediaozhenggonggao"
    url_9 = Create_url(baseurl_9, 5)

    # 评级观察名单(3)
    baseurl_10 = "http://www.pyrating.cn/zh-cn/statement/genzongpingji/budingqigenzongpingji/pingjiguanchamingdan"
    url_10 = Create_url(baseurl_10, 3)

    # 专项评级公告(2)
    baseurl_11 = "http://www.pyrating.cn/zh-cn/statement/genzongpingji/budingqigenzongpingji/zhuanxiangpingjigonggao"
    url_11 = Create_url(baseurl_11, 2)

    # 重点关注公告(15)
    baseurl_12 = "http://www.pyrating.cn/zh-cn/statement/genzongpingji/budingqigenzongpingji/zhongdianguanzhugonggao"
    url_12 = Create_url(baseurl_12, 15)

    # 不定期跟踪公告(1)
    baseurl_13 = "http://www.pyrating.cn/zh-cn/statement/genzongpingji/budingqigenzongpingji/statement/notrack"
    url_13 = Create_url(baseurl_13, 1)

    # 终止评级(8)
    baseurl_14 = "http://www.pyrating.cn/zh-cn/statement/genzongpingji/zhongzhipingji"
    url_14 = Create_url(baseurl_14, 8)

    # 公司治理评级(1)
    baseurl_15 = "http://www.pyrating.cn/zh-cn/statement/gongsizhilipingji"
    url_15 = Create_url(baseurl_15, 1)
    page_1 = 1
    # 发行人一览(46)
    baseurl_16 = "http://www.pyrating.cn/zh-cn/statement/faxingrenyilan"
    url_16 = Create_url(baseurl_16, 46)
    page_2 = 46

    urllist = [url_1, url_2, url_3, url_4, url_5, url_6, url_7, url_8, url_8_2, url_9, url_10, url_11, url_12, url_13, url_14]
    infos = getData(urllist)
    savepath = ".\\鹏元资信评估有限公司.xls"
    saveData(infos, savepath)


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
        response = opener.open(req, timeout=100)

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


def getData(urllist):
    infos = []
    for urls in urllist:
        for url in urls:
            html_text = askURL(url)
            # print(type(html_text))
            html = etree.HTML(html_text)
            # 需要提取的信息：债券类型， 公告名称， 发布日期
            # 先提取tr标签下的所有信息
            trs = html.xpath('//tbody/tr[position()>1][not(@class="empty")]')
            for tr in trs:
                info = {}
                # 1. 债券类型
                bond_list = tr.xpath('./td[1]/text()')
                if len(bond_list) == 0:
                    bond = "None"
                else:
                    bond = bond_list[0]
                # 2. 公告名称
                notice_list = tr.xpath('./td[@class="text_left"]/a/text()')
                if len(notice_list) == 0:
                    notice = "None"
                else:
                    notice = notice_list[0]
                # 3. 评级公告链接
                link_list = tr.xpath('./td[@class="text_left"]/a/@href')
                if len(link_list) == 0:
                    link = "None"
                else:
                    link = "http://www.pyrating.cn" + link_list[0]

                # 4. 日期
                date_list = tr.xpath('./td[3]/text()')
                if len(date_list) == 0:
                    date = "None"
                else:
                    date = date_list[0]

                info = {
                    "bond": bond,
                    "notice": notice,
                    "link": link,
                    "date": date,
                }
                infos.append(info)
    # print(infos)
    return infos


def saveData(infos, savepath):
    num = len(infos)
    print("save...")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.add_sheet('联合信用评级公告', cell_overwrite_ok=True)
    col = ("债券类型", "评级公告名称", "评级公告链接", "发布日期")
    for i in range(0, 4):
        sheet.write(0, i, col[i])
    for i in range(0, num):
        print("第%d条" % (i+1))
        j = 0
        for value in infos[i].values():
            sheet.write(i+1, j, value)
            j = j + 1
    book.save(savepath)


def Create_url(baseurl, page):
    url = []
    for i in range(0, page):
        url_page = "?page=%d" % (i+1)
        url_data = baseurl + url_page
        url.append(url_data)
    return url

if __name__ == "__main__":
    main()
    print("爬取完毕")