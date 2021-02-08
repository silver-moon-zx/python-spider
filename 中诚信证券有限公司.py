# -*- coding = utf-8 -*-
# @Time : 2020/8/12 10:04
# @Author : 陈子翔
# @File : 中诚信证券有限公司.py
# @software : PyCharm


import urllib.request, urllib.error, urllib.parse
import http.cookiejar
import re   # 正则表达式，进行文字匹配
import xlwt  # 进行excel操作



def main():
    baseurl = "http://www.ccxr.com.cn/info_list/selInfo_list.do"
    '''
    初始评级，普通公司债：htmllist_1——37个元素
    初始评级，可交换公司债：htmllist_2——1个元素
    初始评级，可转换公司债：htmllist_3——3个元素
    定期跟踪：htmllist_4——65个元素
    不定期跟踪：htmllist_5——5个元素
    其他：htmllist_6——1个元素
    '''
    data, data_2, data_3, data_4, data_5, data_6 = getData(baseurl)
    savepath =".\\中诚信证券评级公告.xls"
    saveData(data, data_2, data_3, data_4, data_5, data_6, savepath)



def askURL(url):
    # 获取网页源码
    # 获取网页cookie，构造opener
    filename = 'wk_cookie.txt'
    cookie = http.cookiejar.LWPCookieJar(filename)
    handler = urllib.request.HTTPCookieProcessor(cookie)
    opener = urllib.request.build_opener(handler)
    # 模拟Request Headers
    headers = {
        "Host": "www.ccxr.com.cn",
        "Origin": "http://www.ccxr.com.cn",
        "Referer": "http://www.ccxr.com.cn/notice.html?sm_id=0&su_id=35",
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 84.0.4147.105Safari / 537.36",
        "X-Requested-With": "XMLHttpRequest"
    }

    htmllist_1 = []
    htmllist_2 = []
    htmllist_3 = []
    htmllist_4 = []
    htmllist_5 = []
    htmllist_6 = []


# 初始评级，普通公司债
    for i in range(0, 37):
        dict_1 = {
            'pageNow': i+1,  # 1~37
            'sm_value': 30,
            'su_value': ''
        }

        data_1 = bytes(urllib.parse.urlencode(dict_1), encoding='utf-8')
        req_1 = urllib.request.Request(url, data=data_1, headers=headers, method="POST")
        response_1 = opener.open(req_1, timeout=5)
        html_1 = response_1.read().decode('utf-8')
        htmllist_1.append(html_1)


# 初始评级，可交换公司债
    dict_2 = {
        'pageNow': 1,
        'sm_value': 31,
        'su_value': ''
    }

    data_2 = bytes(urllib.parse.urlencode(dict_2), encoding='utf-8')
    req_2 = urllib.request.Request(url, data=data_2, headers=headers, method="POST")
    response_2 = opener.open(req_2, timeout=5)
    html_2 = response_2.read().decode('utf-8')
    htmllist_2.append(html_2)


# 初始评级，可转换公司债
    for i in range(0, 3):
        dict_3 = {
            'pageNow': i+1,  # 1~3
            'sm_value': 32,
            'su_value': ''
        }
        data_3 = bytes(urllib.parse.urlencode(dict_3), encoding='utf-8')
        req_3 = urllib.request.Request(url, data=data_3, headers=headers, method="POST")
        response_3 = opener.open(req_3, timeout=5)
        html_3 = response_3.read().decode('utf-8')
        htmllist_3.append(html_3)


# 定期跟踪
    for i in range(0, 65):
        dict_4 = {
            'pageNow': i+1,  # 1~65
            'sm_value': '',
            'su_value': 33
        }
        data_4 = bytes(urllib.parse.urlencode(dict_4), encoding='utf-8')
        req_4 = urllib.request.Request(url, data=data_4, headers=headers, method="POST")
        response_4 = opener.open(req_4, timeout=5)
        html_4 = response_4.read().decode('utf-8')
        htmllist_4.append(html_4)


# 不定期跟踪
    for i in range(0, 5):
        dict_5 = {
            'pageNow': i+1,  # 1~5
            'sm_value': '',
            'su_value': 34
        }
        data_5 = bytes(urllib.parse.urlencode(dict_5), encoding='utf-8')
        req_5 = urllib.request.Request(url, data=data_5, headers=headers, method="POST")
        response_5 = opener.open(req_5, timeout=5)
        html_5 = response_5.read().decode('utf-8')
        htmllist_5.append(html_5)


# 其他
    dict_6 = {
        'pageNow': 1,
        'sm_value': '',
        'su_value': 35
    }
    data_6 = bytes(urllib.parse.urlencode(dict_6), encoding='utf-8')
    req_6 = urllib.request.Request(url, data=data_6, headers=headers, method="POST")
    response_6 = opener.open(req_6, timeout=5)
    html_6 = response_6.read().decode('utf-8')
    htmllist_6.append(html_6)

    return htmllist_1, htmllist_2, htmllist_3, htmllist_4, htmllist_5, htmllist_6


#   正则提取
findName = re.compile(r'"in_title":"(.*?)","su_name"')  # 公司名称
findPDF = re.compile(r'"in_picture":"(.*?)"')   # pdf链接
findHope = re.compile(r'"info6":"(.*?)"', re.DOTALL)   # 展望
findDate = re.compile(r'"addtime":"(.*?)"', re.DOTALL)     # 日期
findLevel = re.compile(r'"info1":"(.*?)"', re.DOTALL)      # 等级
findDebt = re.compile(r'"info2":"(.*?)"', re.DOTALL)       # 债项


def getData(baseurl):
    data = []
    data_2 = []
    data_3 = []
    data_4 = []
    data_5 = []
    data_6 = []
    htmllist_1, htmllist_2, htmllist_3, htmllist_4, htmllist_5, htmllist_6 = askURL(baseurl)

# 初始评级，普通公司债
    for i in range(0, 37):
        html_1 = htmllist_1[i]
        name = re.findall(findName, html_1)
        pdf = re.findall(findPDF, html_1)
        level = re.findall(findLevel, html_1)
        debt = re.findall(findDebt, html_1)
        hope = re.findall(findHope, html_1)
        date = re.findall(findDate, html_1)

        # print("第%d项列表" % (i+1))
        # print(level)      # level有可能会改变，若改变，需要更改下面的语句
        # print(debt)
        # print(hope)
        # print(date)

        for j in range(0, len(name)):
            data.append(name[j])
            pdf[j] = "http://www.ccxr.com.cn/pdf/" + pdf[j]
            data.append(pdf[j])

            # 调整level
            if len(level) != len(name):
                if level == []:
                    for a in range(0, len(name)):
                        level.append('undefined')
                elif level == ['']:
                    level[0] = 'undefined'
                    for a in range(0, 39):
                        level.append('undefined')
                elif level == ['AAA']:
                    for a in range(0, len(name)-9):
                        level.insert(0, 'undefined')
                    for b in range(0, 8):
                        level.append('undefined')
                elif level == ['', '', '', '', '', '', '', '', '', 'AA']:
                    for a in range(0, 9):
                        level[a] = 'undefined'
                    for b in range(0, 12):
                        level.insert(0, 'undefined')
                    for c in range(0, len(name)-22):
                        level.append('undefined')
            else:
                for a in range(0, len(name)):
                    if level[a] == '':
                        level[a] = 'undefined'
            data.append(level[j])

        # print("第%d项列表" % (i+1))
        # print(level)

            # 调整debt
            if len(debt) != len(name):
                if debt == []:
                    for a in range(0, len(name)):
                        debt.append('undefined')
                elif debt == ['']:
                    debt[0] = 'undefined'
                    for a in range(0, 39):
                        debt.append('undefined')
                elif debt == ['AAA']:
                    for a in range(0, len(name) - 9):
                        debt.insert(0, 'undefined')
                    for b in range(0, 8):
                        debt.append('undefined')
                elif debt == ['', '', '', '', '', '', '', '', '', 'AA']:
                    for a in range(0, 9):
                        debt[a] = 'undefined'
                    for b in range(0, 12):
                        debt.insert(0, 'undefined')
                    for c in range(0, len(name) - 22):
                        debt.append('undefined')
            else:
                for a in range(0, len(name)):
                    if debt[a] == '':
                        debt[a] = 'undefined'
            data.append(debt[j])

            # 调整hope
            if len(hope) != len(name):
                if hope == []:
                    for a in range(0, len(name)):
                        hope.append('undefined')
                elif hope == ['']:
                    hope[0] = 'undefined'
                    for a in range(0, 39):
                        hope.append('undefined')
                elif hope == ['稳定']:
                    for a in range(0, len(name) - 9):
                        hope.insert(0, 'undefined')
                    for b in range(0, 8):
                        hope.append('undefined')
                elif hope == ['', '', '', '', '', '', '', '', '', '稳定']:
                    for a in range(0, 9):
                        hope[a] = 'undefined'
                    for b in range(0, 12):
                        hope.insert(0, 'undefined')
                    for c in range(0, len(name) - 22):
                        hope.append('undefined')
            else:
                for a in range(0, len(name)):
                    if hope[a] == '':
                        hope[a] = 'undefined'
            data.append(hope[j])
        # print("第%d项列表" % (i+1))
        # print(hope)
            data.append(date[j])


# 初始评级，可交换公司债
    html_2 = htmllist_2[0]
    name_2 = re.findall(findName, html_2)
    pdf_2 = re.findall(findPDF, html_2)
    level_2 = re.findall(findLevel, html_2)
    debt_2 = re.findall(findDebt, html_2)
    hope_2 = re.findall(findHope, html_2)
    date_2 = re.findall(findDate, html_2)

    # print(level_2)
    for j in range(0, len(name_2)):
        data_2.append(name_2[j])
        pdf_2[j] = "http://www.ccxr.com.cn/pdf/" + pdf_2[j]
        data_2.append(pdf_2[j])
        data_2.append(level_2[j])
        data_2.append(debt_2[j])
        data_2.append(hope_2[j])
        data_2.append(date_2[j])


# 初始评级，可转换公司债
    for i in range(0, 3):
        html_3 = htmllist_3[i]
        name = re.findall(findName, html_3)
        pdf = re.findall(findPDF, html_3)
        level = re.findall(findLevel, html_3)
        debt = re.findall(findDebt, html_3)
        hope = re.findall(findHope, html_3)
        date = re.findall(findDate, html_3)

        # print(level)
        for j in range(0, len(name)):
            data_3.append(name[j])
            pdf[j] = "http://www.ccxr.com.cn/pdf/" + pdf[j]
            data_3.append(pdf[j])

            if len(level) == 0:
                for a in range(0, len(name)):
                    level.append('undefined')
            elif len(level) != len(name):
                while len(level) < len(name):
                    level.append('undefined')
            data_3.append(level[j])

            if len(debt) == 0:
                for a in range(0, len(name)):
                    debt.append('undefined')
            elif len(debt) != len(name):
                while len(debt) < len(name):
                    debt.append('undefined')
            data_3.append(debt[j])

            if len(hope) == 0:
                for a in range(0, len(name)):
                    hope.append('undefined')
            elif len(hope) != 40:
                while len(hope) < len(name):
                    hope.append('undefined')
            data_3.append(hope[j])

            data_3.append(date[j])


# 定期跟踪
    for i in range(0, 65):
        html_4 = htmllist_4[i]
        name = re.findall(findName, html_4)
        pdf = re.findall(findPDF, html_4)
        level = re.findall(findLevel, html_4)
        debt = re.findall(findDebt, html_4)
        hope = re.findall(findHope, html_4)
        date = re.findall(findDate, html_4)

        # print(pdf)
        # print("第%d项" % (i+1))
        # print(len(name))
        # print(len(pdf))
        # print(len(level))
        # print(level)
        # print(len(debt))
        # print(debt)
        # print(len(hope))
        # print(hope)
        # print(len(date))
        # print(date)

        for j in range(0, len(name)):
            data_4.append(name[j])

            if len(pdf) != len(name):
                pdf.insert(1, 'undefined')
                pdf.insert(1, 'undefined')
            pdf[j] = "http://www.ccxr.com.cn/pdf/" + pdf[j]
            data_4.append(pdf[j])

            if len(level) == 24:
                level = ['AAA', 'AAA', 'AAA', 'AAA', 'AAA', 'AAA', 'AA+', 'AA+', 'AA+', 'AA+', 'AA-', 'AA-', 'AA', 'AA', 'AA+', 'AAA', 'AAA']
                while len(level) != len(name):
                    level.append('undefined')
            elif len(level) == 1:
                level[0] = 'undefined'
                while len(level) != len(name):
                    level.append('undefined')
            elif len(level) == 3:
                while len(level) != len(name):
                    level.insert(0, 'undefined')
            elif len(level) == 0:
                while len(level) != len(name):
                    level.append('undefined')
            data_4.append(level[j])

        # print("第%d项" % (i+1))
        # print(len(name))
        # print(len(level))
        # print(level)

            if len(debt) == 24:
                debt = ['AAA', 'AAA', 'AAA', 'AAA', 'AAA', 'AAA', 'AA+', 'AA+', 'AA+', 'AA+', 'AA-', 'AA-', 'AA', 'AA', 'AA+', 'AAA', 'AAA']
                while len(debt) != len(name):
                    debt.append('undefined')
            elif len(debt) == 1:
                debt[0] = 'undefined'
                while len(debt) != len(name):
                    debt.append('undefined')
            elif len(debt) == 3:
                while len(debt) != len(name):
                    debt.insert(0, 'undefined')
            elif len(debt) == 0:
                while len(debt) != len(name):
                    debt.append('undefined')
            data_4.append(debt[j])


            if len(hope) == 24:
                hope = ['稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定', '稳定']
                while len(hope) != len(name):
                    hope.append('undefined')
            elif len(hope) == 1:
                hope[0] = 'undefined'
                while len(hope) != len(name):
                    hope.append('undefined')
            elif len(hope) == 3:
                while len(hope) != len(name):
                    hope.insert(0, 'undefined')
            elif len(hope) == 0:
                while len(hope) != len(name):
                    hope.append('undefined')
            data_4.append(hope[j])

            if len(date) != len(name):
                while len(date) != len(name):
                    date.append('undefined')
            data_4.append(date[j])


# 不定期跟踪
    for i in range(0, 5):
        html_5 = htmllist_5[i]
        name = re.findall(findName, html_5)
        pdf = re.findall(findPDF, html_5)
        level = re.findall(findLevel, html_5)
        debt = re.findall(findDebt, html_5)
        hope = re.findall(findHope, html_5)
        date = re.findall(findDate, html_5)


        # print("第%d项" % (i+1))
        # print(len(name))
        # print(len(level))
        # print(level)
        # print(len(hope))
        # print(hope)

        for j in range(0, len(name)):
            data_5.append(name[j])
            pdf[j] = "http://www.ccxr.com.cn/pdf/" + pdf[j]
            data_5.append(pdf[j])

            if level[j] == '' or level[j] == '——':
                level[j] = 'undefined'
            data_5.append(level[j])
        # print(level)

            if debt[j] == '' or debt[j] == '——':
                debt[j] = 'undefined'
            data_5.append(debt[j])

            if hope[j] == '' or hope[j] == '——':
                hope[j] = 'undefined'
            data_5.append(hope[j])

            data_5.append(date[j])

# 其他
    html_6 = htmllist_6[0]
    name_6 = re.findall(findName, html_6)
    pdf_6 = re.findall(findPDF, html_6)
    level_6 = re.findall(findLevel, html_6)
    debt_6 = re.findall(findDebt, html_6)
    hope_6 = re.findall(findHope, html_6)
    date_6 = re.findall(findDate, html_6)

    # print(len(name_6))
    # print(name_6)
    # print(pdf_6)
    # print(level_6)
    # print(debt_6)
    # print(hope_6)
    # print(date_6)

    level_6[0] = 'undefined'
    debt_6[0] = 'undefined'
    hope_6[0] = 'undefined'
    date_6[0] = 'undefined'

    data_6.append(name_6[0])
    pdf_6[0] = "http://www.ccxr.com.cn/pdf/" + pdf_6[0]
    data_6.append(pdf_6[0])
    data_6.append(level_6[0])
    data_6.append(debt_6[0])
    data_6.append(hope_6[0])
    data_6.append(date_6[0])

    # print(data)
    # print("_"*40)
    # print(data_2)
    # print("_"*40)
    # print(data_3)
    # print("_"*40)
    # print(data_4)
    # print("_"*40)
    # print(data_5)
    # print("_"*40)
    # print(data_6)

    return data, data_2, data_3, data_4, data_5, data_6


def saveData(data, data_2, data_3, data_4, data_5, data_6, savepath):
    print("save...")
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet = book.add_sheet('中诚信证券评级公告', cell_overwrite_ok=True)
    col = ("类别", "名称", "pdf链接", "等级", "债项", "展望", "日期")
    for i in range(0, 7):
        sheet.write(0, i, col[i])

# 普通公司债
    for i in range(0, 1441):
        sheet.write(i+1, 0, "普通公司债")
        for j in range(0, 6):
            sheet.write(i+1, j+1, data[j+6*i])

# 可交换公司债
    for i in range(0, 10):
        sheet.write(i+1442, 0, "可交换公司债")
        for j in range(0, 6):
            sheet.write(i+1442, j+1, data_2[j+6*i])

# 可转换公司债
    for i in range(0, 86):
        sheet.write(i+1452, 0, "可转换公司债")
        for j in range(0, 6):
            sheet.write(i+1452, j+1, data_3[j+6*i])

# 定期跟踪
    for i in range(0, 2575):
        sheet.write(i+1538, 0, "定期跟踪")
        for j in range(0, 6):
            sheet.write(i+1538, j+1, data_4[j+6*i])

# 不定期跟踪
    for i in range(0, 192):
        sheet.write(i+4113, 0, "不定期跟踪")
        for j in range(0, 6):
            sheet.write(i+4113, j+1, data_5[j+6*i])

# 其他
    for i in range(0, 1):
        sheet.write(i+4305, 0, "其他")
        for j in range(0, 6):
            sheet.write(i+4305, j+1, data_6[j+6*i])

    book.save(savepath)


if __name__ == "__main__":
    main()
    print("爬取完毕")

