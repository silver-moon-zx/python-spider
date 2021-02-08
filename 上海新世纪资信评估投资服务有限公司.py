# -*- coding = utf-8 -*-
# @Time : 2020/9/8 22:47
# @Author : 陈子翔
# @File : 上海新世纪资信评估投资服务有限公司.py
# @software : PyCharm

import urllib.request, urllib.error, urllib.parse
import http.cookiejar
import json
import xlwt


def main():
    """
    一、首次评级：
        短期融资券：cid: 52
                limit: 10
                page: 1~164
        中期票据：cid: 53
                limit: 10
                page: 1~164
        公司债券：cid: 54
                limit: 10
                page: 1~63
        企业债券：cid: 55
                limit: 10
                page: 1~46
        中小企业集合票据：cid: 56
                       limit: 10
                        page: 1~5
        金融债券：cid: 57
                limit: 10
                page: 1~14
        资产支持证券：cid: 58
                    limit: 10
                    page: 1~19
        地方政府债券：cid: 59
                    limit: 10
                    page: 1~83
        其他：cid: 60
            limit: 10
            page: 1~4
    二、跟踪评级：
        短期融资券：cid: 62
                limit: 10
                page: 1~144
        中期票据：cid: 63
                limit: 10
                page: 1~424
        公司债券：cid: 64
                limit: 10
                page: 1~168
        企业债券：cid: 65
                limit: 10
                page: 1~192
        中小企业集合票据：cid: 66
                       limit: 10
                        page: 1~9
        金融债券：cid: 67
                limit: 10
                page: 1~39
        资产支持证券：cid: 68
                    limit: 10
                    page: 1
        地方政府债券：cid: 69
                    limit: 10
                    page: 1~4
        其他：cid: 70
            limit: 10
            page: 1
    三、其他信息：
        重点关注：cid: 73
                limit: 10
                page: 1~53
        级别调整：cid: 72
                limit: 10
                page: 1~14
        其他公告：cid: 74
                limit: 10
                page: 1~7
    """
    url = "http://www.shxsj.com/serve/api/article_api/get_cid_article"
    datalist_first, datalist_later, datalist_other = create_FormData()
    list_1 = []
    list_2 = []
    for form_Data in datalist_first:
        data = getData(url, form_Data)
        # print(data)
        list_1.append(data)
    for form_Data in datalist_later:
        data = getData(url, form_Data)
        list_2.append(data)

    savepath = ".\\上海新世纪资信评级公告.xls"
    saveData(list_1, list_2, savepath)


def create_FormData():

    # 首次评级
    datalist_first = []
    for i in range(0, 164):
            page = i+1
            form_Data = [52, 10, page]
            datalist_first.append(form_Data)
    for i in range(0, 164):
            page = i+1
            form_Data = [53, 10, page]
            datalist_first.append(form_Data)
    for i in range(0, 63):
            page = i+1
            form_Data = [54, 10, page]
            datalist_first.append(form_Data)
    for i in range(0, 46):
            page = i+1
            form_Data = [55, 10, page]
            datalist_first.append(form_Data)
    # for i in range(0, 5):
    #         page = i+1
    #         form_Data = [56, 10, page]
    #         datalist_first.append(form_Data)
    # for i in range(0, 14):
    #         page = i+1
    #         form_Data = [57, 10, page]
    #         datalist_first.append(form_Data)
    # for i in range(0, 19):
    #         page = i+1
    #         form_Data = [58, 10, page]
    #         datalist_first.append(form_Data)
    # for i in range(0, 83):
    #         page = i+1
    #         form_Data = [59, 10, page]
    #         datalist_first.append(form_Data)
    # for i in range(0, 4):
    #         page = i+1
    #         form_Data = [60, 10, page]
    #         datalist_first.append(form_Data)

    # 跟踪评级
    datalist_later = []
    for j in range(0, 144):
            page = j+1
            form_Data = [62, 10, page]
            datalist_later.append(form_Data)
    for j in range(0, 424):
            page = j+1
            form_Data = [63, 10, page]
            datalist_later.append(form_Data)
    for j in range(0, 168):
            page = j+1
            form_Data = [64, 10, page]
            datalist_later.append(form_Data)
    for j in range(0, 192):
            page = j+1
            form_Data = [65, 10, page]
            datalist_later.append(form_Data)
    # for j in range(0, 9):
    #         page = j+1
    #         form_Data = [66, 10, page]
    #         datalist_later.append(form_Data)
    # for j in range(0, 39):
    #         page = j+1
    #         form_Data = [67, 10, page]
    #         datalist_later.append(form_Data)
    # for j in range(0, 1):
    #         page = j+1
    #         form_Data = [68, 10, page]
    #         datalist_later.append(form_Data)
    # for j in range(0, 4):
    #         page = j+1
    #         form_Data = [69, 10, page]
    #         datalist_later.append(form_Data)
    # for j in range(0, 1):
    #         page = j+1
    #         form_Data = [70, 10, page]
    #         datalist_later.append(form_Data)

    # 其他公告
    datalist_other = []
    for k in range(0, 53):
            page = k+1
            form_Data = [73, 10, page]
            datalist_other.append(form_Data)
    for k in range(0, 14):
            page = k+1
            form_Data = [72, 10, page]
            datalist_other.append(form_Data)
    for k in range(0, 7):
            page = k+1
            form_Data = [74, 10, page]
            datalist_other.append(form_Data)

    return datalist_first, datalist_later, datalist_other


def askURL(url, form_Data):
    # 获取网页cookie，构造opener
    filename = 'wk_cookie.txt'
    cookie = http.cookiejar.LWPCookieJar(filename)
    handler = urllib.request.HTTPCookieProcessor(cookie)
    opener = urllib.request.build_opener(handler)

    # 模拟Request Headers
    headers = {
        "Cookie": "menu_id=3; menu_opennames=[%2251%22]; template_id=4;page_id=57; active_name=51-57",
        "Host": "www.shxsj.com",
        "Origin": "http://www.shxsj.com",
        "Referer": "http://www.shxsj.com/page?template=4&pageid=57&mid=3",
        "User-Agent": "Mozilla / 5.0(Windows NT 10.0;Win64;x64) AppleWebKit / 537.36(KHTML, likeGecko) Chrome / 84.0.4147.105Safari / 537.36",
        "X-Requested-With": "XMLHttpRequest"
    }

    # 模拟POST的data内容
    form_Data = {
        "cid": form_Data[0],
        "limit": form_Data[1],
        "page": form_Data[2],
        "sort_val": '',
        "sort_key": '',
        "rank_stime": '',
        "rank_etime": '',
        "keyword": '',
        "rank_type": '',
        "main_type": '',
        "rank_level": '',
        "bound_type": ''
    }

    data = bytes(urllib.parse.urlencode(form_Data), encoding='utf-8')
    request = urllib.request.Request(url, data=data, headers=headers, method="POST")
    response = opener.open(request, timeout=5)
    html_json = response.read().decode('utf-8')
    html = json.loads(html_json)
    return html


def getData(url, form_Data):
    datas = []
    data = {}
    html = askURL(url, form_Data)
    origin_dict = html['data']
    second_dict = origin_dict['list']
    datalist = second_dict['data']    # 列表
    for dict in datalist:
        # print(type(dict))
        # print(dict)
        # for key in dict.keys():
        #     print(key)
        pdf_list = dict.get('pdf')

        project_dict = dict.get('project')
        # 问题所在地！！！！！！！！！！！！！！！！！
        if pdf_list == None or project_dict == None:
            issuer = dict['issuer']
            rank_type = dict['rank_type']
            main_type = dict['main_type']
            rank_level = dict['rank_level']
            bound_type = dict['bound_type']
            rank_time = dict['rank_time']
            name = dict.get('name')
            url = dict.get('url')
            if name == None or url == None:
                name_url = 'None'
            else:
                name_url = name + ':' + "http://www.shxsj.com" + url
        elif len(pdf_list) != 0:
            pdf_dict = pdf_list[0]
            # 1. 项目名称：'issuer'
            issuer = project_dict['issuer']
            # 2. 评级类别：'rank_type'
            rank_type = project_dict['rank_type']
            # 3. 主体级别：'main_type'
            main_type = project_dict['main_type']
            # 4. 评级展望：'rank_level'
            rank_level = project_dict['rank_level']
            # 5. 债项级别：'bound_type'
            bound_type = project_dict['bound_type']
            # 6. 评级时间：'rank_time'
            rank_time = project_dict['rank_time']
            # 7. 公告名称：'name'
            name = pdf_dict['name']
            # 8. 公告链接：'url'
            url = pdf_dict['url']

            name_url = name + ':' + "http://www.shxsj.com" + url
        else:
            # 1. 项目名称：'issuer'
            issuer = project_dict['issuer']
            # 2. 评级类别：'rank_type'
            rank_type = project_dict['rank_type']
            # 3. 主体级别：'main_type'
            main_type = project_dict['main_type']
            # 4. 评级展望：'rank_level'
            rank_level = project_dict['rank_level']
            # 5. 债项级别：'bound_type'
            bound_type = project_dict['bound_type']
            # 6. 评级时间：'rank_time'
            rank_time = project_dict['rank_time']

            name_url = 'None'


        data = {
            'issuer': issuer,
            'rank_type': rank_type,
            'main_type': main_type,
            'rank_level': rank_level,
            'bound_type': bound_type,
            'rank_time': rank_time,
            'name_url': name_url
        }

        datas.append(data)
    return datas




def saveData(list_1, list_2, savepath):
    print("save...")
    len_1 = 0
    len_2 = 0
    book = xlwt.Workbook(encoding="utf-8", style_compression=0)
    sheet_1 = book.add_sheet('首次评级', cell_overwrite_ok=True)
    sheet_2 = book.add_sheet('跟踪评级', cell_overwrite_ok=True)
    col = ("项目名称", "评级类别", "主体级别", "评级展望", "债项级别", "评级时间", "公告链接")
    for i in range(0, 7):
        sheet_1.write(0, i, col[i])
        sheet_2.write(0, i, col[i])

    # for data_1 in list_1:
    #     len_1 = len(data_1) + len_1
    # for data_2 in list_2:
    #     len_2 = len(data_2) + len_2
    i = 0
    for data_1 in list_1:
        for dict_1 in data_1:
            i = i + 1
            j = 0
            for value in dict_1.values():
                j = j + 1
                sheet_1.write(i, j-1, value)
    k = 0
    for data_2 in list_2:
        for dict_2 in data_2:
            k = k + 1
            j = 0
            for value in dict_2.values():
                j = j + 1
                sheet_2.write(k, j-1, value)

    book.save(savepath)



if __name__ == "__main__":
    main()
    print("爬取完毕")
