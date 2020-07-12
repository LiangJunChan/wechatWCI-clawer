# coding: utf-8
import time
import random
import json
import requests
from requests.adapters import HTTPAdapter
import ssl
import sys
import importlib
import openpyxl
from datetime import date, datetime, timedelta
ssl._create_default_https_context = ssl._create_unverified_context

importlib.reload(sys)
s = requests.Session()
requests.adapters.DEFAULT_RETRIES = 5
s.mount('http://', HTTPAdapter(max_retries=5))
s.mount('https://', HTTPAdapter(max_retries=5))


def search(open_id, search_key, keyword):

    s_url = "https://search.weixin.qq.com/cgi-bin/searchweb/wxindex/querywxindexgroup?wxindex_query_list=%s&gid=&openid=%s&search_key=%s" % (
        keyword, open_id, search_key)
    headers = {
        "Referer":
        "https://servicewechat.com/wxc026e7662ec26a3a/7/page-frame.html",
        "User-Agent":
        "Mozilla/5.0 (Linux; Android 6.0.1; Nexus 5 Build/M4B30Z; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/44.0.2403.117 Mobile Safari/537.36 MicroMessenger/6.7.3.1360(0x26070333) NetType/WIFI Language/zh_CN Process/appbrand0",
    }
    ss = s.get(s_url, headers=headers, timeout=3, verify=False)
    content = json.loads(ss.content)

    wxindex_str = content["data"]["group_wxindex"][0]["wxindex_str"]

    if wxindex_str == "":
        pass

    return wxindex_str.split(",")


def sleep():
    time.sleep(random.randint(2, 4))


if __name__ == '__main__':

    # 爬取的 keyword，这里以创造营成员为例子
    keywords = [
        '张艺凡', "陈卓璇", '希林娜依·高', '赵粤', '王艺瑾', '郑乃馨', '刘些宁','创造营'
    ]
    # open_id 及search_key 值需要从抓包软件抓取小程序
    open_id = ""
    search_key = ""

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['name', 'type', 'value', 'date'])
    for keyword in keywords:
        wis = search(open_id, search_key, keyword)
        # 微信指数查看的是近 90 天的指数值，因此从 90 天前初始化
        day2 =90
        for data in wis:
            day = date.today()
            now = datetime.now()
            t = timedelta(days=day2)
            date1 = now -t
            ws.append([keyword, "czy2020", data,"{}".format(date1.strftime('%Y-%m-%d')) ])
            day2 -=1
            print(data)

    wb.save("weixin.xlsx")
