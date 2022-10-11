#coding:utf-8
from bs4 import BeautifulSoup # 抓标签里面元素的方法
import re
import random
import pandas as pd
import os
import datetime
import time
import xlwings
import xlsxwriter
import math
import requests
import json
import sys
from queue import Queue
from settings_sso import Settings_sso
from sqlalchemy import create_engine
from settings import Settings
from emailControl import EmailControl
from multiprocessing.dummy import Pool
from fake_user_agent import user_agent   # 用fake_useragent模块来设置一个请求头，用来进行伪装成浏览器

from random import randint

# -*- coding:utf-8 -*-
class QueryTwo(Settings, Settings_sso):
    def __init__(self, userMobile, password):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self._online()
        # self.sso_online_Two()
        self.engine1 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql1['user'],
                                                                                    self.mysql1['password'],
                                                                                    self.mysql1['host'],
                                                                                    self.mysql1['port'],
                                                                                    self.mysql1['datebase']))
        self.engine2 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql2['user'],
                                                                                    self.mysql2['password'],
                                                                                    self.mysql2['host'],
                                                                                    self.mysql2['port'],
                                                                                    self.mysql2['datebase']))
        self.engine20 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql20['user'],
                                                                                    self.mysql20['password'],
                                                                                    self.mysql20['host'],
                                                                                    self.mysql20['port'],
                                                                                    self.mysql20['datebase']))
        self.engine3 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql3['user'],
                                                                                    self.mysql3['password'],
                                                                                    self.mysql3['host'],
                                                                                    self.mysql3['port'],
                                                                                    self.mysql3['datebase']))
        self.e = EmailControl()
    def reSetEngine(self):
        self.engine1 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql1['user'],
                                                                                    self.mysql1['password'],
                                                                                    self.mysql1['host'],
                                                                                    self.mysql1['port'],
                                                                                    self.mysql1['datebase']))
        self.engine2 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql2['user'],
                                                                                    self.mysql2['password'],
                                                                                    self.mysql2['host'],
                                                                                    self.mysql2['port'],
                                                                                    self.mysql2['datebase']))
    # 获取签收表内容
    def readFormHost(self):
        start = datetime.datetime.now()
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath)
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath):
        fileType = os.path.splitext(filePath)[1]
        app = xlwings.App(visible=False, add_book=False)
        app.display_alerts = False
        if 'xls' in fileType:
            wb = app.books.open(filePath, update_links=False, read_only=True)
            for sht in wb.sheets:
                if sht.api.Visible == -1:
                    try:
                        db = None
                        tm = ''
                        # print(sht.name)
                        # db = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False, dtype=str).value
                        # db = pd.read_excel(filePath, sheet_name=sht.name, header=1, names=int, index_col=False, dtype=str)
                        db = pd.read_excel(filePath, sheet_name=sht.name)
                        if '运单编号' in db.columns:
                            tm = '运单编号'
                        elif '查件单号' in db.columns:
                            tm = '查件单号'
                        elif '运单号' in db.columns:
                            tm = '运单号'
                        elif '物流单号' in db.columns:
                            tm = '物流单号'
                        db = db[[tm]]
                        db[tm] = db[tm].astype(str)
                        db.dropna(axis=0, how='any', inplace=True)                  # 空值（缺失值），将空值所在的行/列删除后
                    except Exception as e:
                        print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                    if db is not None and len(db) > 0:
                        # print(db)
                        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
                        print('++++正在获取：' + sht.name + ' 表；共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                        # 将获取到的运单号 查询轨迹
                        self.SearchGoods(db, tm)
                    else:
                        print('----------数据为空,查询失败：' + sht.name)
                else:
                    print('----不需查询：' + sht.name)
            wb.close()
        app.quit()

    #、随机验证码
    def SearchGoods(self, db, tm):
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        orderId = list(db[tm])
        print(orderId)
        max_count = len(orderId)                 # 使用len()获取列表的长度，上节学的
        if max_count > 0:
            df = pd.DataFrame([])                # 创建空的dataframe数据框
            dlist = []
            # pool = Pool(4)
            # result = pool.map(self._SearchGoods,orderId)
            for ord in orderId:
                print(ord)
                data = self._SearchGoods(ord)
                if data is not None and len(data) > 0:
                    dlist.append(data)
                # print('暂停5秒')
                # time.sleep(15)
            dp = df.append(dlist, ignore_index=True)
            # dp.dropna(axis=0, how='any', inplace=True)
        else:
            dp = None
        print(dp)
        dp.to_excel('G:\\输出文件\\新竹快递-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
        print('查询已导出+++')
        print('*' * 50)

    def _SearchGoods(self,wayBillNumber):
        # 生成随机的User-Agent
        USER_AGENTS = [ "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Mobile Safari/537.36",
                        "User-Agent:Mozilla/5.0",
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"
        ]
        random_agent = USER_AGENTS[randint(0, len(USER_AGENTS) - 1)]

        # 使用代理服务器
        proxy_list = ['47.242.154.178:37466', '47.242.154.178:37467', '47.242.154.178:37468', '47.242.154.178:37469', '47.242.154.178:37460',
                 '39.105.167.0:37466', '39.105.167.0:37467', '39.105.167.0:37468', '39.105.167.0:37469', '39.105.167.0:37460',
                 '47.242.154.178:46566', '47.242.85.200:46566', '47.242.85.200:46565',
                 '39.105.167.0:17467']
        proxy = random.choice(proxy_list)

        #0、获取验证码
        code = ''
        for i in range(4):
            code += str(random.randint(0,9))
        #1、构建url
        url = "https://www.hct.com.tw/search/searchgoods_n.aspx"   #url为机器人的webhook
        #2、构建一下请求头部
        r_header = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
                    "Accept-Encoding": "gzip, deflate, br",
                    "Accept-Language": "zh-CN,zh;q=0.9",
                    "Cache-Control": "no-cache",
                    # "Connection": "keep-alive",
                    "Connection": "close",
                    "Content-Length": '92',
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Host": "www.hct.com.tw",
                    'Origin': 'https://www.hct.com.tw',
                    # "Pragma": "no-cache",
                    "Pragma": "1",
                    "Referer": "https://www.hct.com.tw/search/searchgoods_n.aspx",
                    "Sec-Fetch-Mode": "navigate",
                    "Sec-Fetch-Site": "same-origin",
                    "Sec-Fetch-User": "?1",
                    "Upgrade-Insecure-Requests": '1',
                    "Charset": "UTF-8",
                    # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"
                    "User-Agent": user_agent()
                    # "User-Agent": random.choice(user_agent_list)
                    }

        #3、构建请求数据
        data = {'__VIEWSTATE': 'HmTo9prl6CO4ytnFMzgfgTsdbJ5MSx7l5gm0chzu2Wx+HKF1cFyEPs1OAwLWOlymmInrgTgSPdY75BwFB3qB0JDXY02XSC14LXp/dC4hZBrHB66Fe5CxoJng7cw=',
                'ctl00$ContentFrame$txtpKey': wayBillNumber,
                'ctl00$ContentFrame$txtpKey2': '',
                'ctl00$ContentFrame$txtpKey3': '',
                'ctl00$ContentFrame$txtpKey4': '',
                'ctl00$ContentFrame$txtpKey5': '',
                'ctl00$ContentFrame$txtpKey6': '',
                'ctl00$ContentFrame$txtpKey7': '',
                'ctl00$ContentFrame$txtpKey8': '',
                'ctl00$ContentFrame$txtpKey9': '',
                'ctl00$ContentFrame$txtpKey10': '',
                'ctl00$ContentFrame$b13ca230fd18402cad0febf14d8a11bc': code,
                'ctl00$ContentFrame$Button1': '查詢'
                }

        # proxy = '47.242.154.178:37466'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}

        # 检测代理ip是否使用
        rq_ip = requests.get('http://httpbin.org/ip', proxies=proxies, timeout=3)
        # rq_ip = json.loads(rq_ip.text)
        # print('代理ip(一)：' + rq_ip['origin'])

        # 使用代理ip发送请求
        req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # req = self.session.post(url=url, headers=r_header, data=data, verify=False)
        # req = requests.post(url=url, headers=r_header, data=data, verify=False)
        # print(req.headers)

        soup = BeautifulSoup(req.text, 'lxml')      # 创建 beautifulsoup 对象
        no = soup.input.get('value')
        chk = soup.input.next_sibling.get('value')
        # print(no)
        # print(chk)
        # print('----------获取验证值成功-------------')

        # time.sleep(1)
        url = "https://www.hct.com.tw/search/SearchGoods.aspx"   #url为机器人的webhook
        r_header = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
                    "Accept-Encoding": "gzip, deflate, br",
                    "Accept-Language": "zh-CN,zh;q=0.9",
                    "Cache-Control": "no-cache",
                    # "Connection": "keep-alive",
                    "Connection": "close",
                    "Content-Length": '92',
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Host": "www.hct.com.tw",
                    'Origin': 'https://www.hct.com.tw',
                    # "Pragma": "no-cache",
                    "Pragma": "1",
                    "Referer": "https://www.hct.com.tw/search/searchgoods_n.aspx",
                    "Sec-Fetch-Mode": "navigate",
                    "Sec-Fetch-Site": "same-origin",
                    "Sec-Fetch-User": "?1",
                    "Upgrade-Insecure-Requests": '1',
                    "Charset": "UTF-8",
                    # "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"
                    "User-Agent": user_agent()
                    # "User-Agent": random.choice(user_agent_list)
                    }
        data = {'no': no,
                'chk': chk
                }

        # proxy = '47.242.154.178:37466'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}

        # 检测代理ip是否使用
        rq_ip = requests.get('http://httpbin.org/ip', proxies=proxies, timeout=3)
        # rq_ip = json.loads(rq_ip.text)
        # print('代理ip(二)：' + rq_ip['origin'])

        # 使用代理ip发送请求
        req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # req = self.session.post(url=url, headers=r_header, data=data, verify=False)
        # req = requests.post(url=url, headers=r_header, data=data, verify=False)

        # print(req)
        # print('----------数据获取返回成功-----------')
        rq = BeautifulSoup(req.text, 'lxml')
        rq = rq.find_all('tr') 
        ordersDict = []
        online_time = ''
        res = ''
        res2 = ''
        for index, val in enumerate(rq):
            result = {}
            # print(val)
            if "請檢查您的查貨資料。" in str(val):            # 没有查询到货态
                data = pd.DataFrame([[res, res2, '', '', online_time]], columns=['查件单号', '查货时间', '轨迹时间', '轨迹内容', '上线时间'])
                # data.dropna(axis=0, how='any', inplace=True)
                data.dropna(axis=0, subset=["查件单号"], inplace=True)
                data.sort_values(by="轨迹时间", inplace=True, ascending=True)  # inplace: 原地修改; ascending：升序
                # print(data)
                return data
            if "ctl00_ContentFrame_lblInvoiceNo" in str(val):       # 查询到货态（一）
                result['序号'] = index
                result['查件单号'] = str(val).split(r'ctl00_ContentFrame_lblInvoiceNo">')[1].split('</')[0]
                result['查货时间'] = str(val).split('時間：')[1].split('</')[0]
                res = result['查件单号'].strip()
                res2 = result['查货时间']
            if "L_time" in str(val):                                # 查询到货态（二）
                if "貨件已退回" in str(val):
                    # pattern = re.compile(r'<[^>]+>', re.S)
                    L_time = str(val).split('L_time">')[1].split('</')[0]
                    L_cls = str(val).split(')">')[1]
                    L_cls = L_cls.split('</u>')[0].replace('<font color="blue"><u color="blue">',' ')
                elif "送達。貨物件" in str(val):
                    L_time = str(val).split('L_time">')[1].split('</')[0]
                    L_cls = str(val).split('L_cls">')[1]
                    L_cls = L_cls.split('。</')[0].replace('</a>',' ')
                elif "原貨號：" in str(val):
                    L_time = str(val).split('L_time">')[1].split('</')[0]
                    L_cls = str(val).split(')">')[1]
                    L_cls = L_cls.split('</u>')[0].replace('<font color="blue"><u color="blue">',' ')
                else:
                    pattern = re.compile(r'<[^>]+>', re.S)
                    L_time = str(val).split('L_time">')[1].split('</')[0]
                    L_cls = str(val).split('L_cls">')[1]
                    L_cls = pattern.sub('', L_cls)

                if '貨件已抵達' in str(val) and ('土城營業所' in str(val) or '桃園營業所' in str(val)) and '貨件整理中' in str(val) or '集荷' in str(val):
                    online_time = L_time
                # print(L_time) 
                # print(res)
                result['序号'] = index
                result['查件单号'] = res
                result['查货时间'] = res2
                result['轨迹时间'] = L_time
                result['轨迹内容'] = L_cls.replace("\n\n", "").strip()
                result['上线时间'] = online_time
            ordersDict.append(result)
        # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        data = pd.json_normalize(ordersDict)
        # print(data)
        data.dropna(axis=0, how='any', inplace=True)
        data.sort_values(by=["查件单号", "轨迹时间"], inplace=True, ascending=[True, True])
        # data.sort_values(by="轨迹时间", inplace=True, ascending=True)  # inplace: 原地修改; ascending：升序
        # print(data)
        # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # data.to_excel('G:\\输出文件\\新竹快递 {0} .xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        return data

if __name__ == '__main__':
    m = QueryTwo('+86-18538110674', 'qyz04163510')
    start: datetime = datetime.datetime.now()
    print(datetime.datetime.now())
    '''
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    '''
    m.readFormHost()

    # m._SearchGoods('7532082106')

    print('查询耗时：', datetime.datetime.now() - start)