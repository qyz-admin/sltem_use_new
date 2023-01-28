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
                data = self._SearchGoodsT(ord)
                if data is not None and len(data) > 0:
                    dlist.append(data)
                # print('暂停5秒')
                # time.sleep(15)
            dp = df.append(dlist, ignore_index=True)
            nan_value = float("NaN")   # 用null替换所有空位，然后用dropna函数删除所有空值列。
            dp.replace("", nan_value, inplace=True)
            dp.dropna(axis=0, how='any', inplace=True, subset=['运单号'])
            # dp.dropna(axis=0, how='any', inplace=True)
        else:
            dp = None
        print(dp)
        dp.to_sql('tem', con=self.engine1, index=False, if_exists='replace')
        dp.to_excel('G:\\输出文件\\黑猫宅急便-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
        print('查询已导出+++')
        print('*' * 50)
    def _SearchGoodsT(self,wayBillNumber):
        #0、获取验证码
        code = ''
        for i in range(4):
            code += str(random.randint(0,9))
        #1、构建url
        url = "https://www.t-cat.com.tw/Inquire/TraceDetail.aspx?BillID=" + wayBillNumber   #url为机器人的webhook
        #2、构建一下请求头部
        r_header = {"Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
                    "Accept-Encoding": "gzip, deflate, br",
                    "Accept-Language": "zh-CN,zh;q=0.9",
                    "Cache-Control": "no-cache",
                    "Connection": "keep-alive",
                    # "Connection": "close",
                    # "Content-Length": '92',
                    # "Content-Type": "application/x-www-form-urlencoded",
                    "Host": "www.t-cat.com.tw",
                    'Origin': 'https://www.t-cat.com.tw',
                    # "Pragma": "no-cache",
                    # "Pragma": "1",
                    "Referer": "https://www.t-cat.com.tw/Inquire/trace.aspx",
                    "Sec-Fetch-Mode": "navigate",
                    "Sec-Fetch-Site": "same-origin",
                    "Sec-Fetch-User": "?1",
                    "Upgrade-Insecure-Requests": '1',
                    "Charset": "UTF-8",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36"
                    # "User-Agent": user_agent()
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
        data = {'BillID': 620430597712}
        proxy = '47.242.154.178:37466'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}

        # 检测代理ip是否使用
        # rq_ip = requests.get('http://httpbin.org/ip', proxies=proxies, timeout=3)
        # rq_ip = json.loads(rq_ip.text)
        # print('代理ip(一)：' + rq_ip['origin'])

        # 使用代理ip发送请求
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # req = self.session.post(url=url, headers=r_header, data=data, verify=False)
        # req = requests.post(url=url, headers=r_header, data=data, verify=False)
        # print(req.headers)
        # print(req)
        # print('----------数据获取返回成功-----------')
        rq = BeautifulSoup(req.text, 'lxml')
        rq = rq.find_all('tr')
        # print(88)
        # print(rq)
        # print('-' * 50)
        ordersDict = []
        L_waybill = ''
        L_time = ''
        L_info = ''
        L_info2 = ''
        L_info3 = ''
        for index, val in enumerate(rq):
            result = {}
            # print(index)
            # print(val)
            # print('-' * 10)
            if "top" in str(val) and "time" not in str(val):            # 没有查询到货态
                # result = pd.DataFrame([['', '', '', '']], columns=['查件单号', '轨迹时间', '轨迹内容', '負責營業所'])
                # data.dropna(axis=0, how='any', inplace=True)
                # data.dropna(axis=0, subset=["查件单号"], inplace=True)
                # data.sort_values(by="轨迹时间", inplace=True, ascending=True)  # inplace: 原地修改; ascending：升序
                result['序号'] = ""
                result['订单编号'] = ""
                result['运单号'] = ""
                result['物流轨迹时间'] = ""
                result['出货时间'] = ""
                result['上线时间'] = ""
                result['保管时间'] = ""
                result['完成时间'] = ""
                result['物流轨迹'] = ""
                result['負責營業所'] = ""
                result['轨迹备注'] = ""
                # print(result)
                # print(808)
                # return result
            if "height" in str(val) and "rowspan" in str(val):       # 查询到货态（一）
                L_waybill = str(val).split('</span>')[0].split('bl12">')[1]
                L_time_val = str(val).split('<br/>')
                L_time10 = L_time_val[0].split('bl12">')
                L_time11 = L_time_val[0].split('bl12">')[len(L_time10)-1]
                L_time22 = L_time_val[1].split('</span>')[0]
                L_time = L_time11 + '' + L_time22
                if 'strong' in str(val):
                    L_info = str(val).split('<strong>')[1].split('</strong>')[0]
                else:
                    L_info = str(val).split('bl12">')[2].split('</span>')[0]
                L_info2 = str(val).split('foothold.aspx?n=')[1].split('</a>')[0]
                L_info2 = L_info2.split('>')[1]
                L_info3 = str(val).split('title=')[1].split('>')[0]
                result['序号'] = index
                result['订单编号'] = ""
                result['运单号'] = L_waybill
                result['物流轨迹时间'] = L_time
                result['出货时间'] = ""
                result['上线时间'] = ""
                result['保管时间'] = ""
                result['完成时间'] = ""
                result['物流轨迹'] = L_info
                result['轨迹备注'] = L_info3
                result['負責營業所'] = L_info2
                result['轨迹备注'] = L_info3
                if '包裹已經送達收件人' in L_info3:
                    result['完成时间'] = L_time
                elif '順利送達' in L_info:
                    result['完成时间'] = L_time
            if "height" not in str(val) and "rowspan" not in str(val) and "<br/>" in str(val):
                L_time_val = str(val).split('<br/>')
                L_time10 = L_time_val[0].split('bl12">')
                L_time11 = L_time_val[0].split('bl12">')[len(L_time10)-1]
                L_time22 = L_time_val[1].split('</span>')[0]
                L_time = L_time11 + '' + L_time22
                if 'strong' in str(val):
                    L_info = str(val).split('<strong>')[1].split('</strong>')[0]
                else:
                    L_info = str(val).split('bl12">')[1].split('</span>')[0]
                L_info2 = str(val).split('foothold.aspx?n=')[1].split('</a>')[0]
                L_info2 = L_info2.split('>')[1]
                L_info3 = str(val).split('title=')[1].split('>')[0]
                result['序号'] = index
                result['订单编号'] = ""
                result['运单号'] = L_waybill
                result['物流轨迹时间'] = L_time
                result['出货时间'] = ""
                result['上线时间'] = ""
                result['保管时间'] = ""
                result['完成时间'] = ""
                result['物流轨迹'] = L_info
                result['負責營業所'] = L_info2
                result['轨迹备注'] = L_info3

                if 'sd已經至寄件人指定地點收到包裹' in L_info3:
                    result['出货时间'] = L_time
                elif '已集貨' in L_info:
                    result['出货时间'] = L_time

                if '包裹正從營業所送到轉運中心，或從轉運中心送到營業所' in L_info3:
                    result['上线时间'] = L_time
                elif '轉運中' in L_info or 'sd正在將包裹配送到收件人途中' in L_info3:
                    result['上线时间'] = L_time

                if '暫置營業所保管中' in L_info:
                    result['保管时间'] = L_time
                elif '調查處理中' in L_info or '另約時間' in L_info:
                    result['保管时间'] = L_time

                if '包裹已經送達收件人' in L_info3:
                    result['完成时间'] = L_time
                elif '順利送達' in L_info:
                    result['完成时间'] = L_time

            # print(858)
            # print(result)
            # print(848)
            ordersDict.append(result)
        # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        data = pd.json_normalize(ordersDict)
        data.sort_values(by=["运单号", "物流轨迹时间"], inplace=True, ascending=[True, True])
        print(data)
        # print(99)
        # data.dropna(axis=0, how='any', inplace=True)
        # data.dropna(axis=0, how='any', inplace=True, subset=['查件单号'])
        # data.sort_values(by=["查件单号", "轨迹时间"], inplace=True, ascending=[True, True])
        # data.sort_values(by="轨迹时间", inplace=True, ascending=True)  # inplace: 原地修改; ascending：升序
        # print(data)
        # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # data.to_excel('G:\\输出文件\\黑猫宅急便 {0} .xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
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

    # m._SearchGoodsT('620434003756')

    print('查询耗时：', datetime.datetime.now() - start)