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
    def readFormHost(self, proxy_id, proxy_handle):
        start = datetime.datetime.now()
        path = r'F:\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath, proxy_id, proxy_handle)
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath, proxy_id, proxy_handle):
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
                        self.SearchGoods(db, tm, proxy_id, proxy_handle)
                    else:
                        print('----------数据为空,查询失败：' + sht.name)
                else:
                    print('----不需查询：' + sht.name)
            wb.close()
        app.quit()

    #、随机验证码
    def SearchGoods(self, db, tm, proxy_id, proxy_handle):
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
                data = self._SearchGoods(ord, proxy_id, proxy_handle)
                if data is not None and len(data) > 0:
                    dlist.append(data)
                time.sleep(30)
            dp = df.append(dlist, ignore_index=True)
            # dp.dropna(axis=0, how='any', inplace=True)
        else:
            dp = None
        print(dp)
        dp.to_excel('F:\\输出文件\\宅配通快递-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
        print('查询已导出+++')
        print('*' * 50)

    def SearchGoods_write(self, proxy_id, proxy_handle):
        sql = '''SELECT * FROM 宅配通 z
                    WHERE z.`运单编号` NOT IN (SELECT DISTINCT wayBillNumber FROM 宅配通记录 p GROUP BY wayBillNumber);'''
        db = pd.read_sql_query(sql=sql, con=self.engine1)
        orderId = list(db['运单编号'])
        print(orderId)
        max_count = len(orderId)                 # 使用len()获取列表的长度，上节学的
        print(max_count)
        if max_count > 0:
            for ord in orderId:
                print(ord)
                data = self._SearchGoods(ord, proxy_id, proxy_handle)
                if data is not None and len(data) > 0:
                    data.to_sql('query_cache', con=self.engine1, index=False, if_exists='replace')
                    sql = '''REPLACE INTO 宅配通记录(序号,	orderNumber,wayBillNumber,track_date,出货时间,上线时间,保管时间,完成时间,track_info,track_status,负责营业所,轨迹备注,便利店) 
                            SELECT 序号,	orderNumber,wayBillNumber,track_date,出货时间,上线时间,保管时间,完成时间,track_info,track_status,负责营业所,轨迹备注,便利店
                            FROM query_cache;'''
                    pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
                    print('查询已写入+++')
                time.sleep(10)
        else:
            dp = None
        print('查询已导出+++')
        print('*' * 50)

    def _SearchGoods(self,wayBillNumber, proxy_id, proxy_handle):
        # #1、构建url
        # url = "http://query2.e-can.com.tw/ECAN_APP/search.shtm"   #url为机器人的webhook
        # #2、构建一下请求头部
        # r_header = {"Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        #             "Accept-Encoding": "gzip, deflate",
        #             "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        #             'Host': 'query2.e-can.com.tw',
        #             'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.42'
        #             }
        # #3、构建请求数据
        # data = {'txtMainID': wayBillNumber, 'B1': '查詢'}
        # if proxy_handle == '代理服务器':
        #     proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
        #     req = self.session.get(url=url, headers=r_header, data=data, proxies=proxies, allow_redirects=True)
        # else:
        #     req = self.session.get(url=url, headers=r_header, data=data)
        #     # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=True)
        # req.encoding = "utf-8"                      # 新增编码格式
        # print('----------获取验证值成功-------------')
        # print(req)
        # print(req.headers)
        # print(req.text)

        #1、构建url
        url = "http://query2.e-can.com.tw/ECAN_APP/DS_LINK.asp"   #url为机器人的webhook
        #2、构建一下请求头部
        r_header = {'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
                    'Accept-Encoding': 'gzip, deflate',
                    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
                    'Cache-Control': 'max-age=0',
                    'Connection': 'keep-alive',
                    'Content-Length': '44',
                    'Content-Type': 'application/x-www-form-urlencoded',
                    'Cookie': '__utma=73744262.1058373605.1685364822.1685364822.1685364822.1; __utmz=73744262.1685364822.1.1.utmcsr=query2.e-can.com.tw|utmccn=(referral)|utmcmd=referral|utmcct=/; _ga=GA1.3.1058373605.1685364822',
                    'Host': 'query2.e-can.com.tw',
                    'Origin': 'http://query2.e-can.com.tw',
                    'Referer': 'http://query2.e-can.com.tw/ECAN_APP/search.shtm',
                    'Upgrade-Insecure-Requests': '1',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.42'
                    }
        #3、构建请求数据
        data = {'txtMainID': wayBillNumber, 'B1': '查詢'}
        USER_AGENTS = ['192.168.13.89:37466', '192.168.13.89:37467']
        proxy_id = USER_AGENTS[randint(0, len(USER_AGENTS) - 1)]
        print(proxy_id)
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies, allow_redirects=False)
        else:
            req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        req.encoding = "utf-8"                      # 新增编码格式
        # print('----------获取验证值成功-------------')
        print(req)
        print(req.headers)
        # print(req.text)
        print('----------数据获取返回成功-----------')
        rq = BeautifulSoup(req.text, 'lxml')
        ordersDict = []
        order_number = str(rq).split(r'出貨單號：')[1].split('</h2>')[0]              # 订单编号
        waybill_number = ""
        k = 0
        k2 = 0
        k3 = 0
        rq = rq.find_all('tr')
        print('*' * 50)
        for index, val in enumerate(rq):
            result = {}
            # print(index)
            # print(val)
            if 'colspan="4"' in str(val):       # 查询到货态（一）
                waybill_number = str(val).split(r'單號：')[1].split('-00')[0]
            elif 'class="date">' in str(val):       # 查询到货态（一）
                rq_val = val.find_all('td')
                result['序号'] = index
                result['orderNumber'] = order_number
                result['wayBillNumber'] = waybill_number
                result['track_date'] = str(rq_val[0]).split(r'class="date">')[1].split('</span')[0]
                result['出货时间'] = ""
                result['上线时间'] = ""
                result['保管时间'] = ""
                result['完成时间'] = ""
                result['track_info'] = str(rq_val[1]).split(r'<td>')[1].split('</td>')[0]   # 物流轨迹
                if '取件完成' in result['track_info']:
                    if k < 1:
                        result['出货时间'] = result['track_date']
                        k = k + 1
                if '轉運作業中' in result['track_info']:
                    if k2 < 1:
                        result['上线时间'] = result['track_date']
                        k2 = k2 + 1
                if '異常狀況' in result['track_info']:
                    if k3 < 1:
                        result['保管时间'] = result['track_date']
                        k3 = k3 + 1
                result['track_status'] = ""
                result['负责营业所'] = str(rq_val[3]).split(r'<td>')[1].split('</td>')[0]
                result['轨迹备注'] = str(rq_val[2]).split(r'<td>')[1].split('</td>')[0]
                result['序号'] = index - 1
                result['便利店'] = ""
            ordersDict.append(result)
        data = pd.json_normalize(ordersDict)
        data.dropna(axis=0, how='any', inplace=True)
        data.sort_values(by=["wayBillNumber", "track_date"], inplace=True, ascending=[True, True])      # inplace: 原地修改; ascending：升序
        print(data)
        print('*' * 50)
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # data.to_excel('F:\\输出文件\\宅配通快递 {0} .xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        return data

    def _SearchGoodsT(self,wayBillNumber, proxy_id, proxy_handle):
        #1、构建url
        url = "http://query2.e-can.com.tw/self_link/id_link_c.asp?txtMainid=" + wayBillNumber  #url为机器人的webhook
        #2、构建一下请求头部
        r_header = {"Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
                    "Accept-Encoding": "gzip, deflate",
                    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Charset": "UTF-8",
                    'Host': 'query2.e-can.com.tw',
                    'Origin': 'http://query2.e-can.com.tw',
                    'Referer': 'http://query2.e-can.com.tw/%E5%A4%9A%E7%AD%86%E6%9F%A5%E4%BB%B6_oo4o.asp',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36 Edg/113.0.1774.42'
                    }
        #3、构建请求数据
        data = {'txtMainID': wayBillNumber,
                'B1': '查詢'}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies, allow_redirects=False)
        else:
            req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
            # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=True)
        req.encoding = "utf-8"                      # 新增编码格式
        # print('----------获取验证值成功-------------')
        print(req)
        print(req.text)
        print('----------数据获取返回成功-----------')
        rq = BeautifulSoup(req.text, 'lxml')
        ordersDict = []
        order_number = str(rq).split(r'出貨單號：')[1].split('</h2>')[0]              # 订单编号
        waybill_number = ""
        k = 0
        k2 = 0
        k3 = 0
        rq = rq.find_all('tr')
        print('*' * 50)
        for index, val in enumerate(rq):
            result = {}
            # print(index)
            # print(val)
            if 'colspan="4"' in str(val):       # 查询到货态（一）
                waybill_number = str(val).split(r'單號：')[1].split('-00')[0]
            elif 'class="date">' in str(val):       # 查询到货态（一）
                rq_val = val.find_all('td')
                result['序号'] = index
                result['orderNumber'] = order_number
                result['wayBillNumber'] = waybill_number
                result['track_date'] = str(rq_val[0]).split(r'class="date">')[1].split('</span')[0]
                result['出货时间'] = ""
                result['上线时间'] = ""
                result['保管时间'] = ""
                result['完成时间'] = ""
                result['track_info'] = str(rq_val[1]).split(r'<td>')[1].split('</td>')[0]   # 物流轨迹
                if '取件完成' in result['track_info']:
                    if k < 1:
                        result['出货时间'] = result['track_date']
                        k = k + 1
                if '轉運作業中' in result['track_info']:
                    if k2 < 1:
                        result['上线时间'] = result['track_date']
                        k2 = k2 + 1
                if '異常狀況' in result['track_info']:
                    if k3 < 1:
                        result['保管时间'] = result['track_date']
                        k3 = k3 + 1
                result['track_status'] = ""
                result['负责营业所'] = str(rq_val[3]).split(r'<td>')[1].split('</td>')[0]
                result['轨迹备注'] = str(rq_val[2]).split(r'<td>')[1].split('</td>')[0]
                result['序号'] = index - 1
                result['便利店'] = ""
            ordersDict.append(result)
        data = pd.json_normalize(ordersDict)
        data.dropna(axis=0, how='any', inplace=True)
        data.sort_values(by=["wayBillNumber", "track_date"], inplace=True, ascending=[True, True])      # inplace: 原地修改; ascending：升序
        print(data)
        print('*' * 50)
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        data.to_excel('F:\\输出文件\\宅配通快递 {0} .xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        return data

if __name__ == '__main__':
    m = QueryTwo('+86-18538110674', 'qyz04163510')
    start: datetime = datetime.datetime.now()
    print(datetime.datetime.now())
    '''
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    '''
    # m.readFormHost()
    proxy_handle = '代理服务器0'
    proxy_id = '192.168.13.89:37468'  # 输入代理服务器节点和端口
    # m.readFormHost(proxy_id, proxy_handle)
    m.SearchGoods_write(proxy_id, proxy_handle)
    # m._SearchGoods('377194099656', proxy_id, proxy_handle)
    # m._SearchGoodsT('377194099656', proxy_id, proxy_handle)

    print('查询耗时：', datetime.datetime.now() - start)