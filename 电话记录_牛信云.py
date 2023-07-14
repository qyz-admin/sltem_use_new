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
        # self.e = EmailControl()
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
            dp = df.append(dlist, ignore_index=True)
            dp.dropna(axis=0, how='any', inplace=True)
        else:
            dp = None
        print(dp)
        dp.to_excel('G:\\输出文件\\新竹17track-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
        print('查询已导出+++')
        print('*' * 50)
    def _SearchGoods(self,wayBillNumber):
        #0、获取验证码
        code = ''
        for i in range(4):
            code += str(random.randint(0, 9))
        #1、构建url
        url = "https://www.17track.net/zh-cn"   #url为机器人的webhook
        #2、构建一下请求头部
        r_header = {"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
                    'accept-encoding': 'gzip, deflate, br',
                    'accept-language': 'zh-CN,zh;q=0.9',
                    'Host': 't.17track.net',
                    'Origin': 'https://t.17track.net',
                    'Referer': 'https://features.17track.net/zh-cn/carriersettlein',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
                    }
        req = self.session.get(url=url, headers=r_header)
        print(req)
        print(req.headers)
        # print(req.text)
        print(88)

        #1、构建url
        url = "https://t.17track.net/zh-cn?v=2"   #url为机器人的webhook
        #2、构建一下请求头部
        r_header = {"accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
                    'accept-encoding': 'gzip, deflate, br',
                    'accept-language': 'zh-CN,zh;q=0.9',
                    'Referer': 'https://t.17track.net/zh-cn',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
                    }
        req = self.session.get(url=url, headers=r_header)
        print(req)
        print(req.headers)
        # print(req.text)
        print(88088)



        #1、构建url
        url = "https://t.17track.net/track/restapi"   #url为机器人的webhook
        #2、构建一下请求头部
        r_header = {"Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                    "accept": "application/json, text/javascript, */*; q=0.01",
                    'accept-encoding': 'gzip, deflate, br',
                    'accept-language': 'zh-CN,zh;q=0.9',
                    'Host': 't.17track.net',
                    'Origin': 'https://t.17track.net',
                    'Referer': 'https://t.17track.net/zh-cn?v=2',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
                    }
        #3、构建请求数据
        data = {'data': [{'num': '8555614292',
                          'fc': '190466',
                          'sc': '0'
                 }],
                'guid': 'ac2d6a83922746a2829e944e2b00190f',
                'timeZoneOffset': '-480'
                }
        print(data)
        req = self.session.post(url=url, headers=r_header, data=data)
        print(req)
        print(req.headers)
        soup = BeautifulSoup(req.text, 'lxml')      # 创建 beautifulsoup 对象
        print(req.next)
        print(soup)
        print('----------获取验证值成功-------------')

        url = "https://t.17track.net/track/restapi"   #url为机器人的webhook
        r_header = {"Content-Type": "application/x-www-form-urlencoded",
                    "Charset": "UTF-8",
                    'Host': 'www.hct.com.tw',
                    'Origin': 'https://www.hct.com.tw',
                    'Referer': 'https://www.hct.com.tw/search/searchgoods_n.aspx',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
                    }
        # data = {'no': no,
        #         'chk': chk
        #         }
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
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
                data = pd.DataFrame([[res, res2, '', '', online_time]], columns=['查货号码', '查货时间', '轨迹时间', '轨迹内容', '上线时间'])
                # data.dropna(axis=0, how='any', inplace=True)
                data.dropna(axis=0, subset=["查货号码"], inplace=True)
                data.sort_values(by="轨迹时间", inplace=True, ascending=True)  # inplace: 原地修改; ascending：升序
                # print(data)
                return data
            if "ctl00_ContentFrame_lblInvoiceNo" in str(val):       # 查询到货态（一）
                result['序号'] = index
                result['查货号码'] = str(val).split(r'ctl00_ContentFrame_lblInvoiceNo">')[1].split('</')[0]
                result['查货时间'] = str(val).split('時間：')[1].split('</')[0]
                res = result['查货号码'].strip()
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
                result['查货号码'] = res
                result['查货时间'] = res2
                result['轨迹时间'] = L_time
                result['轨迹内容'] = L_cls.replace("\n\n", "").strip()
                result['上线时间'] = online_time
            ordersDict.append(result)
        # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        data = pd.json_normalize(ordersDict)
        # print(data)
        # data.dropna(axis=0, how='any', inplace=True)
        data.sort_values(by=["查货号码", "轨迹时间"], inplace=True, ascending=[True, True])
        # data.sort_values(by="轨迹时间", inplace=True, ascending=True)  # inplace: 原地修改; ascending：升序
        # print(data)
        # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # data.to_excel('G:\\输出文件\\新竹快递 {0} .xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        return data
    def TW_SearchGoods(self):
        url = "https://emap.pcsc.com.tw/EMapSDK.aspx"
        #2、构建一下请求头部
        r_header = {"Content-Type": "application/x-www-form-urlencoded",
                    "Charset": "UTF-8",
                    'Host': 'emap.pcsc.com.tw',
                    'Origin': 'https://emap.pcsc.com.tw',
                    'Referer': 'https://emap.pcsc.com.tw/emap.aspx',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
                    }
        #3、构建请求数据
        data = {'commandid': 'Search0007',
                'x1': 116894531,
                'y1': 22055096,
                'x2': 124837646,
                'y2': 26588527
                }
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        print(req)
        print("*" * 50)
        print(req.text)
        soup = BeautifulSoup(req.text, 'lxml')      # 创建 beautifulsoup 对象

        print("*" * 50)
        print(soup)
        no = soup.input.get('value')
        chk = soup.input.next_sibling.get('value')
        # print(no)

        # return data

    # 、电话记录查询_牛新云
    def getNewSipCdrList(self, begin, end, requestid, usertoken):
        match = {886: '台湾', 852: '香港'}
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        dlist = []
        df = pd.DataFrame([])
        for code in [886, 852]:
            # dlist = []
            # df = pd.DataFrame([])
            for i in range((end - begin).days):  # 按天循环获取订单状态
                day = begin + datetime.timedelta(days=i)
                day_time = str(day)
                print('正在获取 ' + match[code] + '： ' + day_time + ' 号电话记录…………')
                data = self._getNewSipCdrList(day_time, code, requestid, usertoken)
                dlist.append(data)
        dp = df.append(dlist, ignore_index=True)
        dp = dp[['线路', '日期', 'name', 'connect_size', 'no_connect_size', 'total_size',  'total_duration']]
        dp.columns = ['线路', '日期', '话机昵称', '接通总数', '未接通总数', '呼叫总数',  '总时长(秒)']
        print(dp)
        dp.to_excel('F:\\输出文件\\电话记录-牛信云 {0}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        print('写入成功......')

    def _getNewSipCdrList(self, day_time, code, requestid, usertoken):
        match = {886: '台湾', 852: '香港'}
        url = "https://backend.nxcloud.com/user/newSipTrunk/getDayPhoneTotalInfo"
        #2、构建一下请求头部
        r_header = {"Content-Type": "application/x-www-form-urlencoded",
                    "Charset": "UTF-8",
                    "requestid": requestid,
                    'usertoken': usertoken,
                    'Origin': 'https://www.nxcloud.com',
                    'Referer': 'https://www.nxcloud.com/',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36'
                    }
        #3、构建请求数据
        data = {'startDate': day_time + ' 00:00:00',
                'endDate': day_time + ' 23:59:59',
                'phone': '',
                'effective_called_number': '',
                'duration_flag': '',
                'sip_group_id': '',
                'sip_mobile_id': '',
                'code': code,
                'pageSize': 10,
                'type': '',
                'page': 1
                }
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print("*" * 50)
        req = json.loads(req.text)
        ordersdict = []
        try:
            for result in req['info']['rows']:
                result['线路'] = match[code]
                result['日期'] = day_time
                ordersdict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        df = pd.json_normalize(ordersdict)
        # print(df)
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return df



        # print(no)

        # return data

if __name__ == '__main__':
    m = QueryTwo('+86-18538110674', 'qyz04163510')
    start: datetime = datetime.datetime.now()
    '''
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    '''
    # m.readFormHost()
    # m._SearchGoods('8555379070')
    # m._SearchGoods('7532082106')
    

    begin = datetime.date(2023, 7, 3)  # 单点更新
    end = datetime.date(2023, 7, 9)
    requestid = "e99afb02-65f9-4f77-bfba-66d7bd13ad2f"
    usertoken = "6cdeb266-5ccf-40e3-8334-a8f7ea9b4e8d"
    m.getNewSipCdrList(begin, end, requestid, usertoken)



    print('查询耗时：', datetime.datetime.now() - start)