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
from dateutil.relativedelta import relativedelta
from threading import Thread #  使用 threading 模块创建线程
import pandas.io.formats.excel

from sqlalchemy import create_engine
from settings import Settings
from settings_sso import Settings_sso
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色


# -*- coding:utf-8 -*-
class QueryTwo(Settings, Settings_sso):
    def __init__(self, userMobile, password, login_TmpCode, handle):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self._online()
        # self.sso_online_Two()
        # self.sso__online_handle(login_TmpCode)
        # self.sso__online_auto()
        if handle == '手动':
            self.sso__online_handle(login_TmpCode)
        else:
            self.sso__online_auto()

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
    def readFormHost(self, isReal):
        start = datetime.datetime.now()
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath, isReal)
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath, isReal):
        fileType = os.path.splitext(filePath)[1]
        app = xlwings.App(visible=False, add_book=False)
        app.display_alerts = False
        if 'xls' in fileType:
            wb = app.books.open(filePath, update_links=False, read_only=True)
            for sht in wb.sheets:
                if sht.api.Visible == -1:
                    try:
                        tem = ''
                        db = None
                        db = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                        # columns_value = list(db.columns)                             # 获取数据的标题名，转为列表
                        # for column_val in columns_value:
                        #     if '运单编号' == column_val:
                        #         tem = column_val
                        #     elif '运单号' == column_val:
                        #         tem = column_val
                        if '运单编号' in db.columns:
                            tem = '运单编号'
                        elif '查件单号' in db.columns:
                            tem = '查件单号'
                        elif '运单号' in db.columns:
                            tem = '运单号'
                        db = db[[tem]]
                        db.dropna(axis=0, how='any', inplace=True)                  # 空值（缺失值），将空值所在的行/列删除后
                    except Exception as e:
                        print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                    if db is not None and len(db) > 0:
                        print(db)
                        print('++++正在获取：' + sht.name + ' 表；共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                        # 将获取到的运单号 查询轨迹
                        self.Search_online(db, isReal, tem)
                    else:
                        print('----------数据为空,查询失败：' + sht.name)
                else:
                    print('----不需查询：' + sht.name)
            wb.close()
        app.quit()

    #  查询运单轨迹-按订单查询（一）
    def Search_online(self, db, isReal, tem):
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        orderId = list(db[tem])
        print(orderId)
        max_count = len(orderId)                 # 使用len()获取列表的长度，上节学的
        if max_count > 0:
            df = pd.DataFrame([])                # 创建空的dataframe数据框
            dlist = []
            for ord in orderId:
                print(ord)
                data = self._order_online(ord, isReal)
                # print(data)
                if data is not None and len(data) > 0:
                    dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
            # dp.dropna(axis=0, how='any', inplace=True)
        else:
            dp = None
        print(dp)
        dp = dp[['orderNumber', 'wayBillNumber', 'track_date', '出货时间', '上线时间', '保管时间', '完成时间', 'track_info', 'track_status']]
        dp.columns = ['订单编号', '运单号', '物流轨迹时间', '出货时间', '上线时间', '保管时间', '完成时间', '物流轨迹', '轨迹代码']
        dp.to_excel('G:\\输出文件\\运单轨迹-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
        print('查询已导出+++')
        print('*' * 50)

    #  查询运单轨迹-按时间查询（二）
    def order_online(self, timeStart, timeEnd, isReal):  # 进入运单轨迹界面
        # print('正在获取需要订单信息......')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        start = datetime.datetime.now()
        begin = datetime.datetime.strptime(timeStart, '%Y-%m-%d')
        end = datetime.datetime.strptime(timeEnd, '%Y-%m-%d')
        for i in range((end - begin).days):
            day = begin + datetime.timedelta(days=i + 1)
            day = day.strftime('%Y-%m-%d')
            print('正在获取 ' + day + ' 号订单信息…………')
            sql = '''SELECT id,`运单编号`  FROM gat_order_list sl WHERE sl.`日期` = '{0}' AND sl.运单编号 IS NOT NULL;'''.format(day)
            # sql = '''SELECT id,`运单编号`  FROM gat_order_list sl WHERE sl.`下单时间` BETWEEN  '2022-02-03 09:00:00' AND '2022-02-03 09:30:00' AND sl.运单编号 IS NOT NULL;'''.format(day)
            ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
            if ordersDict.empty:
                print('无需要更新订单信息！！！')
                return
            # print(ordersDict['运单编号'][0])
            orderId = list(ordersDict['运单编号'])
            max_count = len(orderId)            # 使用len()获取列表的长度，上节学的
            if max_count > 0:
                df = pd.DataFrame([])           # 创建空的dataframe数据框
                dlist = []
                for ord in orderId:
                    print(ord)
                    data = self._order_online(ord, isReal)
                    if data is not None and len(data) > 0:
                        dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
                dp.to_excel('G:\\输出文件\\运单轨迹 {0} .xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        print('++++++查询成功+++++++')
        print('查询耗时：', datetime.datetime.now() - start)
        print('*' * 50)

    #  查询运单轨迹-按订单查询（一）
    def _order_online(self, ord, isReal):  # 进入订单检索界面
        print('+++正在实时_搜索轨迹信息中')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getLogisticsTrace'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/logisticsTrajectory'}
        data = {'numbers': ord,
                'searchType': 1,
                'isReal': isReal
                }
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        if req['data']['list'] == []:
            data = self._order_online_data(ord, 0)
            return data
        else:
            ordersDict = []
            try:
                if req['data']['list'][0]['list'] == []:
                    # print(req['data']['list'])
                    return None
                else:
                    for result in req['data']['list']:
                        for res in result['list']:
                            res['出货时间'] = ''
                            res['上线时间'] = ''
                            res['完成时间'] = ''
                            res['保管时间'] = ''
                            res['orderNumber'] = result['order_number']
                            res['wayBillNumber'] = result['track_no']
                            if '.' in res['track_date']:
                                res['track_date'] = res['track_date'].split('.')[0]
                            else:
                                res['track_date'] = res['track_date']

                            if '二次出貨' in res['track_info'] or '二次出貨貼標' in res['track_info'] or '出貨理貨中' in res['track_info'] or '首發出貨貼標' in res['track_info'] or '已发货' in res['track_info']:
                                res['出货时间'] = res['track_date']
                            elif '已核重,集运仓' in res['track_info'] or '已核重-集运仓' in res['track_info']:
                                res['出货时间'] = res['track_date']

                            if '貨件已抵達 土城營業所，貨件整理中' in res['track_info'] or '貨件已抵達 桃園營業所 ，貨件整理中' in res['track_info'] or '貨件已抵達土城營業所，貨件整理中' in res['track_info'] or '貨件已抵達桃園營業所 ，貨件整理中' in res['track_info'] or '顺丰速运 已收取快件' in res['track_info']:
                                res['上线时间'] = res['track_date']
                            elif '貨件已抵達 大園營業所 ，貨件整理中' in res['track_info'] or '貨件已抵達 大園營業所，貨件整理中' in res['track_info']:
                                res['上线时间'] = res['track_date']
                            elif '營業所發出，前往配送站途中' in res['track_info']:
                                res['上线时间'] = res['track_date']
                            elif '營業所，分貨中' in res['track_info']:
                                res['上线时间'] = res['track_date']

                            if ('送達客戶不在' in res['track_info'] or '貨件拒收' in res['track_info'] or '與客戶另約時間' in res['track_info'] or '收貨人' in res['track_info']) and '保管中' in res['track_info'] or '營業所送至' in res['track_info'] or '集配站送至' in res['track_info']:
                                res['保管时间'] = res['track_date']
                            elif ('貨件由' in res['track_info'] and '保管中' in res['track_info']) or '收件地址為公司行號，本日休假' in res['track_info'] or (('收貨人' in res['track_info'] or '該地址' in res['track_info']) and '營業所處理中' in res['track_info']):
                                res['保管时间'] = res['track_date']

                            if '送達。貨物件數共' in res['track_info'] or '貨件已退回，退貨號碼' in res['track_info']:
                                res['完成时间'] = res['track_date']

                            # print(res)track_info
                            ordersDict.append(res)
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            data = pd.json_normalize(ordersDict)
            data.sort_values(by=["orderNumber", "track_date"], inplace=True, ascending=[True, True])  # inplace: 原地修改; ascending：升序 （是否升序排序，默认为true，降序则为false。如果是列表，则需和by指定的列表数量相同，指明每一列的排序方式）
            print('++++++本次获取成功+++++++')
            # print('*' * 50)
            return data
    # 物流轨迹数据库
    def _order_online_data(self, ord, isReal):  # 进入订单检索界面
        print('+++正在_搜索轨迹信息中')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getLogisticsTrace'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/logisticsTrajectory'}
        data = {'numbers': ord,
                'searchType': 1,
                'isReal': isReal
                }
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        if req['data']['list'] == []:
            return None
        else:
            ordersDict = []
            try:
                if req['data']['list'][0]['list'] == []:
                    # print(req['data']['list'])
                    return None
                else:
                    for result in req['data']['list']:
                        for res in result['list']:
                            # res['出货上线时间'] = ''
                            # res['orderNumber'] = result['order_number']
                            # res['wayBillNumber'] = result['track_no']
                            # if '.' in res['track_date']:
                            #     res['track_date'] = res['track_date'].split('.')[0]
                            # else:
                            #     res['track_date'] = res['track_date']
                            # if '已核重-集运仓发货' in res['track_info'] or '顺丰速运 已收取快件' in res['track_info'] or '貨件整理中' in res['track_info'] or '二次出貨貼標' in res['track_info']:
                            #     res['出货上线时间'] = res['track_date']

                            res['出货时间'] = ''
                            res['上线时间'] = ''
                            res['完成时间'] = ''
                            res['保管时间'] = ''
                            res['orderNumber'] = result['order_number']
                            res['wayBillNumber'] = result['track_no']
                            if '.' in res['track_date']:
                                res['track_date'] = res['track_date'].split('.')[0]
                            else:
                                res['track_date'] = res['track_date']

                            if '二次出貨' in res['track_info'] or '二次出貨貼標' in res['track_info'] or '出貨理貨中' in res['track_info'] or '首發出貨貼標' in res['track_info'] or '已发货' in res['track_info']:
                                res['出货时间'] = res['track_date']
                            elif '已核重,集运仓' in res['track_info'] or '已核重-集运仓' in res['track_info']:
                                res['出货时间'] = res['track_date']

                            if '貨件已抵達 土城營業所，貨件整理中' in res['track_info'] or '貨件已抵達 桃園營業所 ，貨件整理中' in res['track_info'] or '貨件已抵達土城營業所，貨件整理中' in res['track_info'] or '貨件已抵達桃園營業所 ，貨件整理中' in res['track_info'] or '顺丰速运 已收取快件' in res['track_info']:
                                res['上线时间'] = res['track_date']
                            elif '貨件已抵達 大園營業所 ，貨件整理中' in res['track_info'] or '貨件已抵達 大園營業所，貨件整理中' in res['track_info']:
                                res['上线时间'] = res['track_date']
                            elif '營業所發出，前往配送站途中' in res['track_info']:
                                res['上线时间'] = res['track_date']
                            elif '營業所，分貨中' in res['track_info']:
                                res['上线时间'] = res['track_date']

                            if ('送達客戶不在' in res['track_info'] or '貨件拒收' in res['track_info'] or '與客戶另約時間' in res['track_info'] or '收貨人' in res['track_info']) and '保管中' in res['track_info'] or '營業所送至' in res['track_info'] or '集配站送至' in res['track_info']:
                                res['保管时间'] = res['track_date']
                            elif ('貨件由' in res['track_info'] and '保管中' in res['track_info']) or '收件地址為公司行號，本日休假' in res['track_info'] or (('收貨人' in res['track_info'] or '該地址' in res['track_info']) and '營業所處理中' in res['track_info']):
                                res['保管时间'] = res['track_date']

                            if '送達。貨物件數共' in res['track_info'] or '貨件已退回，退貨號碼' in res['track_info']:
                                res['完成时间'] = res['track_date']

                            ordersDict.append(res)
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            data = pd.json_normalize(ordersDict)
            data.sort_values(by=["orderNumber", "track_date"], inplace=True, ascending=[True, True])  # inplace: 原地修改; ascending：升序 （是否升序排序，默认为true，降序则为false。如果是列表，则需和by指定的列表数量相同，指明每一列的排序方式）
            print('++++++本次获取成功+++++++')
            # print('*' * 50)
            return data





    #  查询运单轨迹-上线及派送（三）
    def order_bind_status(self, timeStart, timeEnd):  # 进入运单轨迹界面
        # print('正在获取需要订单信息......')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        start = datetime.datetime.now()
        begin = datetime.datetime.strptime(timeStart, '%Y-%m-%d')
        end = datetime.datetime.strptime(timeEnd, '%Y-%m-%d')
        for i in range((end - begin).days):
            day = begin + datetime.timedelta(days=i + 1)
            day = day.strftime('%Y-%m-%d')
            print('正在获取 ' + day + ' 号订单信息…………')
            # sql = '''SELECT id,`运单编号`  FROM gat_order_list sl WHERE sl.`日期` = '{0}' AND sl.运单编号 IS NOT NULL;'''.format(day)
            sql = '''SELECT id,`运单编号`  FROM gat_order_list sl WHERE sl.`下单时间` BETWEEN  '2022-02-03 00:00:00' AND '2022-02-03 23:59:59' AND sl.运单编号 IS NOT NULL;'''.format(day)
            ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
            if ordersDict.empty:
                print('无需要更新订单信息！！！')
                return
            # print(ordersDict['运单编号'][0])
            orderId = list(ordersDict['运单编号'])
            max_count = len(orderId)            # 使用len()获取列表的长度，上节学的
            if max_count > 0:
                df = pd.DataFrame([])           # 创建空的dataframe数据框
                dlist = []
                for ord in orderId:
                    print(ord)
                    data = self._order_bind_status(ord)
                    if data is not None and len(data) > 0:
                        dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
                dp.to_excel('G:\\输出文件\\运单轨迹 {0} .xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        print('++++++查询成功+++++++')
        print('查询耗时：', datetime.datetime.now() - start)
        print('*' * 50)
    def _order_bind_status(self, ord):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getLogisticsTrace'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/logisticsTrajectory'}
        data = {'numbers': ord,
                'searchType': 1,
                'isReal': 1
                }
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        ordersDict = []
        val = {}
        try:
            if req['data']['list'] == []:
                print(req['data']['list'])
                return None
            else:
                for result in req['data']['list']:
                    # print(result)
                    # print(999)
                    k = 0
                    t = 0
                    g = 0
                    j = 0
                    h = 0
                    val['订单编号'] = result['order_number']
                    val['运单编号'] = result['track_no']
                    for res in result['list']:
                        # print(0)
                        # print(res)
                        if '快件到达 【香港青衣中轉場】' in res['track_info']:         # 立邦顺丰
                            while t == 0:
                                val['上线时间'] = res['track_date']
                                t = t + 1
                            for res in result['list']:
                                # print(11)
                                # print(res)
                                if '正在派送途中' in res['track_info']:
                                    g = g + 1
                                    val[str(g) +'派'] = res['track_date']
                                    # print('g:' + str(g) +' 派')

                        elif '項目營業部' in res['track_info']:                       # 天马顺丰
                            val['上线时间'] = res['track_date']
                            for res in result['list']:
                                # print(22)
                                # print(res)
                                if '正在派件...' in res['track_info'] or '快件派送不成功' in res['track_info'] or '快件已签收' in res['track_info'] or '快件代签收' in res['track_info']:
                                    j = j + 1
                                    val[str(j) +'派'] = res['track_date']
                                    # print('j:' + str(j))

                        elif '貨件已抵達土城營業所，貨件整理中' in res['track_info'] or '貨件已抵達桃園營業所，貨件整理中' in res['track_info'] or '集荷' in res['track_info']:      # 新竹 貨件已抵達桃園營業所，貨件整理中
                            val['上线时间'] = res['track_date']
                            for res in result['list']:
                                # print(33)
                                # print(res)
                                if '配送中' in res['track_info']:
                                    k = k + 1
                                    val[str(k) +'派'] = res['track_date']
                                    # print('k:' + str(k))

                        elif '結轉物流中心' in res['track_info']:                         # 7-11 
                            val['上线时间'] = res['track_date']
                            for res in result['list']:
                                # print(44)
                                # print(res)
                                if '门市配达' in res['track_info']:
                                    h = h + 1
                                    val[str(h) +'派'] = res['track_date']
                                    # print('h:' + str(h))
                        elif '已安排航班' in res['track_info'] or '起飞' in res['track_info']:                         # 7-11 
                            val['上线时间'] = res['track_date']
                            for res in result['list']:
                                # print(55)
                                # print(res)
                                if '配送中' in res['track_info']:
                                    h = h + 1
                                    val[str(h) +'派'] = res['track_date']
                                    # print('h:' + str(h))
                        # if '貨件已抵達土城營業所，貨件整理中' in res['track_info'] or '貨件已抵達桃園營業所，貨件整理中' in res['track_info']:      # 新竹
                        #     val['上线时间'] = res['track_date']
                        # if '配送中' in res['track_info']:
                        #     k = k + 1
                        #     val[str(k) +'派'] = res['track_date']

                        # if '結轉物流中心' in res['track_info']:         # 7-11 
                        #     val['上线时间'] = res['track_date']
                        # if '门市配达' in res['track_info']:
                        #     val[str(k) +'派'] = res['track_date']

                        # if '項目營業部' in res['track_info']:         # 顺丰
                        #     val['上线时间'] = res['track_date']
                        # if '正在派件' in res['track_info'] or '快件已' in res['track_info'] or '快件代' in res['track_info'] or '快件派送' in res['track_info']:
                        #     val[str(k) +'派'] = res['track_date']

                        # if '快件到达 【香港青衣中轉場】' in res['track_info']:         # 立邦
                        #     while t == 0:
                        #         t= t + 1
                        #         val['上线时间'] = res['track_date']
                        # if '正在派送途中' in res['track_info']:
                        #     g = g + 1
                        #     val[str(g) +'派'] = res['track_date']
                    ordersDict.append(val)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print(data)
        # data.sort_values(by="track_date", inplace=True, ascending=True)  # inplace: 原地修改; ascending：升序
        # data['name'] = data['name'].str.strip()
        # data['cate_id'] = data['cate_id'].str.strip()
        # data['second_cate_id'] = data['second_cate_id'].str.strip()
        # data['third_cate_id'] = data['third_cate_id'].str.strip()
        # data = data[['id', 'name', 'cate_id', 'second_cate_id', 'third_cate_id', 'status', 'price', 'selectionName',
        #              'sellerCount', 'buyerName', 'saleCount', 'logisticsCost', 'lender', 'isGift', 'createTime',
        #              'categorys', 'image']]
        # data.columns = ['产品id', '产品名称', '一级分类', '二级分类', '三级分类', '产品状态', '价格(￥)', '选品人',
        #                 '供应商数', '采购人', '商品数', 'logisticsCost', '出借人', '特殊信息', '添加时间',
        #                 '产品分类', '产品图片']
        # data.to_excel('G:\\输出文件\\运单轨迹 {0} .xlsx'.format(rq), sheet_name='查询', index=False,engine='xlsxwriter')
        print('++++++本批次查询成功+++++++')
        # print('*' * 50)
        return data



if __name__ == '__main__':
    m = QueryTwo('+86-18538110674', 'qyz04163510.','0d74e19877643f1895752a1daf15325e','手动')
    start: datetime = datetime.datetime.now()
    match1 = {'gat': '港台', 'gat_order_list': '港台', 'slsc': '品牌'}
    '''
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    # 1、 正在按订单查询；2、正在按时间查询；--->>数据更新切换
    # isReal: 0 查询后台保存的运单轨迹； 1 查询物流的实时运单轨迹 
    '''
    isReal = 1
    select = 1
    if int(select) == 1:
        print("1-->>> 正在按运单号查询+++")
        m.readFormHost(isReal)       # 导入；，更新--->>数据更新切换
    elif int(select) == 2:
        print("2-->>> 正在按时间查询+++")
        m.order_online('2022-01-01', '2022-01-05', isReal)

    # m.order_bind_status('2022-01-01', '2022-01-02')

    # m._order_bind_status('7449201841')

    print('查询耗时：', datetime.datetime.now() - start)