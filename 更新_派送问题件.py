import pandas as pd
import os
import datetime
import time
from tqdm import tqdm

import xlwings
import xlsxwriter
import math
import requests
import json
import re
import sys
from queue import Queue
from dateutil.relativedelta import relativedelta
from threading import Thread #  使用 threading 模块创建线程
import pandas.io.formats.excel
import zhconv          # transform2_zh_hant：转为繁体;transform2_zh_hans：转为简体

from mysqlControl import MysqlControl
from sqlalchemy import create_engine
from settings import Settings
from settings_sso import Settings_sso
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色
from 更新_已下架_压单_头程导入提单号 import QueryTwoLower
from 查询_订单检索 import QueryOrder

# -*- coding:utf-8 -*-
class QueryTwo(Settings, Settings_sso):
    def __init__(self, userMobile, password, login_TmpCode,handle):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self._online()
        # self.sso_online_Two()
        # self.sso__online_handle(login_TmpCode)
        # # self.sso__online_auto()

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
        # self.my = MysqlControl()

    # 获取查询时间
    def readInfo(self, team):
        print('>>>>>>正式查询中<<<<<<')
        print('正在获取需要订单信息......')
        start = datetime.datetime.now()
        if team == '派送问题件_跟进表':
            last_time = datetime.datetime.now().strftime('%Y-%m') + '-01'
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')

        else:
            sql = '''SELECT DISTINCT 处理时间 FROM {0} d GROUP BY 处理时间 ORDER BY 处理时间 DESC'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['处理时间'][0])
            last_time = (rq + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        print('******************起止时间：' + team + last_time + ' - ' + now_time + ' ******************')
        return last_time, now_time


    # 查询更新（新后台的获取-派送问题件）
    def waybill_delivery_updata(self, timeStart, timeEnd):  # 进入订单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deliveryProblemPackage'}
        data = {'order_number': None, 'waybill_number': None, 'question_level': None, 'question_type': None, 'order_trace_id': None, 'ship_phone': None, 'page': 1, 'pageSize': 90,
                'addtime': None, 'question_time': None, 'trace_time': None, 'create_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'finishtime': None,
                'sale_id': None, 'product_id': None, 'logistics_id': None, 'area_id': None, 'currency_id': None, 'order_status': None, 'logistics_status': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)          # json类型数据转换为dict字典
        max_count = req['data']['count']    # 获取 请求订单量

        ordersDict = []
        if max_count != 0 and max_count != []:
            try:
                for result in req['data']['list']:                  # 添加新的字典键-值对，为下面的重新赋值
                    ordersDict.append(result.copy())
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            df = pd.json_normalize(ordersDict)
            print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
            print('*' * 50)
            if max_count > 90:
                in_count = math.ceil(max_count/90)
                dlist = []
                n = 1
                while n < in_count:  # 这里用到了一个while循环，穿越过来的
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                    data = self._waybill_delivery_updata(timeStart, timeEnd, n)
                    dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = df
            dp = dp[['order_number',  'currency', 'addtime', 'create_time', 'finishtime', 'lastQuestionName', 'orderStatus', 'logisticsStatus',
                     'reassignmentTypeName', 'logisticsName',  'questionAddtime', 'userName', 'traceName', 'traceTime', 'content']]
            dp.columns = ['订单编号', '币种', '下单时间', '创建时间', '完成时间', '派送问题', '订单状态', '物流状态',
                          '订单类型', '物流渠道',  '派送问题首次时间', '处理人', '处理记录', '处理时间', '备注']
            print('正在写入......')
            dp.to_sql('customer_up', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('G:\\输出文件\\派送问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 派送问题件_跟进表(订单编号,币种, 下单时间,完成时间,订单状态,物流状态,订单类型,物流渠道, 创建日期, 创建时间, 派送问题, 派送问题首次时间, 处理人, 处理记录, 处理时间,备注, 记录时间) 
                    SELECT 订单编号,币种, 下单时间,完成时间,订单状态,物流状态,订单类型,物流渠道, DATE_FORMAT(创建时间,'%Y-%m-%d') 创建日期, 创建时间, 派送问题, 派送问题首次时间, 处理人, 处理记录, IF(处理时间 = '',NULL,处理时间) 处理时间,备注,NOW() 记录时间 
                    FROM customer_up;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _waybill_delivery_updata(self, timeStart, timeEnd, n):  # 进入派送问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deliveryProblemPackage'}
        data = {'order_number': None, 'waybill_number': None, 'question_level': None, 'question_type': None, 'order_trace_id': None, 'ship_phone': None, 'page': n, 'pageSize': 90,
                'addtime': None, 'question_time': None, 'trace_time': None, 'create_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'finishtime': None,
                'sale_id': None, 'product_id': None, 'logistics_id': None, 'area_id': None, 'currency_id': None, 'order_status': None, 'logistics_status': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data






if __name__ == '__main__':
    start: datetime = datetime.datetime.now()
    '''
    # -----------------------------------------------自动获取 问题件 状态运行（一）-----------------------------------------
    # 1、 物流问题件；2、物流客诉件；3、物流问题件；4、全部；--->>数据更新切换
    '''
    select = 99
    if int(select) == 99:
        handle = '手0动'
        login_TmpCode = '7e00200b074b38be93d83578da27e666'
        m = QueryTwo('+86-18538110674', 'qyz35100416', login_TmpCode,handle)
        start: datetime = datetime.datetime.now()

        if int(select) == 1:
            timeStart, timeEnd = m.readInfo('物流问题件')

        elif int(select) == 99:
            timeStart, timeEnd = m.readInfo('派送问题件_跟进表')
            m.waybill_delivery_updata(timeStart, timeEnd)                        # 查询更新-派送问题件




    print('查询耗时：', datetime.datetime.now() - start)