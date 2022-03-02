import pandas as pd
import os
import datetime
import time
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
from 查询_已下架_压单 import QueryTwoLower


# -*- coding:utf-8 -*-
class QueryTwo(Settings, Settings_sso):
    def __init__(self, userMobile, password):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self._online()
        self.sso_online_Two()
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
        self.my = MysqlControl()
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
    #  登录后台中-停用
    def _online(self):  # 登录系统保持会话状态
        print('正在登录后台系统中......')
        # print('第一阶段获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171628645731266586976965831628645747396525G1E2B0816DEBF96BC4199761B6A1F3C0FCD91FB',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        req = req.json()
        # print(req)
        req_url = req['data']
        loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        # print(loginTmpCode)
        # print('+++已获取loginTmpCode值+++')

        time.sleep(1)
        # print('第二阶段请求-登录页面......')
        url = r'http://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.text)
        # print('+++请求登录页面url成功+++')

        time.sleep(1)
        # print('第三阶段请求-dingtalk服务器......')
        # print('（一）加载dingtalk_service跳转页面......')
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        gimp = req.headers['Location']
        # print('+++已获取跳转页面+++')
        time.sleep(1)
        # print('（二）请求dingtalk_service的cookie值......')
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print('+++已获取cookie值+++')

        time.sleep(1)
        # print('第四阶段页面-重定向跳转中......')
        # print('（一）加载chooselogin.html页面......')
        url = r'http://gsso.giikin.com/admin/login_by_dingtalk/chooselogin.html'
        data = {'user_id': 1343}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': gimp,
                    'Origin': 'http://gsso.giikin.com'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        index = req.headers['Location']
        # print('+++已获取gimp.giikin.com页面')
        time.sleep(1)
        # print('（二）加载gimp.giikin.com页面......')
        url = index
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': index}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index2 = req.headers['Location']
        # print('+++已获取index.html页面')

        time.sleep(1)
        # print('（三）加载index.html页面......')
        url = 'http://gimp.giikin.com/' + index2
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system = req.headers['Location']
        # print('+++已获取index.html?_system=18正式页面')

        time.sleep(1)
        # print('第五阶段正式页面-重定向跳转中......')
        # print('（一）加载index.html?_system页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system2 = req.headers['Location']
        # print('+++已获取index.html?_ticker=页面......')
        time.sleep(1)
        # print('（二）加载index.html?_ticker=页面......')
        url = index_system2
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)
        print('++++++已成功登录++++++')


    # 获取查询时间
    def readInfo(self, team):
        print('>>>>>>正式查询中<<<<<<')
        print('正在获取需要订单信息......')
        start = datetime.datetime.now()
        if team == '派送问题件':
            sql = '''SELECT DISTINCT 派送问题首次时间 FROM {0} d GROUP BY 派送问题首次时间 ORDER BY 派送问题首次时间 DESC'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['派送问题首次时间'][0])
            last_time = rq.strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now()).strftime('%Y-%m-%d')
        else:
            sql = '''SELECT DISTINCT 处理时间 FROM {0} d GROUP BY 处理时间 ORDER BY 处理时间 DESC'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['处理时间'][0])
            last_time = (rq + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        print('起止时间：' + team + last_time + ' - ' + now_time)
        return last_time, now_time

    # 查询更新（新后台的获取-物流问题件）
    def waybill_InfoQuery(self, timeStart, timeEnd):  # 进入订单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': None,
                'question_type': None, 'critical': None, 'read_status': None, 'operator_type': None, 'operator': None, 'create_time': None,
                'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_collection': None, 'logistics_status': None, 'user_id': None,
                'page': 1, 'pageSize': 90}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['traceRecord'] = zhconv.convert(result['traceRecord'], 'zh-hans')
                if ';' in result['traceRecord']:
                    trace_record = result['traceRecord'].split(";")
                    for record in trace_record:
                        result['deal_time'] = ''
                        result['result_reson'] = ''
                        result['result_info'] = ''

                        result['deal_time'] = record.split()[0]
                        rec = record.split("#处理结果：")[1]
                        # print(rec)
                        if len(rec.split()) > 2:
                            result['result_info'] = rec.split()[2]        # 客诉原因
                        if len(rec.split()) > 1:
                            result['result_reson'] = rec.split()[1]     # 处理内容
                        result['dealContent'] = rec.split()[0]            # 最新处理结果
                        rec_name = record.split("#处理结果：")[0]
                        if len(rec_name.split()) > 2:
                            if (rec_name.split())[2] != '' or (rec_name.split())[2] != []:
                                result['traceUserName'] = (rec_name.split())[2]
                        else:
                            result['traceUserName'] = ''
                        ordersDict.append(result.copy())
                else:
                    result['deal_time'] = ''
                    result['result_reson'] = ''
                    result['result_info'] = ''
                    if '拒收' in result['dealContent']:
                        if len(result['dealContent'].split()) > 2:
                            result['result_info'] = result['dealContent'].split()[2]
                        if len(result['dealContent'].split()) > 1:
                            result['result_reson'] = result['dealContent'].split()[1]
                        result['dealContent'] = result['dealContent'].split()[0]
                    if result['traceRecord'] != '' or result['traceRecord'] != []:
                        result['deal_time'] = result['traceRecord'].split()[0]
                    if result['traceUserName'] != '' or result['traceUserName'] != []:
                        result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                    result['dealContent'] = result['dealContent'].strip()
                    ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        df = pd.json_normalize(ordersDict)
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        print('*' * 50)
        if max_count != 0:
            if max_count > 90:
                in_count = math.ceil(max_count/90)
                dlist = []
                n = 1
                while n < in_count:  # 这里用到了一个while循环，穿越过来的
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                    data = self._waybillInfoQuery(timeStart, timeEnd, n)
                    dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = df
            dp = dp[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'dealStatus', 'dealContent',
                     'deal_time', 'result_reson', 'result_info', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus',
                     'gift_reissue_order_number', 'update_time']]
            dp.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '最新处理状态', '最新处理结果',
                          '处理时间', '拒收原因', '具体原因', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态', '赠品补发订单编号', '更新时间']
            dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞|张陈平', na=False))]
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('G:\\输出文件\\物流问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 物流问题件(处理时间,物流反馈时间,处理人,订单编号,处理结果, 拒收原因, 币种, 记录时间) 
                    SELECT 处理时间,导入时间 AS 物流反馈时间,处理人,订单编号,最新处理结果 AS 处理结果, 拒收原因, 币种, NOW() 记录时间 
                    FROM customer'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _waybillInfoQuery(self, timeStart, timeEnd, n):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': None,
                'question_type': None, 'critical': None, 'read_status': None, 'operator_type': None, 'operator': None, 'create_time': None,
                'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_collection': None, 'logistics_status': None, 'user_id': None,
                'page': n, 'pageSize': 90}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['traceRecord'] = zhconv.convert(result['traceRecord'], 'zh-hans')
                if ';' in result['traceRecord']:
                    trace_record = result['traceRecord'].split(";")
                    for record in trace_record:
                        result['deal_time'] = ''
                        result['result_reson'] = ''
                        result['result_info'] = ''

                        result['deal_time'] = record.split()[0]
                        rec = record.split("#处理结果：")[1]
                        # print(rec)
                        if len(rec.split()) > 2:
                            result['result_info'] = rec.split()[2]  # 客诉原因
                        if len(rec.split()) > 1:
                            result['result_reson'] = rec.split()[1]  # 处理内容
                        result['dealContent'] = rec.split()[0]  # 最新处理结果
                        rec_name = record.split("#处理结果：")[0]
                        if len(rec_name.split()) > 2:
                            if (rec_name.split())[2] != '' or (rec_name.split())[2] != []:
                                result['traceUserName'] = (rec_name.split())[2]
                        else:
                            result['traceUserName'] = ''
                        ordersDict.append(result.copy())
                else:
                    result['deal_time'] = ''
                    result['result_reson'] = ''
                    result['result_info'] = ''
                    if '拒收' in result['dealContent']:
                        if len(result['dealContent'].split()) > 2:
                            result['result_info'] = result['dealContent'].split()[2]
                        if len(result['dealContent'].split()) > 1:
                            result['result_reson'] = result['dealContent'].split()[1]
                        result['dealContent'] = result['dealContent'].split()[0]
                    if result['traceRecord'] != '' or result['traceRecord'] != []:
                        result['deal_time'] = result['traceRecord'].split()[0]
                    if result['traceUserName'] != '' or result['traceUserName'] != []:
                        result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                    result['dealContent'] = result['dealContent'].strip()
                    ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    # 查询更新（新后台的获取-派送问题件）
    def waybill_deliveryList(self, timeStart, timeEnd):  # 进入订单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deliveryProblemPackage'}
        data = {'order_number': None, 'waybill_number': None, 'question_level': None, 'question_type': None, 'order_trace_id': None, 'ship_phone': None,
                'page': 1, 'pageSize': 90, 'addtime': None, 'question_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'trace_time': None,
                'create_time': None, 'finishtime': None, 'sale_id': None, 'product_id': None, 'logistics_id': None, 'area_id': None, 'currency_id': None,
                'order_status': None, 'logistics_status': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        ordersDict = []
        if max_count != 0:
            try:
                for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
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
                    data = self._waybill_deliveryList(timeStart, timeEnd, n)
                    dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = df
            dp = dp[['order_number',  'currency', 'addtime', 'create_time', 'finishtime', 'lastQuestionName', 'orderStatus', 'logisticsStatus',
                     'reassignmentTypeName', 'logisticsName',  'questionAddtime', 'userName', 'traceName', 'traceTime', 'content']]
            dp.columns = ['订单编号', '币种', '下单时间', '创建时间', '完成时间', '派送问题', '订单状态', '物流状态',
                          '订单类型', '物流渠道',  '派送问题首次时间', '处理人', '处理记录', '处理时间', '备注']
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('G:\\输出文件\\派送问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 派送问题件(订单编号,创建时间, 派送问题, 派送问题首次时间, 处理人, 处理记录, 处理时间,备注, 记录时间) 
                    SELECT 订单编号,创建时间, 派送问题, 派送问题首次时间, 处理人, 处理记录, IF(处理时间 = '',NULL,处理时间) 处理时间,备注,NOW() 记录时间 
                    FROM customer'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _waybill_deliveryList(self, timeStart, timeEnd, n):  # 进入派送问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deliveryProblemPackage'}
        data = {'order_number': None, 'waybill_number': None, 'question_level': None, 'question_type': None, 'order_trace_id': None, 'ship_phone': None,
                'page': n, 'pageSize': 90, 'addtime': None, 'question_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'trace_time': None,
                'create_time': None, 'finishtime': None, 'sale_id': None, 'product_id': None, 'logistics_id': None, 'area_id': None, 'currency_id': None,
                'order_status': None, 'logistics_status': None}
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



    # 查询更新（新后台的获取-物流客诉件）
    def waybill_Query(self, timeStart, timeEnd):  # 进入物流客诉件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.orderCustomerComplaint&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'order_trace_id': None, 'question_type': None, 'critical': None, 'read_status': None,
                'operator_type': None, 'operator': None, 'create_time': None, 'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_gift_reissue': None,
                'is_collection': None, 'logistics_status': None, 'user_id': None, 'page': 1, 'pageSize': 90}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        # print(req)
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['traceRecord'] = zhconv.convert(result['traceRecord'], 'zh-hans')
                if ';' in result['traceRecord']:
                    trace_record = result['traceRecord'].split(";")
                    for record in trace_record:
                        result['deal_time'] = ''
                        result['result_reson'] = ''
                        result['result_info'] = ''
                        result['result_content'] = ''
                        result['deal_time'] = record.split()[0]
                        rec = record.split("#处理结果：")[1]
                        # print(rec)
                        if len(rec.split()) > 3:
                            result['result_reson'] = rec.split()[3]       # 具体原因
                        if len(rec.split()) > 2:
                            result['result_info'] = rec.split()[2]        # 客诉原因
                        if len(rec.split()) > 1:
                            result['result_content'] = rec.split()[1]     # 处理内容
                        result['dealContent'] = rec.split()[0]            # 最新处理结果

                        rec_name = record.split("#处理结果：")[0]
                        if '赠品' in rec.split()[0] or '退款' in rec.split()[0] or '补发' in rec.split()[0] or '换货' in rec.split()[0]:                    # 筛选无用的通话记录
                            if len(rec_name.split()) > 2:
                                if (rec_name.split())[2] != '' or (rec_name.split())[2] != []:
                                    result['traceUserName'] = (rec_name.split())[2]
                            else:
                                result['traceUserName'] = ''
                        else:
                            result['traceUserName'] = ''
                        ordersDict.append(result.copy())    # append()方法只是将字典的地址存到list中，而键赋值的方式就是修改地址，所以才导致覆盖的问题;  使用copy() 或者 deepcopy()  当字典中存在list的时候需要使用deepcopy()
                else:
                    result['deal_time'] = ''
                    result['result_reson'] = ''
                    result['result_info'] = ''
                    result['result_content'] = ''
                    if len(result['dealContent'].split()) > 3:
                        result['result_reson'] = result['dealContent'].split()[3]       # 具体原因
                    if len(result['dealContent'].split()) > 2:
                        result['result_info'] = result['dealContent'].split()[2]        # 客诉原因
                    if len(result['dealContent'].split()) > 1:
                        result['result_content'] = result['dealContent'].split()[1]     # 处理内容
                    result['dealContent'] = result['dealContent'].split()[0]            # 最新处理结果

                    if result['traceRecord'] != '' or result['traceRecord'] != []:
                        result['deal_time'] = result['traceRecord'].split()[0]
                    if result['traceUserName'] != '' or result['traceUserName'] != []:
                        result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                    result['dealContent'] = result['dealContent'].strip()
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
                data = self._waybill_Query(timeStart, timeEnd, n)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        else:
            dp = df
        dp = dp[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'dealStatus', 'dealContent',
                 'deal_time', 'result_content', 'result_info', 'result_reson', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus',
                 'gift_reissue_order_number', 'update_time']]
        dp.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '最新处理状态', '最新处理结果',
                      '处理时间', '处理内容', '客诉原因', '具体原因', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态',
                      '赠品补发订单编号', '更新时间']
        print('正在写入......')
        dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞|张陈平', na=False))]
        dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
        dp.to_excel('G:\\输出文件\\物流客诉件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        sql = '''REPLACE INTO 物流客诉件(处理时间,物流反馈时间,处理人,订单编号,处理方案, 处理结果, 客诉原因, 赠品补发订单编号,币种, 记录时间) 
                SELECT 处理时间,导入时间 AS 物流反馈时间,处理人,订单编号,最新处理结果 AS 处理方案, 处理内容 AS 处理结果, 客诉原因, 赠品补发订单编号, 币种, NOW() 记录时间 
                FROM customer;'''
        # pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('写入成功......')
        print('*' * 50)
    def _waybill_Query(self, timeStart, timeEnd, n):  # 进入物流客诉件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.orderCustomerComplaint&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'order_trace_id': None, 'question_type': None, 'critical': None, 'read_status': None,
                'operator_type': None, 'operator': None, 'create_time': None, 'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_gift_reissue': None,
                'is_collection': None, 'logistics_status': None, 'user_id': None, 'page': n, 'pageSize': 90}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['traceRecord'] = zhconv.convert(result['traceRecord'], 'zh-hans')
                if ';' in result['traceRecord']:
                    trace_record = result['traceRecord'].split(";")
                    for record in trace_record:
                        result['deal_time'] = ''
                        result['result_reson'] = ''
                        result['result_info'] = ''
                        result['result_content'] = ''
                        result['deal_time'] = record.split()[0]
                        rec = record.split("#处理结果：")[1]
                        if len(rec.split()) > 3:
                            result['result_reson'] = rec.split()[3]       # 具体原因
                        if len(rec.split()) > 2:
                            result['result_info'] = rec.split()[2]        # 客诉原因
                        if len(rec.split()) > 1:
                            result['result_content'] = rec.split()[1]     # 处理内容
                        result['dealContent'] = rec.split()[0]            # 最新处理结果
                        rec_name = record.split("#处理结果：")[0]
                        if '赠品' in rec.split()[0] or '退款' in rec.split()[0] or '补发' in rec.split()[0] or '换货' in rec.split()[0]:                    # 筛选无用的通话记录
                            if len(rec_name.split()) > 2:
                                if (rec_name.split())[2] != '' or (rec_name.split())[2] != []:
                                    result['traceUserName'] = (rec_name.split())[2]
                            else:
                                result['traceUserName'] = ''
                        else:
                            result['traceUserName'] = ''
                        ordersDict.append(result.copy())    # append()方法只是将字典的地址存到list中，而键赋值的方式就是修改地址，所以才导致覆盖的问题;  使用copy() 或者 deepcopy()  当字典中存在list的时候需要使用deepcopy()
                else:
                    result['deal_time'] = ''
                    result['result_reson'] = ''
                    result['result_info'] = ''
                    result['result_content'] = ''
                    if len(result['dealContent'].split()) > 3:
                        result['result_reson'] = result['dealContent'].split()[3]       # 具体原因
                    if len(result['dealContent'].split()) > 2:
                        result['result_info'] = result['dealContent'].split()[2]        # 客诉原因
                    if len(result['dealContent'].split()) > 1:
                        result['result_content'] = result['dealContent'].split()[1]     # 处理内容
                    result['dealContent'] = result['dealContent'].split()[0]            # 最新处理结果

                    if result['traceRecord'] != '' or result['traceRecord'] != []:
                        result['deal_time'] = result['traceRecord'].split()[0]
                    if result['traceUserName'] != '' or result['traceUserName'] != []:
                        result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                    result['dealContent'] = result['dealContent'].strip()
                    ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data


    # 查询更新（新后台的获取-采购问题件）（一、简单查询）
    def sale_Query(self, timeStart, timeEnd):  # 进入采购问题件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': 1, 'pageSize': 90, 'areaId': None, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': None, 'add_time_end': None,
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': timeStart,
                'update_time_end': timeEnd}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['total']
        ordersDict = []
        try:
            for result in req['data']['data']:  # 添加新的字典键-值对，为下面的重新赋值用
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print(data)
        data = data[(data['currencyName'].str.contains('台币|港币', na=False))]
        df = data[['orderNumber',  'create_time', 'dealTime', 'dealName', 'dealProcess', 'description']]
        df.columns = ['订单编号', '反馈时间', '处理时间', '处理人', '处理结果', '取消原因']
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        print(max_count)
        if max_count > 90:
            in_count = math.ceil(max_count/90)
            dlist = []
            n = 1
            while n < in_count:  # 这里用到了一个while循环，穿越过来的
                print('剩余查询次数' + str(in_count - n))
                n = n + 1
                data = self._sale_Query(timeStart, timeEnd, n)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('G:\\输出文件\\采购问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        else:
            print('正在写入......')
            df.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            df.to_excel('G:\\输出文件\\采购问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        sql = '''REPLACE INTO 采购异常(订单编号,处理结果,反馈时间,处理时间,取消原因, 处理人, 电话联系人, 联系时间,记录时间) 
                SELECT 订单编号,处理结果,反馈时间,处理时间,取消原因, 处理人, null 电话联系人, null 联系时间, NOW() 记录时间 
                FROM customer'''
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('写入成功......')
        print('*' * 50)
    def _sale_Query(self, timeStart, timeEnd, n):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': n, 'pageSize': 90, 'areaId': None, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': None, 'add_time_end': None,
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': timeStart,
                'update_time_end': timeEnd}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['data']:  # 添加新的字典键-值对，为下面的重新赋值用
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        # print(data)
        data = data[(data['currencyName'].str.contains('台币|港币', na=False))]
        df = data[['orderNumber',  'create_time', 'dealTime', 'dealName', 'dealProcess', 'description']]
        df.columns = ['订单编号', '反馈时间', '处理时间', '处理人', '处理结果', '取消原因']
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return df

    # 查询更新（新后台的获取-采购问题件）(二、补充查询)
    def sale_Query_info(self, timeStart, timeEnd):  # 进入采购问题件界面--明细查询
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('正在获取 补充订单信息......')
        start = datetime.datetime.now()
        sql = '''SELECT id,`订单编号`  FROM 采购异常 sl WHERE DATE(sl.`处理时间`) BETWEEN '{0}' AND '{1}';'''.format(timeStart, timeEnd)
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        if ordersDict.empty:
            print('无需要更新订单信息！！！')
            return
        print(ordersDict['订单编号'][0])
        df = self._sale_Query_info(ordersDict['订单编号'][0])
        order_list = list(ordersDict['订单编号'])
        max_count = len(order_list)    # 使用len()获取列表的长度，上节学的
        if max_count > 1:
            dlist = []
            for ord in order_list:
                print(ord)
                data = self._sale_Query_info(ord)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        else:
            dp = df
        print('正在写入......')
        dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
        dp.to_excel('G:\\输出文件\\采购问题件-查询-副本{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        sql = '''update 采购异常 a, customer b
                    set a.`电话联系人`= b.`name`,
                        a.`联系时间`= IF( b.`addTime` = '', a.`联系时间`,  b.`addTime`)
                where a.`订单编号`=b.`orderNumber`;'''
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('查询耗时：', datetime.datetime.now() - start)
    def _sale_Query_info(self, ord):  # 进入采购问题件界面--明细查询
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=abnormalDisposeLog'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/purchaseFeedback'}
        data = {'orderNumber': ord}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        data_count = req['data']
        # print(data_count)
        print(len(data_count))
        orders_dict = {}
        try:
            if len(data_count) > 2:
                orders_dict = req['data'][1]
            elif len(data_count) == 2:
                orders_dict = req['data'][0]
            else:
                orders_dict['id'] = ''
                orders_dict['content'] = ''
                orders_dict['orderNumber'] = req['data'][0]['orderNumber']
                orders_dict['userId'] = ''
                orders_dict['addTime'] = ''
                orders_dict['name'] = ''
                orders_dict['avatar'] = ''
                orders_dict['roleName'] = ''
                orders_dict['dealProcess'] = ''
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(orders_dict)
        print(data)
        print('++++++本次 查询成功+++++++')
        print('*' * 50)
        return data



    # 查询更新（新后台的获取-采购问题件）
    def ssale_Query(self, timeStart, timeEnd):  # 进入采购问题件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': 1, 'pageSize': 90, 'areaId': None, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': None, 'add_time_end': None,
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': timeStart,
                'update_time_end': timeEnd}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['total']
        ordersDict = []
        try:
            for result in req['data']['data']:  # 添加新的字典键-值对，为下面的重新赋值用
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        df = pd.json_normalize(ordersDict)
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')      # 获取总单量
        print('*' * 50)
        if max_count != 0:
            if max_count > 90:
                in_count = math.ceil(max_count/90)
                dlist = []
                n = 1
                while n < in_count:  # 这里用到了一个while循环，穿越过来的
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                    data = self._ssale_Query(timeStart, timeEnd, n)                     # 分页获取详情
                    dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = df
            dp = dp[(dp['currencyName'].str.contains('台币|港币', na=False))]           # 筛选币种
            dp = dp[['orderNumber']]
            # print(dp)
            print('++++++明细查询中+++++++')
            # print(dp['orderNumber'][0])
            # dt = self._ssale_Query_info(dp['orderNumber'][0])                           # 查询第一个订单信息
            order_list = list(dp['orderNumber'])
            # print(order_list)
            # dtlist = []
            for ord in order_list:
                data = self._ssale_Query_info(ord)                                      # 查询全部订单信息
                # dtlist.append(data)
            # print(99)
            # print(dtlist)
            # dtlist.to_excel('G:\\输出文件\\采购问题件-查询-副本{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _ssale_Query(self, timeStart, timeEnd, n):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': n, 'pageSize': 90, 'areaId': None, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': None, 'add_time_end': None,
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': timeStart,
                'update_time_end': timeEnd}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['data']:  # 添加新的字典键-值对，为下面的重新赋值用
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print('++++++单次查询成功+++++++')
        print('*' * 50)
        return data
    def _ssale_Query_info(self, ord):  # 进入采购问题件界面--明细查询
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=abnormalDisposeLog'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/purchaseFeedback'}
        data = {'orderNumber': ord}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        req = json.loads(req.text)  # json类型数据转换为dict字典
        order_dict = []     # 初始化列表
        try:
            for result in req['data']:  # 添加新的字典键-值对，为下面的重新赋值用
                order_dict.insert(0, result)   # 指定位置添加数据
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        orders_dict = []    # 数据的列表
        try:
            for ord in order_dict:  # 添加新的字典键-值对，为下面的重新赋值用
                ord['content'] = zhconv.convert(ord['content'], 'zh-hans')
                orders_dict.append(ord.copy())
                data = pd.json_normalize(ord)
                dp = data[['orderNumber', 'content', 'addTime', 'addTime', 'name', 'dealProcess']]
                dp.columns = ['订单编号', '反馈内容', '处理时间', '详细处理时间', '处理人', '处理结果']
                dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞|张陈平', na=False))]
                dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
                sql = '''REPLACE INTO 采购异常(订单编号,处理结果,处理时间,详细处理时间,反馈内容, 处理人,记录时间) 
                        SELECT 订单编号,处理结果,处理时间,详细处理时间,反馈内容, 处理人, NOW() 记录时间 
                        FROM customer;'''
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(orders_dict)
        print('++++++本次 查询成功+++++++')
        print('*' * 50)
        return data

    def order_js_Query(self, timeStart, timeEnd):  # 进入拒收问题件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getRejectList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerRejection'}
        data = {'page': 1, 'pageSize': 500, 'orderPrefix': None, 'shipUsername': None, 'shippingNumber': None, 'email': None, 'saleIds': None, 'ip': None,
                'productIds': None, 'phone': None, 'optimizer': None, 'payment': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                'emailStatus': None, 'befrom': None, 'areaId': None, 'orderStatus': None, 'timeStart': None, 'timeEnd': None, 'payType': None, 'questionId': None,
                'autoVerifys': None, 'reassignmentType': None, 'logisticsStatus': None, 'logisticsId': None, 'traceItemIds': None, 'finishTimeStart': None,
                'finishTimeEnd': None, 'traceTimeStart': timeStart + ' 00:00:00', 'traceTimeEnd': timeEnd + ' 23:59:59','newCloneNumber': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        # print(req)
        ordersDict = []
        if max_count != 0:
            try:
                for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                    # print(result['orderNumber'])
                    result['订单编号'] = result['orderNumber']
                    result['再次克隆下单'] = result['newCloneNumber']
                    result['跟进人'] = ''
                    result['时间'] = ''
                    result['内容'] = ''
                    result['联系方式'] = ''
                    result['问题类型'] = ''
                    result['问题原因'] = ''
                    result['处理结果'] = ''
                    result['是否需要商品'] = ''
                    if result['traceItems'] != []:
                        for res in result['traceItems']:
                            resval = res.split(':')[0]
                            if '跟进人' in resval:
                                result['跟进人'] = (res.split('跟进人:')[1]).strip()  # 跟进人
                            if '时间' in resval:
                                result['时间'] = (res.split('时间:')[1]).strip()  # 跟进人
                            if '内容' in resval:
                                result['内容'] = (res.split('内容:')[1]).strip()  # 跟进人
                            if '联系方式' in resval:
                                result['联系方式'] = (res.split('联系方式:')[1]).strip()  # 跟进人
                            if '问题类型' in resval:
                                result['问题类型'] = (res.split('问题类型:')[1]).strip()  # 跟进人
                            if '问题原因' in resval:
                                result['问题原因'] = (res.split('问题原因:')[1]).strip()  # 跟进人
                            if '处理结果' in res:
                                result['处理结果'] = (res.split('处理结果:')[1]).strip()  # 跟进人
                            if '是否需要商品' in res:
                                result['是否需要商品'] = (res.split('是否需要商品:')[1]).strip()  # 跟进人
                    ordersDict.append(result.copy())
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            df = pd.json_normalize(ordersDict)
            print('*' * 50)
            if max_count > 500:
                in_count = math.ceil(max_count/500)
                dlist = []
                n = 1
                while n < in_count:  # 这里用到了一个while循环，穿越过来的
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                    data = self._order_js_Query(timeStart, timeEnd, n)
                    dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = df
            dp = dp[['订单编号', '再次克隆下单', '跟进人', '时间', '联系方式', '问题类型', '问题原因', '内容', '处理结果', '是否需要商品']]
            dp.columns = ['订单编号', '再次克隆下单', '处理人', '处理时间', '联系方式', '核实原因', '具体原因', '备注', '处理结果', '是否需要商品']
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            df.to_excel('G:\\输出文件\\拒收问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 拒收问题件(订单编号,再次克隆下单,处理人,处理时间,联系方式, 核实原因, 具体原因, 备注, 处理结果, 是否需要商品,记录时间)
                    SELECT 订单编号,IF(再次克隆下单 = '',NULL,再次克隆下单) 再次克隆下单,处理人,处理时间,联系方式, 核实原因, 具体原因, 备注, 处理结果,是否需要商品, NOW() 记录时间
                    FROM customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')

            print('获取每日新增核实拒收表......')
            rq = datetime.datetime.now().strftime('%m.%d')
            sql = '''SELECT 处理时间,IF(团队 LIKE "%红杉%","红杉",IF(团队 LIKE "火凤凰%","火凤凰",IF(团队 LIKE "神龙家族%","神龙",IF(团队 LIKE "金狮%","金狮",IF(团队 LIKE "神龙-主页运营1组%","神龙主页运营",IF(团队 LIKE "金鹏%","小虎队",团队)))))) as 团队,
                            js.订单编号,产品id,产品名称,下单时间,完结状态时间,电话号码,核实原因,具体原因,NULL 通话截图,NULL ID,再次克隆下单,NULL 备注,处理人
                    FROM (SELECT * FROM 拒收问题件 WHERE 记录时间 >= TIMESTAMP(CURDATE())) js
                    LEFT JOIN gat_order_list g ON js.订单编号= g.订单编号;'''
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            df.to_excel('F:\\神龙签收率\\(订   单) 拒收原因-核实\\(上传)订单客户反馈-核实原因 & 再次克隆下单汇总\\{} 需核实拒收-每日上传 - 副本.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')

            df = df[['订单编号', '核实原因', '具体原因', '产品id']]
            df.columns = ['订单编号', '客户反馈', '具体原因', '产品ID']
            df.insert(2, '反馈类型', '拒收')
            df.insert(3, '仓库问题', '否')
            df = df.loc[df["客户反馈"] != "未联系上客户"]
            df.to_excel('F:\\神龙签收率\\(订   单) 拒收原因-核实\\(上传)订单客户反馈-核实原因 & 再次克隆下单汇总\\{} 台湾 - 订单客户反馈(上传).xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            print('获取写入成功......')
        else:
            print('****** 没有信息！！！')
        print('*' * 50)
    def _order_js_Query(self, timeStart, timeEnd, n):  # 进入拒收问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getRejectList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerRejection'}
        data = {'page': 1, 'pageSize': 500, 'orderPrefix': None, 'shipUsername': None, 'shippingNumber': None, 'email': None, 'saleIds': None, 'ip': None,
                'productIds': None, 'phone': None, 'optimizer': None, 'payment': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                'emailStatus': None, 'befrom': None, 'areaId': None, 'orderStatus': None, 'timeStart': None, 'timeEnd': None, 'payType': None, 'questionId': None,
                'autoVerifys': None, 'reassignmentType': None, 'logisticsStatus': None, 'logisticsId': None, 'traceItemIds': None, 'finishTimeStart': None,
                'finishTimeEnd': None, 'traceTimeStart': timeStart + ' 00:00:00', 'traceTimeEnd': timeEnd + ' 23:59:59', 'newCloneNumber': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                result['订单编号'] = result['orderNumber']
                result['再次克隆下单'] = result['newCloneNumber']
                result['跟进人'] = ''
                result['时间'] = ''
                result['内容'] = ''
                result['联系方式'] = ''
                result['问题类型'] = ''
                result['问题原因'] = ''
                result['处理结果'] = ''
                result['是否需要商品'] = ''
                if result['traceItems'] != []:
                    for res in result['traceItems']:
                        resval = res.split(':')[0]
                        if '跟进人' in resval:
                            result['跟进人'] = (res.split('跟进人:')[1]).strip()  # 跟进人
                        if '时间' in resval:
                            result['时间'] = (res.split('时间:')[1]).strip()  # 跟进人
                        if '内容' in resval:
                            result['内容'] = (res.split('内容:')[1]).strip()  # 跟进人
                        if '联系方式' in resval:
                            result['联系方式'] = (res.split('联系方式:')[1]).strip()  # 跟进人
                        if '问题类型' in resval:
                            result['问题类型'] = (res.split('问题类型:')[1]).strip()  # 跟进人
                        if '问题原因' in resval:
                            result['问题原因'] = (res.split('问题原因:')[1]).strip()  # 跟进人
                        if '处理结果' in res:
                            result['处理结果'] = (res.split('处理结果:')[1]).strip()  # 跟进人
                        if '是否需要商品' in res:
                            result['是否需要商品'] = (res.split('是否需要商品:')[1]).strip()  # 跟进人
                ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    def orderReturnList_Query(self, team, timeStart, timeEnd):  # 进入退换货界面
        match = {1: '换补',
                 2: '退货'}
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询 ' + match[team] + '表 信息中')
        url = r'https://gimp.giikin.com/service?name=gorder.postSale&action=getOrderReturnList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerRejection'}
        data = {'menu': team, 'is_deal': 1, 'order_number': None, 'waybill_number': None, 'refund_no': None, 'productId': None, 'order_status': None, 'area_id': None,
                'feedback_type': None, 'type': None, 'question_type': None, 'uid': None, 'username': None, 'critical': None, 'refund_no_check': None, 'is_take': None,
                'currency_id': None, 'pay_type': None,'startTime': timeStart + ' 00:00:00', 'endTime': timeEnd + ' 23:59:59', 'startDealTime': None, 'endDealTime': None,
                'startDoorPickTime': None, 'endDoorPickTime': None, 'page': 1, 'pageSize': 90, 'door_pick_status': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        ordersDict = []
        if max_count != 0:
            try:
                for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                    orderInfo_count = 0
                    for dt in result['orderDetails']:
                        if dt != []:
                            all_index = [substr.start() for substr in re.finditer('』x', dt['spec'])]  # 得到所有  '』x' 下标[6, 18, 27, 34, 39]
                            for index in all_index:
                                orderInfo_count = orderInfo_count + int(dt['spec'][index + 2:index + 3])    # 对下标推移位置，获取个数进行加  『白色,均码』x1,『黑色,均码』x1
                    result['orderInfo_count'] = orderInfo_count
                    result['orderInfo.order_number'] = ''
                    result['orderInfo.currency'] = ''
                    result['orderInfo.area'] = ''
                    result['orderInfo.amount'] = ''
                    result['orderInfoAfter.order_number'] = ''
                    result['orderInfoAfter.amount'] = ''
                    result['orderInfo.order_number'] = result['orderInfo']['order_number']
                    result['orderInfo.currency'] = result['orderInfo']['currency']
                    result['orderInfo.area'] = result['orderInfo']['area']
                    result['orderInfo.amount'] = result['orderInfo']['amount']
                    if result['orderInfoAfter'] != []:
                        result['orderInfoAfter.order_number'] = result['orderInfoAfter']['order_number']
                        result['orderInfoAfter.amount'] = result['orderInfoAfter']['amount']
                    ordersDict.append(result.copy())
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            df = pd.json_normalize(ordersDict)
            print('*' * 50)
            if max_count > 90:
                in_count = math.ceil(max_count/90)
                dlist = []
                n = 1
                while n < in_count:  # 这里用到了一个while循环，穿越过来的
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                    data = self._orderReturnList_Query(team, timeStart, timeEnd, n)
                    dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = df
            dp = dp[['orderInfo.order_number',  'orderInfo.currency', 'orderInfo.area', 'orderInfo.amount', 'orderInfo_count', 'feedback_type', 'question_type',
                    'orderInfoAfter.order_number', 'orderInfoAfter.amount', 'refund_amount', 'create_time', 'user', 'deal_time', 'deal_user', 'type']]
            dp.columns = ['订单编号', '币种', '团队', '金额', '数量', '反馈方式', '反馈问题类型', '新订单编号', '克隆后金额', '退款金额', '导入时间',
                        '登记人', '处理时间', '处理人', '售后类型']
            dp = dp[(dp['币种'].str.contains('港币|台币', na=False))]
            print('共有 ' + str(len(dp)) + '条 正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('G:\\输出文件\\{0}-查询{1}.xlsx'.format(match[team], rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 退换货表(订单编号,币种,团队,金额,数量, 反馈方式, 反馈问题类型, 新订单编号,克隆后金额, 退款金额, 导入时间, 登记人, 处理时间, 处理人, 售后类型, 记录时间) 
                    SELECT 订单编号,币种,团队,金额,数量, 反馈方式, 反馈问题类型, 新订单编号, IF(克隆后金额 = '',NULL,克隆后金额) 克隆后金额, IF(退款金额 = '',NULL,退款金额) 退款金额, 
                        导入时间, 登记人, 处理时间, 处理人, IF(售后类型 = '' OR 售后类型 IS NULL,'退货',售后类型) 售后类型, NOW() 记录时间
                    FROM customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('写入成功......')
        print('*' * 50)
    def _orderReturnList_Query(self, team, timeStart, timeEnd, n):  # 进入退换货界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?name=gorder.postSale&action=getOrderReturnList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerRejection'}
        data = {'menu': team, 'is_deal': 1, 'order_number': None, 'waybill_number': None, 'refund_no': None, 'productId': None, 'order_status': None, 'area_id': None,
                'feedback_type': None, 'type': None, 'question_type': None, 'uid': None, 'username': None, 'critical': None, 'refund_no_check': None, 'is_take': None,
                'currency_id': None, 'pay_type': None,'startTime': timeStart + ' 00:00:00', 'endTime': timeEnd + ' 23:59:59', 'startDealTime': None, 'endDealTime': None,
                'startDoorPickTime': None, 'endDoorPickTime': None, 'page': n, 'pageSize': 90, 'door_pick_status': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                orderInfo_count = 0
                for dt in result['orderDetails']:
                    if dt != []:
                        all_index = [substr.start() for substr in re.finditer('』x', dt['spec'])]  # 得到所有  '』x' 下标[6, 18, 27, 34, 39]
                        for index in all_index:
                            orderInfo_count = orderInfo_count + int(dt['spec'][index + 2:index + 3])    # 对下标推移位置，获取个数进行加  『白色,均码』x1,『黑色,均码』x1
                result['orderInfo_count'] = orderInfo_count
                result['orderInfo.order_number'] = ''
                result['orderInfo.currency'] = ''
                result['orderInfo.area'] = ''
                result['orderInfo.amount'] = ''
                result['orderInfoAfter.order_number'] = ''
                result['orderInfoAfter.amount'] = ''
                result['orderInfo.order_number'] = result['orderInfo']['order_number']
                result['orderInfo.currency'] = result['orderInfo']['currency']
                result['orderInfo.area'] = result['orderInfo']['area']
                result['orderInfo.amount'] = result['orderInfo']['amount']
                result['orderInfoAfter.order_number'] = result['orderInfoAfter']['order_number']
                result['orderInfoAfter.amount'] = result['orderInfoAfter']['amount']
                ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data


if __name__ == '__main__':
    '''
    # -----------------------------------------------自动获取 问题件 状态运行（一）-----------------------------------------
    # 1、 物流问题件；2、物流客诉件；3、物流问题件；4、全部；--->>数据更新切换
    '''
    '''
    已下架获取内容 仓库和对应库存类型
    sku  stock_type =1
    组合 stock_type =2
    混合 stock_type =3
    龟山易速配 whid = 70
    速派八股仓 whid =95
    天马新竹仓 whid =102
    立邦香港顺丰 whid =117
    香港易速配 whid =134
    龟山-神龙备货 whid =166
    龟山-火凤凰备货 whid =198
    天马顺丰仓 whid =204
    '''
    print('正在生成每日新文件夹......')
    time_path: datetime = datetime.datetime.now()
    mkpath = "F:\\神龙签收率\\" + time_path.strftime('%m.%d')
    isExists = os.path.exists(mkpath)
    if not isExists:
        os.makedirs(mkpath)
        os.makedirs(mkpath + "\\产品签收率")
        os.makedirs(mkpath + "\\产品签收率\\直发")
        os.makedirs(mkpath + "\\导状态")
        os.makedirs(mkpath + "\\签收率")
        os.makedirs(mkpath + "\\物流表")
        print('创建成功')
    else:
        print(mkpath + ' 目录已存在')
    print('*' * 50)

    '''
    # -----------------------------------------------自动获取 各问题件 状态运行（二）-----------------------------------------
    '''
    m = QueryTwo('+86-18538110674', 'qyz04163510')
    start: datetime = datetime.datetime.now()

    select = 99
    if int(select) == 1:
        timeStart, timeEnd = m.readInfo('物流问题件')
        m.waybill_InfoQuery(timeStart, timeEnd)                     # 查询更新-物流问题件
    elif int(select) == 2:
        timeStart, timeEnd = m.readInfo('物流客诉件')
        m.waybill_Query(timeStart, timeEnd)                         # 查询更新-物流客诉件
    elif int(select) == 3:
        timeStart, timeEnd = m.readInfo('采购异常')
        # m.sale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                        # 查询更新-采购问题件（一、简单查询）
        # m.sale_Query_info(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                   # 查询更新-采购问题件(二、补充查询)
        m.ssale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))  # 查询更新-采购问题件（一、简单查询）

    elif int(select) == 99:
        timeStart, timeEnd = m.readInfo('物流问题件')
        m.waybill_InfoQuery('2021-12-01', '2021-12-01')             # 查询更新-物流问题件
        m.waybill_InfoQuery(timeStart, timeEnd)                     # 查询更新-物流问题件

        timeStart, timeEnd = m.readInfo('物流客诉件')
        m.waybill_Query(timeStart, timeEnd)                         # 查询更新-物流客诉件

        timeStart, timeEnd = m.readInfo('退换货表')
        for team in [1, 2]:
            m.orderReturnList_Query(team, timeStart, timeEnd)       # 查询更新-退换货

        timeStart, timeEnd = m.readInfo('拒收问题件')
        m.order_js_Query(timeStart, timeEnd)                        # 查询更新-拒收问题件

        timeStart, timeEnd = m.readInfo('派送问题件')
        m.order_js_Query(timeStart, timeEnd)                        # 查询更新-派送问题件

        timeStart, timeEnd = m.readInfo('采购异常')
        m.ssale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                        # 查询更新-采购问题件（一、简单查询）
        # m.sale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                        # 查询更新-采购问题件（一、简单查询）
        # m.sale_Query_info(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                   # 查询更新-采购问题件(二、补充查询)
    print('查询耗时：', datetime.datetime.now() - start)

    '''
    # -----------------------------------------------自动获取 已下架 状态运行（二）-----------------------------------------
    '''
    if int(select) == 99:
        lw = QueryTwoLower('+86-18538110674', 'qyz04163510')
        start: datetime = datetime.datetime.now()
        lw.order_lower('2021-12-31', '2022-01-01', '自动')    # 自动时 输入的时间无效；切为不自动时，有效

        print('查询耗时：', datetime.datetime.now() - start)

    '''
    # -----------------------------------------------自动获取 产品明细、产品预估签收率明细 状态运行（三）-----------------------------------------
    '''
    if int(select) == 99:
        m.my.update_gk_product()  # 更新产品id的列表 --- mysqlControl表
        m.my.update_gk_sign_rate()  # 更新产品预估签收率 --- mysqlControl表

    # -----------------------------------------------测试部分-----------------------------------------
    # timeStart, timeEnd = m.readInfo('物流问题件')

    # m.waybill_InfoQuery('2021-12-01', '2022-01-12')         # 查询更新-物流问题件

    # timeStart, timeEnd = m.readInfo('派送问题件')
    # m.waybill_deliveryList(timeStart, timeEnd)         # 查询更新-派送问题件

    # m.waybill_Query('2022-02-26', '2022-02-26')              # 查询更新-物流客诉件

    # timeStart, timeEnd = m.readInfo('采购异常')
    # m.ssale_Query('2022-02-28', '2022-03-01')                    # 查询更新-采购问题件（一、简单查询）
    # m.sale_Query_info('2021-07-01', '2021-12-01')             # 查询更新-采购问题件 (二、补充查询)

    # m._sale_Query_info('NR112180927421695')

    # for team in [1, 2]:
        # m.orderReturnList_Query(team, '2022-02-15', '2022-02-16')           # 查询更新-退换货

    # timeStart, timeEnd = m.readInfo('拒收问题件')
    # m.order_js_Query('2022-02-24', '2022-02-24')            # 查询更新-拒收问题件



    print('查询耗时：', datetime.datetime.now() - start)