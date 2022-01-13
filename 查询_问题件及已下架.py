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
import zhconv          # transform2_zh_hant：转为繁体;transform2_zh_hans：转为简体

from mysqlControl import MysqlControl
from sqlalchemy import create_engine
from settings import Settings
from settings_sso import Settings_sso
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色
from 查询订单已下架 import QueryTwoLower


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
            dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞', na=False))]
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('G:\\输出文件\\物流问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 物流问题件(处理时间,物流反馈时间,处理人,订单编号,处理结果, 拒收原因, 记录时间) 
                    SELECT 处理时间,导入时间 AS 物流反馈时间,处理人,订单编号,最新处理结果 AS 处理结果, 拒收原因, NOW() 记录时间 
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
                        if len(rec_name.split()) > 2:
                            if (rec_name.split())[2] != '' or (rec_name.split())[2] != []:
                                result['traceUserName'] = (rec_name.split())[2]
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
        dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞', na=False))]
        dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
        dp.to_excel('G:\\输出文件\\物流客诉件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        sql = '''REPLACE INTO 物流客诉件(处理时间,物流反馈时间,处理人,订单编号,处理方案, 处理结果, 客诉原因, 赠品补发订单编号,记录时间) 
                SELECT 处理时间,导入时间 AS 物流反馈时间,处理人,订单编号,最新处理结果 AS 处理方案, 处理内容 AS 处理结果, 客诉原因, 赠品补发订单编号, NOW() 记录时间 
                FROM customer;'''
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
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
                        if len(rec_name.split()) > 2:
                            if (rec_name.split())[2] != '' or (rec_name.split())[2] != []:
                                result['traceUserName'] = (rec_name.split())[2]
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

            print('++++++明细查询中+++++++')
            dt = self._ssale_Query_info(dp['orderNumber'][0])                           # 查询第一个订单信息
            order_list = list(dp['orderNumber'])
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
                dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞', na=False))]
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
    m = QueryTwo('+86-18538110674', 'qyz04163510')
    start: datetime = datetime.datetime.now()

    select = 5
    if int(select) == 1:
        timeStart, timeEnd = m.readInfo('物流问题件')
        m.waybill_InfoQuery(timeStart, timeEnd)                     # 查询更新-物流问题件
    elif int(select) == 2:
        timeStart, timeEnd = m.readInfo('物流客诉件')
        m.waybill_Query(timeStart, timeEnd)                         # 查询更新-物流客诉件
    elif int(select) == 3:
        timeStart, timeEnd = m.readInfo('采购异常')
        m.sale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                        # 查询更新-采购问题件（一、简单查询）
        m.sale_Query_info(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                   # 查询更新-采购问题件(二、补充查询)

    elif int(select) == 4:
        timeStart, timeEnd = m.readInfo('物流问题件')
        m.waybill_InfoQuery('2021-12-01', '2021-12-01')                   # 查询更新-物流问题件
        m.waybill_InfoQuery(timeStart, timeEnd)                     # 查询更新-物流问题件

        timeStart, timeEnd = m.readInfo('物流客诉件')
        m.waybill_Query(timeStart, timeEnd)                         # 查询更新-物流客诉件

        timeStart, timeEnd = m.readInfo('采购异常')
        m.ssale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                        # 查询更新-采购问题件（一、简单查询）
        # m.sale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                        # 查询更新-采购问题件（一、简单查询）
        # m.sale_Query_info(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))                   # 查询更新-采购问题件(二、补充查询)
    print('查询耗时：', datetime.datetime.now() - start)

    '''
    # -----------------------------------------------自动获取 已下架 状态运行（二）-----------------------------------------
    '''
    # lw = QueryTwoLower('+86-18538110674', 'qyz04163510')
    # start: datetime = datetime.datetime.now()
    # lw.order_lower('2021-12-31', '2022-01-01', '自动')
    # print('查询耗时：', datetime.datetime.now() - start)

    '''
    # -----------------------------------------------自动获取 产品明细、产品预估签收率明细 状态运行（三）-----------------------------------------
    '''
    # m.my.update_gk_product()  # 更新产品id的列表 --- mysqlControl表
    # m.my.update_gk_sign_rate()  # 更新产品预估签收率 --- mysqlControl表

    # -----------------------------------------------测试部分-----------------------------------------
    # timeStart, timeEnd = m.readInfo('物流问题件')

    m.waybill_InfoQuery('2021-12-01', '2022-01-12')         # 查询更新-物流问题件
    # m.waybill_Query('2021-12-01', '2022-01-11')             # 查询更新-物流客诉件

    # m.ssale_Query('2021-12-01', '2022-01-12')                    # 查询更新-采购问题件（一、简单查询）
    # m.sale_Query_info('2021-07-01', '2021-12-01')             # 查询更新-采购问题件 (二、补充查询)

    # m._sale_Query_info('NR112180927421695')
    print('查询耗时：', datetime.datetime.now() - start)