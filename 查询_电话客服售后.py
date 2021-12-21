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

from sqlalchemy import create_engine
from settings import Settings
from settings_sso import Settings_sso
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色


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
    # 获取签收表内容-停用
    def readInfo(self, team, last_month):
        print('>>>>>>正式查询中<<<<<<')
        print('正在获取需要订单信息......')
        start = datetime.datetime.now()
        sql = '''SELECT id,`订单编号`  FROM {0} sl WHERE sl.`处理时间` = '{1}';'''.format(team, last_month)
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        print(ordersDict['订单编号'][0])
        if ordersDict.empty:
            print('无需要更新订单信息！！！')
            # sys.exit()
            return
        orderId = list(ordersDict['订单编号'])
        print('获取耗时：', datetime.datetime.now() - start)
        max_count = len(orderId)    # 使用len()获取列表的长度，上节学的
        n = 0
        while n < max_count:        # 这里用到了一个while循环，穿越过来的
            ord = ', '.join(orderId[n:n + 500])
            # print(ord)
            n = n + 500
            self.orderInfoQuery(ord, team)
        print('单日查询耗时：', datetime.datetime.now() - start)

    # 查询更新（新后台的获取-物流问题件）
    def waybill_InfoQuery(self, timeStart, timeEnd):  # 进入订单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': 62,
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
        # print(req)
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['deal_time'] = ''
                result['result_reson'] = ''
                result['result_info'] = ''
                if '拒收' in result['dealContent']:
                    result['result_reson'] = result['dealContent'].split()[1]
                    result['result_info'] = result['dealContent'].split()[2]
                    result['dealContent'] = result['dealContent'].split()[0]
                if result['traceRecord'] != '':
                    result['deal_time'] = result['traceRecord'].split()[0]
                if result['traceUserName'] != '':
                    result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print(data)
        df = data[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'dealStatus', 'dealContent',
                     'deal_time', 'result_reson', 'result_info', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus',
                     'gift_reissue_order_number', 'update_time']]
        df.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '最新处理状态', '最新处理结果',
                        '处理时间', '拒收原因', '具体原因', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态', '赠品补发订单编号', '更新时间']
        # df['处理人'] = (data['处理人'].replace('客服：', '')).copy()
        # df['处理人'] = data['处理人'].replace('客服：', '')
        df['最新处理结果'] = df['最新处理结果'].str.strip()
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
                data = self._waybillInfoQuery(timeStart, timeEnd, n)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
            print('正在写入......')
            dp.to_excel('G:\\输出文件\\物流问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        else:
            print('正在写入......')
            df.to_excel('G:\\输出文件\\物流问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        # return data
        print('写入成功......')
        print('*' * 50)
    def _waybillInfoQuery(self, timeStart, timeEnd, n):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': 62,
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
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['deal_time'] = ''
                result['result_reson'] = ''
                result['result_info'] = ''
                if '拒收' in result['dealContent']:
                    result['result_reson'] = result['dealContent'].split()[1]
                    result['result_info'] = result['dealContent'].split()[2]
                    result['dealContent'] = result['dealContent'].split()[0]
                if result['traceRecord'] != '':
                    result['deal_time'] = result['traceRecord'].split()[0]
                if result['traceUserName'] != '':
                    result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print(data)
        data = data[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'deal_time', 'dealStatus', 'dealContent',
                    'result_reson', 'result_info', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus',
                    'gift_reissue_order_number', 'update_time']]
        data.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '处理时间', '最新处理状态', '最新处理结果',
                        '拒收原因', '具体原因', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态', '赠品补发订单编号', '更新时间']
        # data['处理人'] = (data['处理人'].replace('客服：', '')).copy()
        data['最新处理结果'] = (data['最新处理结果'].str.strip()).copy()
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
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'order_trace_id': '40,41', 'question_type': None, 'critical': None, 'read_status': None,
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
                result['deal_time'] = ''
                result['result_reson'] = ''
                result['result_info'] = ''
                if '拒收' in result['dealContent']:
                    result['result_reson'] = result['dealContent'].split()[1]
                    result['result_info'] = result['dealContent'].split()[2]
                    result['dealContent'] = result['dealContent'].split()[0]
                if result['traceRecord'] != '':
                    result['deal_time'] = result['traceRecord'].split()[0]
                if result['traceUserName'] != '':
                    result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                result['dealContent'] = result['dealContent'].strip()
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print(data)
        df = data[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'dealStatus', 'dealContent',
                     'deal_time', 'result_reson', 'result_info', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus',
                     'gift_reissue_order_number', 'update_time']]
        df.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '最新处理状态', '最新处理结果',
                        '处理时间', '拒收原因', '具体原因', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态', '赠品补发订单编号', '更新时间']
        # df['处理人'] = (data['处理人'].replace('客服：', '')).copy()
        # df['处理人'] = data['处理人'].replace('客服：', '')
        # df['最新处理结果'] = (df['最新处理结果'].str.strip()).copy()
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
                data = self._waybill_Query(timeStart, timeEnd, n)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
            print('正在写入......')
            dp.to_excel('G:\\输出文件\\物流客诉件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        else:
            print('正在写入......')
            df.to_excel('G:\\输出文件\\物流客诉件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        # return data
        print('写入成功......')
        print('*' * 50)
    def _waybill_Query(self, timeStart, timeEnd, n):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.orderCustomerComplaint&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'order_trace_id': '40,41', 'question_type': None, 'critical': None, 'read_status': None,
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
                result['deal_time'] = ''
                result['result_reson'] = ''
                result['result_info'] = ''
                if '拒收' in result['dealContent']:
                    result['result_reson'] = result['dealContent'].split()[1]
                    result['result_info'] = result['dealContent'].split()[2]
                    result['dealContent'] = result['dealContent'].split()[0]
                if result['traceRecord'] != '':
                    result['deal_time'] = result['traceRecord'].split()[0]
                if result['traceUserName'] != '':
                    result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                result['dealContent'] = result['dealContent'].strip()
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print(data)
        data = data[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'deal_time', 'dealStatus', 'dealContent',
                    'result_reson', 'result_info', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus',
                    'gift_reissue_order_number', 'update_time']]
        data.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '处理时间', '最新处理状态', '最新处理结果',
                        '拒收原因', '具体原因', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态', '赠品补发订单编号', '更新时间']
        # data['处理人'] = (data['处理人'].replace('客服：', '')).copy()
        # data['最新处理结果'] = (data['最新处理结果'].str.strip()).copy()
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    # 查询更新（新后台的获取-采购问题件）
    def sale_Query(self, timeStart, timeEnd):  # 进入物流客诉件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': 1, 'pageSize': 90, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': '2021-12-01', 'add_time_end': '2021-12-20',
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': None,
                'update_time_end': None, '_user': 1343, }
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        print(req)
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['deal_time'] = ''
                result['result_reson'] = ''
                result['result_info'] = ''
                if '拒收' in result['dealContent']:
                    result['result_reson'] = result['dealContent'].split()[1]
                    result['result_info'] = result['dealContent'].split()[2]
                    result['dealContent'] = result['dealContent'].split()[0]
                if result['traceRecord'] != '':
                    result['deal_time'] = result['traceRecord'].split()[0]
                if result['traceUserName'] != '':
                    result['traceUserName'] = result['traceUserName'].replace('客服：', '')
                result['dealContent'] = result['dealContent'].strip()
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print(data)
        df = data[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'dealStatus', 'dealContent',
                     'deal_time', 'result_reson', 'result_info', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus',
                     'gift_reissue_order_number', 'update_time']]
        df.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '最新处理状态', '最新处理结果',
                        '处理时间', '拒收原因', '具体原因', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态', '赠品补发订单编号', '更新时间']
        # df['处理人'] = (data['处理人'].replace('客服：', '')).copy()
        # df['处理人'] = data['处理人'].replace('客服：', '')
        # df['最新处理结果'] = (df['最新处理结果'].str.strip()).copy()
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
                data = self._waybill_Query(timeStart, timeEnd, n)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
            print('正在写入......')
            dp.to_excel('G:\\输出文件\\物流客诉件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        else:
            print('正在写入......')
            df.to_excel('G:\\输出文件\\物流客诉件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        # return data
        print('*' * 50)

if __name__ == '__main__':
    m = QueryTwo('+86-18538110674', 'qyz04163510')
    start: datetime = datetime.datetime.now()
    match1 = {'gat': '港台', 'gat_order_list': '港台', 'slsc': '品牌'}
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------

    # m.waybill_InfoQuery('2021-12-21', '2021-12-21')         # 查询更新-物流问题件

    m.waybill_Query('2021-12-21', '2021-12-21')             # 查询更新-物流客诉件

    m.sale_Query('2021-12-21', '2021-12-21')             # 查询更新-采购问题件




    print('查询耗时：', datetime.datetime.now() - start)