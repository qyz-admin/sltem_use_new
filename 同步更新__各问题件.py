import pandas as pd
import os
import datetime
import time
from tqdm import tqdm

import xlwings
import xlsxwriter
import math
import requests
from requests.adapters import HTTPAdapter
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
from 更新_已下架_压单_驳回发货_头程导入提单号 import QueryTwoLower
from 查询_订单检索 import QueryOrder

# -*- coding:utf-8 -*-
class QueryTwo(Settings, Settings_sso):
    def __init__(self, userMobile, password, login_TmpCode,handle, proxy_handle, proxy_id):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self._online()
        # self.sso_online_Two()
        # self.sso__online_handle(login_TmpCode)
        # self.sso__online_auto()
        self.bulid_file()
        if proxy_handle == '代理服务器':
            if handle == '手动':
                self.sso__online_handle_proxy(login_TmpCode, proxy_id)
            else:
                self.sso__online_auto_proxy(proxy_id)
        else:
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
        data = {'user_id': 4139}
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
        elif team == '压单表_已核实':
            sql = '''SELECT DISTINCT 处理时间 FROM {0} d GROUP BY 处理时间 ORDER BY 处理时间 DESC'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['处理时间'][0])
            last_time = (rq + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now()).strftime('%Y-%m-%d')
        elif team == '短信模板':
            sql = '''SELECT DISTINCT DATE_FORMAT(发送时间,'%Y-%m-%d') 发送日期 FROM 短信日志_发送时间 d  GROUP BY DATE_FORMAT(发送时间,'%Y-%m-%d') ORDER BY 发送时间 DESC;'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['发送日期'][0])
            last_time = (rq + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        elif team == '工单列表':
            sql = '''SELECT DISTINCT DATE_FORMAT(提交时间, '%Y-%m-%d') as 提交日期
                    FROM {0} d 
                    WHERE DATE_FORMAT(提交时间, '%Y%m') >= DATE_FORMAT(DATE_SUB(curdate(), INTERVAL 1 MONTH),'%Y%m')
                    GROUP BY DATE_FORMAT(提交时间, '%Y-%m-%d') 
                    ORDER BY DATE_FORMAT(提交时间, '%Y-%m-%d') DESC;'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['提交日期'][0])
            last_time = (rq - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        elif team == '挽单列表':
            sql = '''SELECT DISTINCT 创建时间 FROM {0}_分析 d GROUP BY 创建时间 ORDER BY 创建时间 DESC;'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['创建时间'][0])
            last_time = (rq + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        elif team == '拒收问题件':
            sql = '''SELECT DISTINCT 处理时间 FROM {0} d GROUP BY 处理时间 ORDER BY 处理时间 DESC'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['处理时间'][0])
            last_time = (rq - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        else:
            sql = '''SELECT DISTINCT 处理时间 FROM {0} d GROUP BY 处理时间 ORDER BY 处理时间 DESC'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['处理时间'][0])
            last_time = (rq - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            # last_time = rq.strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        print('******************起止时间：' + team + last_time + ' - ' + now_time + ' ******************')
        return last_time, now_time

    # 查询更新（新后台的获取-物流问题件）
    def waybill_InfoQuery(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入订单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中---物流问题件')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': None,
                'question_type': None, 'critical': None, 'read_status': None, 'operator_type': None, 'operator': None, 'create_time': None,
                'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_collection': None, 'logistics_status': None, 'user_id': None,
                'page': 1, 'pageSize': 90}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        print('*' * 50)
        if max_count != 0:
            n = 1
            if max_count > 90:
                in_count = math.ceil(max_count/90)
                df = pd.DataFrame([])
                dlist = []
                while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                    data = self._waybillInfoQuery(timeStart, timeEnd, n, proxy_handle, proxy_id)
                    dlist.append(data)
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = self._waybillInfoQuery(timeStart, timeEnd, n, proxy_handle, proxy_id)
            print(dp)
            dp = dp[['order_number',  'currency', 'ship_phone', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'dealStatus', 'dealContent',
                     'deal_time', 'dealTime', 'result_reson', 'result_info', 'questionType', 'questionTypeName', 'question_desc', 'traceRecord', 'traceUserName', 'giftStatus', 'operatorName','contact',
                     'gift_reissue_order_number',  'addtime','update_time']]
            dp.columns = ['订单编号', '币种', '联系电话', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '最新处理状态', '最新处理结果',
                          '处理时间', '处理日期时间', '拒收原因', '具体原因', '问题类型状态', '问题类型', '问题描述', '历史处理记录', '处理人', '赠品补发订单状态', '导入人', '联系方式',
                          '赠品补发订单编号', '下单时间', '更新时间']
            # dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞|张陈平', na=False))]
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('F:\输出文件\\物流问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 物流问题件(订单编号, 下单时间, 联系电话, 币种, 问题类型, 物流反馈时间, 导入人,处理时间, 处理日期时间, 处理人, 联系方式,  处理结果,拒收原因, 克隆订单编号, 记录时间) 
                    SELECT 订单编号, 下单时间, 联系电话, 币种, 问题类型, 导入时间 AS 物流反馈时间, 导入人,IF(处理时间 = '',NULL,处理时间) AS 处理时间, IF(处理日期时间 = '',NULL,处理日期时间) AS 处理日期时间, 处理人, 联系方式, IF(最新处理结果 = '',问题类型状态,最新处理结果) AS 处理结果,拒收原因, 赠品补发订单编号 AS 克隆订单编号, NOW() 记录时间 
                    FROM customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _waybillInfoQuery(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': None,
                'question_type': None, 'critical': None, 'read_status': None, 'operator_type': None, 'operator': None, 'create_time': None,
                'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_collection': None, 'logistics_status': None, 'user_id': None,
                'page': n, 'pageSize': 90}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                # print(result['order_number'])
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                if 'traceRecord' in result:
                    result['traceRecord'] = zhconv.convert(result['traceRecord'], 'zh-hans')
                    if '地址;' in result['traceRecord']:
                        result['traceRecord'] = result['traceRecord'].replace('地址;', '地址:')
                    # if ';' in result['traceRecord']:
                    #     trace_record = result['traceRecord'].split(";")
                    #     for record in trace_record:
                    if ';20' in result['traceRecord']:
                        trace_record = result['traceRecord'].split(";20")
                        for i in range(len(trace_record)):
                            if i == 0:
                                record = trace_record[i]
                            else:
                                record = '20' + trace_record[i]
                            # print(record)
                            if record.split("#处理结果：")[1] != '':
                                result['deal_time'] = record.split()[0]
                                result['result_reson'] = ''
                                result['result_info'] = ''
                                rec = record.split("#处理结果：")[1]
                                if len(rec.split()) > 2:
                                    result['result_info'] = rec.split()[2]        # 客诉原因
                                if len(rec.split()) > 1:
                                    result['result_reson'] = rec.split()[1]     # 处理内容
                                result['dealContent'] = rec.split()[0]            # 最新处理结果
                                rec_name = record.split("#处理结果：")[0]
                                if len(rec_name.split()) > 1:
                                    if (rec_name.split())[2] != '' and (rec_name.split())[2] != []:
                                        if '客服' in (rec_name.split())[2]:
                                            result['traceUserName'] = ((rec_name.split())[2]).split("(客服)")[0]
                                        else:
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
                else:
                    result['deal_time'] = result['update_time']
                    result['result_reson'] = ''
                    result['result_info'] = ''
                    result['dealContent'] = ''
                    result['traceUserName'] = ''
                    ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "物流问题件-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    # 查询更新（新后台的获取-物流问题件- 压单核实）
    def waybill_InfoQuery_yadan(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入订单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中---压单核实')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': None,
                'question_type': 111, 'critical': None, 'read_status': None, 'operator_type': None, 'operator': None, 'create_time': None,
                'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_collection': None, 'logistics_status': None, 'user_id': None,
                'page': 1, 'pageSize': 90}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        print('*' * 50)
        if max_count != 0:
            n = 1
            if max_count > 90:
                in_count = math.ceil(max_count/90)
                df = pd.DataFrame([])
                dlist = []
                while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                    data = self._waybillInfoQuery_yadan(timeStart, timeEnd, n, proxy_handle, proxy_id)
                    dlist.append(data)
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = self._waybillInfoQuery_yadan(timeStart, timeEnd, n, proxy_handle, proxy_id)
            print(dp)
            dp.to_excel('F:\输出文件\\压单核实表-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            # dp = dp[(dp['questionTypeName'].str.contains('订单压单（giikin内部专用）'))]
            dp = dp[['order_number',  'deal_time', 'dealContent', 'traceUserName']]
            dp.columns = ['订单编号', '处理时间', '处理结果', '处理人']
            print('正在写入......')
            dp.to_sql('压单表_已核实_info_copy1', con=self.engine1, index=False, if_exists='replace')
            sql = '''REPLACE INTO 压单表_已核实_info(订单编号, 处理时间, 处理结果, 处理人, 记录时间) 
                    SELECT 订单编号, 处理时间, 处理结果, 处理人, NOW() 记录时间 
                    FROM 压单表_已核实_info_copy1;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            sql = '''REPLACE INTO 压单表_已核实(订单编号, 处理时间, 处理结果, 处理人, 记录时间) 
                    SELECT 订单编号, 处理时间, 处理结果, 处理人, NOW() 记录时间 
                    FROM 压单表_已核实_info_copy1;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            dp.to_excel('F:\输出文件\\压单核实表-2查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _waybillInfoQuery_yadan(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customerQuestion&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerQuestion'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'gift_reissue_order_number': None, 'is_gift_reissue': None, 'order_trace_id': None,
                'question_type': 111, 'critical': None, 'read_status': None, 'operator_type': None, 'operator': None, 'create_time': None,
                'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_collection': None, 'logistics_status': None, 'user_id': None,
                'page': n, 'pageSize': 90}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                result['traceRecord'] = zhconv.convert(result['traceRecord'], 'zh-hans')
                if ';' in result['traceRecord']:
                    trace_record = result['traceRecord'].split(";")

                    # for record in trace_record:
                    # print(result['order_number'])
                    # print(trace_record)
                    # print(len(trace_record))

                    record = trace_record[len(trace_record)-1]
                    result['result_reson'] = ''
                    result['result_info'] = ''
                    result['dealContent'] = ''
                    if record.split("#处理结果：")[1] != '':
                        result['deal_time'] = record.split()[0]

                        rec = record.split("#处理结果：")[1]
                        if len(rec.split()) > 2:
                            result['result_info'] = rec.split()[2]        # 客诉原因
                        if len(rec.split()) > 1:
                            result['result_reson'] = rec.split()[1]     # 处理内容
                        result['dealContent'] = rec.split()[0]            # 最新处理结果
                        rec_name = record.split("#处理结果：")[0]
                        if len(rec_name.split()) > 1:
                            if (rec_name.split())[2] != '' and (rec_name.split())[2] != []:
                                if '客服' in (rec_name.split())[2]:
                                    result['traceUserName'] = ((rec_name.split())[2]).split("(客服)")[0]
                                else:
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
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "压单核实-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    # 查询更新（新后台的获取-派送问题件）
    def waybill_deliveryList(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入订单检索界面
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
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
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
                    data = self._waybill_deliveryList(timeStart, timeEnd, n, proxy_handle, proxy_id)
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
            dp.to_excel('F:\输出文件\\派送问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 派送问题件(订单编号,创建时间, 派送问题, 派送问题首次时间, 处理人, 处理记录, 处理时间,备注, 记录时间) 
                    SELECT 订单编号,创建时间, 派送问题, 派送问题首次时间, 处理人, 处理记录, IF(处理时间 = '',NULL,处理时间) 处理时间,备注,NOW() 记录时间 
                    FROM customer'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _waybill_deliveryList(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入派送问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deliveryProblemPackage'}
        data = {'order_number': None, 'waybill_number': None, 'question_level': None, 'question_type': None, 'order_trace_id': None, 'ship_phone': None,
                'page': n, 'pageSize': 90, 'addtime': None, 'question_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'trace_time': None,
                'create_time': None, 'finishtime': None, 'sale_id': None, 'product_id': None, 'logistics_id': None, 'area_id': None, 'currency_id': None,
                'order_status': None, 'logistics_status': None}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "派送问题件-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data


    # 查询更新（新后台的获取-物流客诉件）
    def waybill_Query(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入物流客诉件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.orderCustomerComplaint&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'order_trace_id': None, 'question_type': None, 'critical': None, 'read_status': None,
                'operator_type': None, 'operator': None, 'create_time': None, 'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_gift_reissue': None,
                'is_collection': None, 'logistics_status': None, 'user_id': None, 'page': 1, 'pageSize': 90}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        print('*' * 50)
        dp = None
        if max_count > 0:
            in_count = math.ceil(max_count/500)
            df = pd.DataFrame([])
            dlist = []
            n = 1
            while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                print('剩余查询次数' + str(in_count - n))
                data = self._waybill_Query(timeStart, timeEnd, n, proxy_handle, proxy_id)
                n = n + 1
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        if dp.empty:
            print("今日无更新数据")
        else:
            dp = dp[['order_number',  'currency', 'amount', 'customer_name', 'customer_mobile', 'arrived_address', 'arrived_time', 'create_time', 'dealStatus',  'dealTime', 'deal_time',
                     'traceUserName', 'trace_UserName', 'dealContent', 'deal_Content', 'result_content', 'result_info', 'result_reson', 'questionTypeName', 'question_desc',
                     'giftStatus', 'gift_reissue_order_number', 'contact', 'traceRecord', 'update_time']]
            dp.columns = ['订单编号', '币种', '订单金额', '客户姓名', '客户电话', '客户地址', '送达时间', '导入时间', '最新处理状态', '最新处理时间', '最新客服处理日期',
                          '最新处理人', '最新客服处理人', '最新处理结果', '最新客服处理', '最新客服处理结果', '客诉原因', '具体原因', '问题类型', '问题描述',
                          '赠品补发订单状态', '赠品补发订单编号', '联系方式', '历史处理记录', '更新时间']
            print('正在写入......')
            dp = dp[(dp['最新客服处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞|张陈平|李晓青', na=False))]
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('F:\输出文件\\物流客诉件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 物流客诉件(处理时间,物流反馈时间,处理人,订单编号,处理方案, 处理结果, 客诉原因, 赠品补发订单编号,币种, 历史处理记录, 记录时间) 
                    SELECT 最新客服处理日期 AS 处理时间,  导入时间 AS 物流反馈时间, 最新客服处理人 AS 处理人, 订单编号,最新客服处理 AS 处理方案, 
                           最新客服处理结果 AS 处理结果, 客诉原因, 赠品补发订单编号, 币种, 历史处理记录, NOW() 记录时间 
                    FROM customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        print('*' * 50)
    def _waybill_Query(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入物流客诉件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.orderCustomerComplaint&action=getCustomerComplaintList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'order_number': None, 'waybill_no': None, 'transfer_no': None, 'order_trace_id': None, 'question_type': None, 'critical': None, 'read_status': None,
                'operator_type': None, 'operator': None, 'create_time': None, 'trace_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_gift_reissue': None,
                'is_collection': None, 'logistics_status': None, 'user_id': None, 'page': n, 'pageSize': 90}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                # print(result['order_number'])
                if 'traceRecord' in result:
                    result['dealContent'] = zhconv.convert(result['dealContent'], 'zh-hans')
                    result['traceRecord'] = zhconv.convert(result['traceRecord'], 'zh-hans')
                    if ';' in result['traceRecord']:
                        trace_record = result['traceRecord'].split(";")
                        result['deal_time'] = ''
                        result['result_reson'] = ''
                        result['result_info'] = ''
                        result['result_content'] = ''
                        result['deal_Content'] = ''
                        result['trace_UserName'] = ''
                        for record in trace_record:
                            if '物流' not in record:
                                rec = record.split("#处理结果：")[1]
                                # print(record)
                                # print(rec)
                                if rec != "" and rec != " ":
                                    result['deal_time'] = record.split()[0]
                                    if len(rec.split()) > 3:
                                        result['result_reson'] = rec.split()[3]       # 最新客服 具体原因
                                    if len(rec.split()) > 2:
                                        result['result_info'] = rec.split()[2]        # 最新客服 客诉原因
                                    if len(rec.split()) > 1:
                                        result['result_content'] = rec.split()[1]     # 最新客服 处理结果
                                    result['deal_Content'] = rec.split()[0]           # 最新客服 处理
                                    rec_name = record.split("#处理结果：")[0]
                                    if '客服' in rec_name:
                                        recname = (rec_name.split())[2]
                                        result['trace_UserName'] = recname.replace('(客服)', '')
                        ordersDict.append(result.copy())    # append()方法只是将字典的地址存到list中，而键赋值的方式就是修改地址，所以才导致覆盖的问题;  使用copy() 或者 deepcopy()  当字典中存在list的时候需要使用deepcopy()
                    else:
                        result['deal_time'] = ''
                        result['result_reson'] = ''
                        result['result_info'] = ''
                        result['result_content'] = ''
                        result['deal_Content'] = ''
                        result['trace_UserName'] = ''
                        if len(result['dealContent'].split()) > 3:
                            result['result_reson'] = result['dealContent'].split()[3]       # 最新客服 具体原因
                        if len(result['dealContent'].split()) > 2:
                            result['result_info'] = result['dealContent'].split()[2]        # 最新客服 客诉原因
                        if len(result['dealContent'].split()) > 1:
                            result['result_content'] = result['dealContent'].split()[1]     # 最新客服 处理内容
                        result['deal_Content'] = result['dealContent'].split()[0]           # 最新客服 处理

                        if result['traceRecord'] != '' or result['traceRecord'] != []:
                            result['deal_time'] = result['traceRecord'].split()[0]
                        if result['traceUserName'] != '' or result['traceUserName'] != []:
                            # if '赠品' in result['traceRecord'] or '退款' in result['traceRecord'] or '补发' in result['traceRecord'] or '换货' in result['traceRecord']:
                            if '客服' in result['traceRecord']:
                                result['trace_UserName'] = result['traceUserName'].replace('客服：', '')
                        ordersDict.append(result.copy())
                else:
                    result['deal_time'] = ''
                    result['result_reson'] = ''
                    result['result_info'] = ''
                    result['result_content'] = ''
                    result['deal_Content'] = ''
                    result['trace_UserName'] = ''
                    ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "物流客诉件-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data


    # 查询更新（新后台的获取-采购问题件）（一、简单查询）
    def sale_Query(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入采购问题件界面
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
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
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
                data = self._sale_Query(timeStart, timeEnd, n, proxy_handle, proxy_id)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('F:\输出文件\\采购问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        else:
            print('正在写入......')
            df.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            df.to_excel('F:\输出文件\\采购问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        sql = '''REPLACE INTO 采购异常(订单编号,处理结果,反馈时间,处理时间,取消原因, 处理人, 电话联系人, 联系时间,记录时间) 
                SELECT 订单编号,处理结果,反馈时间,处理时间,取消原因, 处理人, null 电话联系人, null 联系时间, NOW() 记录时间 
                FROM customer'''
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('写入成功......')
        print('*' * 50)
    def _sale_Query(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入物流问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': n, 'pageSize': 90, 'areaId': None, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': None, 'add_time_end': None,
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': timeStart,
                'update_time_end': timeEnd}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
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
    def sale_Query_info(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入采购问题件界面--明细查询
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('正在获取 补充订单信息......')
        start = datetime.datetime.now()
        sql = '''SELECT id,`订单编号`  FROM 采购异常 sl WHERE DATE(sl.`处理时间`) BETWEEN '{0}' AND '{1}';'''.format(timeStart, timeEnd)
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        if ordersDict.empty:
            print('无需要更新订单信息！！！')
            return
        print(ordersDict['订单编号'][0])
        df = self._sale_Query_info(ordersDict['订单编号'][0], proxy_handle, proxy_id)
        order_list = list(ordersDict['订单编号'])
        max_count = len(order_list)    # 使用len()获取列表的长度，上节学的
        if max_count > 1:
            dlist = []
            for ord in order_list:
                print(ord)
                data = self._sale_Query_info(ord, proxy_handle, proxy_id)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        else:
            dp = df
        print('正在写入......')
        dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
        dp.to_excel('F:\输出文件\\采购问题件-查询-副本{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        sql = '''update 采购异常 a, customer b
                    set a.`电话联系人`= b.`name`,
                        a.`联系时间`= IF( b.`addTime` = '', a.`联系时间`,  b.`addTime`)
                where a.`订单编号`=b.`orderNumber`;'''
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('查询耗时：', datetime.datetime.now() - start)
    def _sale_Query_info(self, ord, proxy_handle, proxy_id):  # 进入采购问题件界面--明细查询
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=abnormalDisposeLog'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/purchaseFeedback'}
        data = {'orderNumber': ord}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
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
    def ssale_Query(self, timeStart, timeEnd, proxy_handle, proxy_id, user_caigou):  # 进入采购问题件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': 1, 'pageSize': 90, 'areaId': None, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': None, 'add_time_end': None,
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': timeStart, 'update_time_end': timeEnd}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        if req['data'] != []:
            max_count = req['data']['total']
            print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
            print('*' * 50)
            if max_count != 0 and max_count != []:
                df = pd.DataFrame([])
                dlist = []
                in_count = math.ceil(max_count / 90)
                n = 1
                while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                    data = self._ssale_Query(timeStart, timeEnd, n, proxy_handle, proxy_id)                     # 分页获取详情
                    dlist.append(data)
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                dp = df.append(dlist, ignore_index=True)
                dp = dp[(dp['currencyName'].str.contains('台币|港币', na=False))]  # 筛选币种
                dp = dp[['orderNumber']]
                print('++++++明细查询中+++++++')
                order_list = list(dp['orderNumber'])
                df2 = pd.DataFrame([])
                dtlist = []
                for ord in order_list:
                    data = self._ssale_Query_info(ord, proxy_handle, proxy_id)  # 查询全部订单信息
                    dtlist.append(data)
                dp2 = df2.append(dtlist, ignore_index=True)
                print(99)
                print(dp2)
                dp3 = dp2[['orderNumber', 'content', 'addTime', 'addTime', 'name', 'dealProcess']]
                dp3.columns = ['订单编号', '反馈内容', '处理时间', '详细处理时间', '处理人', '处理结果']
                dp3 = dp3[(dp3['处理人'].str.contains(user_caigou, na=False))]
                dp3.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
                sql = '''REPLACE INTO 采购异常(订单编号,处理结果,处理时间,详细处理时间,反馈内容, 处理人,记录时间) 
                        SELECT 订单编号,处理结果,处理时间,详细处理时间,反馈内容, 处理人, NOW() 记录时间 
                        FROM customer;'''
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
                # dtlist.to_excel('F:\输出文件\\采购问题件-查询-副本{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
                print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _ssale_Query(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入采购问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=getPurchaseAbnormalList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerComplaint'}
        data = {'page': n, 'pageSize': 90, 'areaId': None, 'userId': None, 'dealUser': None, 'currencyId': None, 'orderNumber': None,
                'productId': None, 'timeStart': None, 'timeEnd': None, 'add_time_start': None, 'add_time_end': None,
                'orderType': None, 'lastProcess': None, 'logisticsStatus': None, 'update_time_start': timeStart,
                'update_time_end': timeEnd}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['data']:  # 添加新的字典键-值对，为下面的重新赋值用
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "采购问题件-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(ordersDict)
        print('++++++单次查询成功+++++++')
        print('*' * 50)
        return data
    def _ssale_Query_info(self, ord, proxy_handle, proxy_id):  # 进入采购问题件界面--明细查询
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.afterSale&action=abnormalDisposeLog'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/purchaseFeedback'}
        data = {'orderNumber': ord}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
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
                # data = pd.json_normalize(ord)
                # dp = data[['orderNumber', 'content', 'addTime', 'addTime', 'name', 'dealProcess']]
                # dp.columns = ['订单编号', '反馈内容', '处理时间', '详细处理时间', '处理人', '处理结果']
                # dp = dp[(dp['处理人'].str.contains('蔡利英|杨嘉仪|蔡贵敏|刘慧霞|张陈平|李晓青', na=False))]
                # dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
                # sql = '''REPLACE INTO 采购异常(订单编号,处理结果,处理时间,详细处理时间,反馈内容, 处理人,记录时间)
                #         SELECT 订单编号,处理结果,处理时间,详细处理时间,反馈内容, 处理人, NOW() 记录时间
                #         FROM customer;'''
                # pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "采购问题件 补充消息-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(orders_dict)
        print('++++++本次 查询成功+++++++')
        print('*' * 50)
        return data

    # 查询更新（新后台的获取-拒收问题件）
    def order_js_Query(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入拒收问题件界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getRejectList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36 Edg/111.0.1661.62',
                    'origin': 'https://gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerRejection'}
        data = {'page': 1, 'pageSize': 100, 'orderPrefix': None, 'shipUsername': None, 'shippingNumber': None, 'email': None, 'saleIds': None, 'ip': None,
                'productIds': None, 'phone': None, 'optimizer': None, 'payment': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                'emailStatus': None, 'befrom': None, 'areaId': None, 'orderStatus': None, 'timeStart': None, 'timeEnd': None, 'payType': None, 'questionId': None,
                'autoVerifys': None, 'reassignmentType': None, 'logisticsStatus': None, 'logisticsId': None, 'traceItemIds': -1, 'finishTimeStart': None,
                'finishTimeEnd': None, 'traceTimeStart': timeStart + ' 00:00:00', 'traceTimeEnd': timeEnd + ' 23:59:59','newCloneNumber': None}
        self.session.mount('http://', HTTPAdapter(max_retries=4))
        self.session.mount('https://', HTTPAdapter(max_retries=4))
        try:
            if proxy_handle == '代理服务器':
                proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
                req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies,)
            else:
                req = self.session.post(url=url, headers=r_header, data=data)
            print('+++已成功发送请求......')
            req = json.loads(req.text)  # json类型数据转换为dict字典
            max_count = req['data']['count']
            print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        except requests.exceptions.RequestException as e:
            print(e)
        ordersDict = []
        if max_count != 0:
            df = pd.DataFrame([])
            in_count = math.ceil(max_count / 100)
            dlist = []
            n = 1
            while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                print('剩余查询次数' + str(in_count - n))
                data = self._order_js_Query(timeStart, timeEnd, n, proxy_handle, proxy_id)
                n = n + 1
                dlist.append(data)
                time.sleep(3)
            dp = df.append(dlist, ignore_index=True)
            dp.to_excel('F:\输出文件\\拒收问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            dp = dp[['id', '订单编号', 'currency', 'percentInfo.orderCount', 'percentInfo.rejectCount', 'percentInfo.arriveCount', 'addTime', 'finishTime', 'tel_phone', 'tel_phone', 'ip','cloneUser', 'newCloneUser', 'newCloneStatus', 'newCloneLogisticsStatus', '再次克隆下单', '跟进人', '时间', '联系方式', '问题类型', '问题原因', '内容', '处理结果', '是否需要商品']]
            dp.columns = ['订单id', '订单编号', '币种', '订单总量', '拒收量', '签收量','下单时间', '完成时间', '电话', '联系电话', 'ip','本单克隆人', '新单克隆人', '新单订单状态', '新单物流状态', '再次克隆下单', '处理人', '处理时间', '联系方式', '核实原因', '具体原因', '备注', '处理结果', '是否需要商品']
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            sql = '''REPLACE INTO 拒收问题件(id, 订单id, 订单编号,币种, 订单总量, 拒收量, 签收量, 下单时间, 完成时间, 电话, 联系电话, ip, 本单克隆人, 新单克隆人, 新单订单状态, 新单物流状态, 再次克隆下单,处理人,处理时间,联系方式, 核实原因, 具体原因, 备注, 处理结果, 是否需要商品,记录时间)
                     SELECT null id, 订单id, 订单编号,币种, 订单总量, 拒收量, 签收量, 下单时间, 完成时间, IF(电话 LIKE "852%",RIGHT(电话,LENGTH(电话)-3),IF(电话 LIKE "886%",RIGHT(电话,LENGTH(电话)-3),电话)) 电话, 联系电话, ip, 本单克隆人, 新单克隆人, 新单订单状态, 新单物流状态,  IF(再次克隆下单 = '',NULL,再次克隆下单) 再次克隆下单,处理人,IF(处理时间 = '',NULL,处理时间) AS 处理时间,联系方式, 核实原因, 具体原因, 备注, 处理结果,是否需要商品,NOW() 记录时间
                    FROM customer;'''.format("", "")
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')

            # print('获取每日新增核实拒收表......')
            # rq = datetime.datetime.now().strftime('%m.%d')
            sql = '''SELECT 处理时间,IF(团队 LIKE "%红杉%","红杉",IF(团队 LIKE "火凤凰%","火凤凰",IF(团队 LIKE "神龙家族%","神龙",IF(团队 LIKE "金狮%","金狮",IF(团队 LIKE "神龙-主页运营1组%","神龙主页运营",IF(团队 LIKE "金鹏%","小虎队",团队)))))) as 团队,
                            js.订单编号,js.币种, 产品id,产品名称, js.下单时间,完结状态时间,电话号码,核实原因 as 问题类型,具体原因 as 核实原因,备注 as 具体原因,NULL 通话截图,NULL ID,再次克隆下单,NULL 备注,处理人
                    FROM (SELECT * FROM 拒收问题件 WHERE 记录时间 >= TIMESTAMP(CURDATE())) js
                    LEFT JOIN gat_order_list g ON js.订单编号= g.订单编号;'''
            # df = pd.read_sql_query(sql=sql, con=self.engine1)
            # df.to_excel('F:\\神龙签收率\\(订   单) 拒收原因-核实\\(上传)订单客户反馈-核实原因 & 再次克隆下单汇总\\{} 需核实拒收-每日上传 - 副本.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')

            # df = df[['订单编号', '核实原因', '具体原因', '产品id']]
            # df.columns = ['订单编号', '客户反馈', '具体原因', '产品ID']
            # df.insert(2, '反馈类型', '拒收')
            # df.insert(3, '仓库问题', '否')
            # df = df.loc[df["客户反馈"] != "未联系上客户"]
            # df.to_excel('F:\\神龙签收率\\(订   单) 拒收原因-核实\\(上传)订单客户反馈-核实原因 & 再次克隆下单汇总\\{} 台湾 - 订单客户反馈(上传).xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            print('获取写入成功......')
        else:
            print('****** 没有信息！！！')
        print('*' * 50)
    def _order_js_Query(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入拒收问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getRejectList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36 Edg/111.0.1661.62',
                    'origin': 'https://gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerRejection'}
        data = {'page': n, 'pageSize': 100, 'orderPrefix': None, 'shipUsername': None, 'shippingNumber': None, 'email': None, 'saleIds': None, 'ip': None,
                'productIds': None, 'phone': None, 'optimizer': None, 'payment': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                'emailStatus': None, 'befrom': None, 'areaId': None, 'orderStatus': None, 'timeStart': None, 'timeEnd': None, 'payType': None, 'questionId': None,
                'autoVerifys': None, 'reassignmentType': None, 'logisticsStatus': None, 'logisticsId': None, 'traceItemIds': -1, 'finishTimeStart': None,
                'finishTimeEnd': None, 'traceTimeStart': timeStart + ' 00:00:00', 'traceTimeEnd': timeEnd + ' 23:59:59', 'newCloneNumber': None}
        self.session.mount('http://', HTTPAdapter(max_retries=4))
        self.session.mount('https://', HTTPAdapter(max_retries=4))
        try:
            if proxy_handle == '代理服务器':
                proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
                req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
            else:
                req = self.session.post(url=url, headers=r_header, data=data)
        except requests.exceptions.RequestException as e:
            print(e)
        try:
            req = json.loads(req.text)  # json类型数据转换为dict字典
        except Exception as e:
            print('转化失败： 请求失败', str(Exception) + str(e))
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                print(result['orderNumber'])
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
                        if '时间' in resval and '跟进' not in resval:
                            result['时间'] = (res.split('时间:')[1]).strip()  # 跟进人
                        if '内容' in resval:
                            result['内容'] = (res.split('内容:')[1]).strip()  # 跟进人
                        if '联系方式' in resval:
                            result['联系方式'] = (res.split('联系方式:')[1]).strip()  # 跟进人
                        if '问题类型' in resval:
                            result['问题类型'] = (res.split('问题类型:')[1]).strip()  # 跟进人
                        if '问题原因' in resval:
                            result['问题原因'] = (res.split('问题原因:')[1]).strip()  # 跟进人
                        if '处理结果' in resval:
                            result['处理结果'] = (res.split('处理结果:')[1]).strip()  # 跟进人
                        if '是否需要商品' in resval:
                            result['是否需要商品'] = (res.split('是否需要商品:')[1]).strip()  # 跟进人

                ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 请检查出错点！！！', str(Exception) + str(e))
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "拒收问题件-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    def orderReturnList_Query(self, team, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入退换货界面
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
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        if max_count != 0:
            in_count = math.ceil(max_count/90)
            dlist = []
            df = pd.DataFrame([])
            n = 1
            while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                print('剩余查询次数' + str(in_count - n))
                data = self._orderReturnList_Query(team, timeStart, timeEnd, n, proxy_handle, proxy_id)
                dlist.append(data)
                n = n + 1
            dp = df.append(dlist, ignore_index=True)
            dp = dp[['orderInfo.order_number',  'orderInfo.currency', 'orderInfo.area', 'orderInfo.amount', 'orderInfo_count', 'feedback_type', 'question_type',
                    'orderInfoAfter.order_number', 'orderInfoAfter.amount', 'refund_amount', 'create_time', 'user', 'deal_time', 'deal_user', 'type']]
            dp.columns = ['订单编号', '币种', '团队', '金额', '数量', '反馈方式', '反馈问题类型', '新订单编号', '克隆后金额', '退款金额', '导入时间', '登记人', '处理时间', '处理人', '售后类型']
            dp = dp[(dp['币种'].str.contains('港币|台币', na=False))]
            print('共有 ' + str(len(dp)) + '条 正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('F:\输出文件\\{0}-查询{1}.xlsx'.format(match[team], rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 退换货表(订单编号,币种,团队,金额,数量, 反馈方式, 反馈问题类型, 新订单编号,克隆后金额, 退款金额, 导入时间, 登记人, 处理时间, 处理人, 售后类型, 记录时间) 
                    SELECT 订单编号,币种,团队,金额,数量, 反馈方式, 反馈问题类型, 新订单编号, IF(克隆后金额 = '',NULL,克隆后金额) 克隆后金额, IF(退款金额 = '',NULL,退款金额) 退款金额, 
                        导入时间, 登记人, 处理时间, 处理人, IF(售后类型 = '' OR 售后类型 IS NULL,'退货',售后类型) 售后类型, NOW() 记录时间
                    FROM customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('写入成功......')
        print('*' * 50)
    def _orderReturnList_Query(self, team, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入退换货界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?name=gorder.postSale&action=getOrderReturnList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/customerRejection'}
        data = {'menu': team, 'is_deal': 1, 'order_number': None, 'waybill_number': None, 'refund_no': None, 'productId': None, 'order_status': None, 'area_id': None,
                'feedback_type': None, 'type': None, 'question_type': None, 'uid': None, 'username': None, 'critical': None, 'refund_no_check': None, 'is_take': None,
                'currency_id': None, 'pay_type': None,'startTime': timeStart + ' 00:00:00', 'endTime': timeEnd + ' 23:59:59', 'startDealTime': None, 'endDealTime': None,
                'startDoorPickTime': None, 'endDoorPickTime': None, 'page': n, 'pageSize': 90, 'door_pick_status': None}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值用
                # print(result['id'])
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
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "退换货-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data


    def bulid_file(self):
        print('正在生成每日新文件夹......')
        # file_path = r'F:\\神龙签收率\\B促单指标\\{0} 日统计.xlsx'.format(datetime.datetime.now().strftime('%m.%d'))
        # df = pd.DataFrame([])
        # df.to_excel(file_path, sheet_name='日统计', index=False, engine='xlsxwriter')

        time_path: datetime = datetime.datetime.now()
        mkpath = "F:\\神龙签收率\\" + time_path.strftime('%m.%d')
        isExists = os.path.exists(mkpath)
        if not isExists:
            os.makedirs(mkpath)
            os.makedirs(mkpath + "\\产品签收率")
            os.makedirs(mkpath + "\\产品签收率\\直发&改派")
            os.makedirs(mkpath + "\\导运单号&提货时间")
            os.makedirs(mkpath + "\\导状态")
            os.makedirs(mkpath + "\\签收率")
            os.makedirs(mkpath + "\\物流签收率")
            os.makedirs(mkpath + "\\物流表")
            print('创建成功')
            file_path = mkpath + '\\导运单号&提货时间\\{} 龟山 无运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path1 = mkpath + '\\导运单号&提货时间\\{} 圆通 无运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path2 = mkpath + '\\导运单号&提货时间\\{} 立邦 无运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path3 = mkpath + '\\导运单号&提货时间\\{} 天马 无运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path4 = mkpath + '\\导运单号&提货时间\\{} 速派 无运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path5 = mkpath + '\\导运单号&提货时间\\{} 协来运普货 无运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path50 = mkpath + '\\导运单号&提货时间\\{} 协来运特货 无运单号.xlsx'.format(time_path.strftime('%m.%d'))
            df = pd.DataFrame([['', '']], columns=['订单编号', '物流单号'])
            df.to_excel(file_path, sheet_name='查询', index=False, engine='xlsxwriter')
            df.to_excel(file_path1, sheet_name='查询', index=False, engine='xlsxwriter')
            df.to_excel(file_path2, sheet_name='查询', index=False, engine='xlsxwriter')
            df.to_excel(file_path3, sheet_name='查询', index=False, engine='xlsxwriter')
            df.to_excel(file_path4, sheet_name='查询', index=False, engine='xlsxwriter')
            df.to_excel(file_path5, sheet_name='查询', index=False, engine='xlsxwriter')
            df.to_excel(file_path50, sheet_name='查询', index=False, engine='xlsxwriter')

            file_path31 = mkpath + '\\导运单号&提货时间\\{} 天马 换新运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path32 = mkpath + '\\导运单号&提货时间\\{} 协来运 换新运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path33 = mkpath + '\\导运单号&提货时间\\{} 立邦 换新运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path34 = mkpath + '\\导运单号&提货时间\\{} 速派 换新运单号.xlsx'.format(time_path.strftime('%m.%d'))
            file_path35 = mkpath + '\\导运单号&提货时间\\{} 龟山 换新运单号.xlsx'.format(time_path.strftime('%m.%d'))
            df2 = pd.DataFrame([['', '', '']], columns=['订单编号', '旧运单号', '新运单号'])
            df2.to_excel(file_path31, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path32, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path33, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path34, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path35, sheet_name='查询', index=False, engine='xlsxwriter')

            file_path91 = mkpath + '\\导运单号&提货时间\\{} 导入提货时间 龟山.xlsx'.format(time_path.strftime('%m.%d'))
            file_path92 = mkpath + '\\导运单号&提货时间\\{} 导入提货时间 天马.xlsx'.format(time_path.strftime('%m.%d'))
            file_path93 = mkpath + '\\导运单号&提货时间\\{} 导入提货时间 速派.xlsx'.format(time_path.strftime('%m.%d'))
            file_path94 = mkpath + '\\导运单号&提货时间\\{} 导入提货时间 协来运.xlsx'.format(time_path.strftime('%m.%d'))
            file_path95 = mkpath + '\\导运单号&提货时间\\{} 导入提货时间 立邦.xlsx'.format(time_path.strftime('%m.%d'))
            df2 = pd.DataFrame([['', '', '']], columns=['订单号', '物流单号', '提货时间'])
            df2.to_excel(file_path91, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path92, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path93, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path94, sheet_name='查询', index=False, engine='xlsxwriter')
            df2.to_excel(file_path95, sheet_name='查询', index=False, engine='xlsxwriter')
            print('创建文件')
        else:
            print(mkpath + ' 目录已存在')
        print('*' * 50)


    # 停用
    # 检查昨日订单是否有重复的 （单点的获取）
    def order_check(self, begin, end): # 进入订单检索界面
        # print('正在获取需要订单信息......')
        for i in range((end - begin).days):             # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            last_month = str(day)
            print('正在检查 港台 ' + last_month + ' 号订单信息…………')
            start = datetime.datetime.now()
            sql = '''SELECT id,`订单编号`  FROM {0} sl WHERE sl.`日期` = '{1}';'''.format('gat_order_list', last_month)
            ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
            if ordersDict.empty:
                print('无需要更新订单信息！！！')
                return
            # print(ordersDict['订单编号'][0])
            orderId = list(ordersDict['订单编号'])
            dlist = []
            for index, ord in enumerate(tqdm(orderId)):
                tem_data = self._order_check(ord)
                if tem_data == 1:
                    dlist.append(ord)
            if dlist == [] or len(orderId) == 0:
                print('今日查询无错误订单：', datetime.datetime.now() - start)
            else:
                print('已发送错误订单中：.......')
                dlist = ','.join(dlist)
                url = "https://oapi.dingtalk.com/robot/send?access_token=bdad3de3c4f5e8cc690a122779a642401de99063967017d82f49663382546f30"  # url为机器人的webhook
                content = dlist                  # 钉钉消息内容，注意test是自定义的关键字，需要在钉钉机器人设置中添加，这样才能接收到消息
                mobile_list = ['18538110674']           # 要@的人的手机号，可以是多个，注意：钉钉机器人设置中需要添加这些人，否则不会接收到消息
                isAtAll = '是'                            # 是否@所有人
                self.send_dingtalk_message(url, content, mobile_list, isAtAll)
            print('查询耗时：', datetime.datetime.now() - start)
    def _order_check(self, ord):  # 进入订单检索界面
        # print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500,
                'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None,
                'type': None, 'collId': None, 'isClone': None,
                'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': '',
                'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None, 'tuan': None,
                'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None, 'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None, 'weightStart': None,
                'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None, 'sortField': None,
                'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None}
        data.update({'orderPrefix': ord,
                    'shippingNumber': None})
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print(req.text)
        # print('+++已成功发送请求......')
        # print('正在处理json数据转化为dataframe…………')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        tem = 0
        try:
            max_count = req['data']['count']
            if max_count > 1:
                tem = 1
        except Exception as e:
            print('更新失败：', str(Exception) + str(e))
        # print('*************************本次获取成功***********************************')
        return tem


    # 短信模板
    def getMessage_Log(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入短信模板界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中---短信模板......')
        url = r'https://gimp.giikin.com/service?service=gorder.sms&action=getMessageLog'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/workOrderCenter'}
        data = {'page': 1, 'pageSize': 90, 'order_number': None, 'waybill_number': None, 'to_phone': None,
                'add_date': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'send_status': None, 'msgid': None,
                'template_id': "217,216,215,214,210,209,208,207,206,205,204,175,174,173,172,171,170,151,148,147,136,127,102,101,100,90,89,88,87,86,85,84,83,82,78,77,76,74,73,72,71,70,69,68,52,50,49,36,35,34,31"}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        print('*' * 50)
        if max_count != 0:
            n = 1
            in_count = math.ceil(max_count/90)
            df = pd.DataFrame([])
            dlist = []
            while n <= in_count:
                data = self._getMessage_Log(timeStart, timeEnd, n, proxy_handle, proxy_id)
                dlist.append(data)
                print('剩余查询次数' + str(in_count - n))
                n = n + 1
            dp = df.append(dlist, ignore_index=True)
            dp = dp[['id','order_number','areaName','currency','to_phone','add_date','sendStatus','content','receiveStatus','receive_msg','typeName','templateName']]
            dp.columns = ['id','订单编号','团队','币种','发送者的电话号码','发送时间','是否发送成功','短信内容','接收状态','接收异常原因','短信用途','短信模板']
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('F:\输出文件\\短信模板-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 短信日志_发送时间(id, 订单编号,团队,币种,发送者的电话号码,发送时间,是否发送成功,短信内容,接收状态,接收异常原因,短信用途,短信模板,记录时间)
                     SELECT id, 订单编号,团队,币种,发送者的电话号码,IF(发送时间 = '',NULL, 发送时间) AS 发送时间,是否发送成功,短信内容,接收状态,接收异常原因,短信用途,短信模板,NOW() 记录时间
                    FROM  customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _getMessage_Log(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入短信模板界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.sms&action=getMessageLog'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/workOrderCenter'}
        data = {'page': n, 'pageSize': 90, 'order_number': None, 'waybill_number': None, 'to_phone': None,
                'add_date': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'send_status': None, 'msgid': None,
                'template_id': "147,148,31,35,36,49,50,52,68,69,70,71,72,73,74,76,77,78,82,83,84,85,86,87,88,89,90,100,101,102,151,136,127,34"}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        ordersDict = []
        if max_count > 0:
            try:
                for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                    ordersDict.append(result.copy())
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
                sso = Settings_sso()
                sso.send_dingtalk_message(
                    "https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "短信日志-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
            data = pd.json_normalize(ordersDict)
        else:
            data = None
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    # 工单列表
    def getOrderCollectionList(self, timeStart, timeEnd, proxy_handle, proxy_id):  # 进入工单列表界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中---工单列表......' + str(timeStart) + ' : ' + str(timeEnd))
        url = r'https://gimp.giikin.com/service?service=gorder.orderCollection&action=getOrderCollectionList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/workOrderCenter'}
        data = {'page': 1, 'pageSize': 500, 'order_number': None, 'waybill_number': None,
                'plate_status': None, 'do_status': None, 'collection_type': None,
                'intime[]': [timeStart + ' 00:00:00', timeEnd + ' 23:59:59']
                }
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print(max_count)
        print('++++++本批次查询成功;  总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
        print('*' * 50)
        if max_count != 0:
            n = 1
            if max_count > 1000:
                in_count = math.ceil(max_count/1000)
                df = pd.DataFrame([])
                dlist = []
                while n <= in_count:
                    data = self._getOrderCollectionList(timeStart, timeEnd, n, proxy_handle, proxy_id)
                    dlist.append(data)
                    print('剩余查询次数' + str(in_count - n))
                    n = n + 1
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = self._getOrderCollectionList(timeStart, timeEnd, n, proxy_handle, proxy_id)
            dp = dp[['detail_id','order_number','area_name','currency_name','waybill_number','ship_phone','payType','order_status','logistics_status','logistics_name','reassignmentTypeName','addtime','delivery_time',
                    'finishtime','question_type','step','channel','source','intime','serviceName','operator','collectionType','dealOperatorName','deal_time',
                    'dealContent','dealStatus','traceRecord', 'do_status','sync_operator','sync_data.deal_id','sync_data.create_time','sync_data.sync_type','sync_data_all']]
            dp.columns = ['工单编号','订单编号','所属团队','币种','运单编号','电话','支付方式','订单状态','物流状态','物流渠道','订单类型','下单时间','发货时间',
                          '完成时间','问题类型','环节问题','来源渠道','提交形式','提交时间','受理客服','登记人','工单类型','最新处理人','最新处理时间',
                          '最新处理描述','最新处理结果','处理记录','是否完成','同步人','同步状态','同步时间','同步类型','同步操作记录']
            print('正在写入......')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            dp.to_excel('F:\输出文件\\工单列表-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 工单列表(工单编号, 订单编号,所属团队,币种,运单编号,电话,支付方式,订单状态,物流状态,物流渠道,订单类型,下单时间,
			                    发货时间,完成时间,问题类型,环节问题,来源渠道,提交形式,提交时间,受理客服,登记人,工单类型,
			                    最新处理人,最新处理时间,最新处理描述,最新处理结果,处理记录,是否完成,同步人,同步状态,同步时间,同步类型,同步操作记录,记录时间)
                    SELECT 工单编号, 订单编号,所属团队,币种,运单编号,电话,支付方式,订单状态,物流状态,物流渠道,订单类型,下单时间,IF(发货时间 = '',NULL, 发货时间) AS 发货时间,IF(完成时间 = '',NULL, 完成时间) AS 完成时间,
                                问题类型,环节问题,来源渠道,提交形式,IF(提交时间 = '',NULL, 提交时间) AS 提交时间,受理客服,登记人,工单类型,最新处理人,IF(最新处理时间 = '',NULL, 最新处理时间) AS 最新处理时间,最新处理描述,
                                最新处理结果,处理记录,是否完成,同步人,同步状态,IF(同步时间 = '',NULL, 同步时间) AS 同步时间,同步类型,同步操作记录,NOW() 记录时间
                    FROM  customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _getOrderCollectionList(self, timeStart, timeEnd, n, proxy_handle, proxy_id):  # 进入工单界面界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.orderCollection&action=getOrderCollectionList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/workOrderCenter'}
        data = {'page': n, 'pageSize': 1000, 'order_number': None, 'waybill_number': None,
                'plate_status': None, 'do_status': None, 'collection_type': None,
                'intime[]': [timeStart + ' 00:00:00', timeEnd + ' 23:59:59']
                }
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        ordersDict = []
        if max_count > 0:
            try:
                for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                    # print(result['order_number'])
                    result['sync_data.deal_id'] = ""         # 添加新的字典键-值对，为下面的重新赋值用
                    result['sync_data.create_time'] = ""
                    result['sync_data.sync_type'] = ""
                    if result['sync_data'] != []:
                        result['sync_data.deal_id'] = result['sync_data'][0]['deal_id']
                        result['sync_data.create_time'] = result['sync_data'][0]['create_time']
                        result['sync_data.sync_type'] = result['sync_data'][0]['sync_type']
                    ordersDict.append(result.copy())
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
                sso = Settings_sso()
                sso.send_dingtalk_message(
                    "https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e","工单列表-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
            data = pd.json_normalize(ordersDict)
        else:
            data = None
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data

    # 挽单列表
    # 绩效-查询 挽单列表（一.2）
    def getRedeemOrderList(self, timeStart, timeEnd, proxy_handle, proxy_id):    # 进入订单检索界面     挽单列表查询
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('正在查询 挽单列表 起止时间：' + str(timeStart) + " *** " + str(timeEnd))
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getRedeemOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/saveOrder'}
        data = {'order_number': None, 'type': None, 'order_status': None, 'logistics_status': None, 'old_order_status': None, 'old_logistics_status': None, 'operator': None,
                'create_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_del': None, 'page': 1, 'pageSize': 10, 'area_id': None}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        req = json.loads(req.text)  # json类型数据转换为dict字典
        max_count = req['data']['count']
        print('共...' + str(max_count) + '...单量')
        if max_count != 0:
            df = pd.DataFrame([])
            n = 1
            in_count = math.ceil(max_count / 90)
            dlist = []
            while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                data = self._getRedeemOrderList(timeStart, timeEnd, n, proxy_handle, proxy_id)
                dlist.append(data)
                print('剩余查询次数' + str(in_count - n))
                n = n + 1
            dp = df.append(dlist, ignore_index=True)
            dp = dp[['id', 'order_number', 'redeemType', 'oldOrderStatus', 'oldLogisticsStatus', 'oldAmount', 'orderStatus','logisticsStatus','amount','logisticsName','operatorName','create_time','save_money','currencyName', 'delOperatorName','del_reason']]
            dp.columns = ['id', '订单编号', '挽单类型', '原订单状态', '原物流状态', '原订单金额', '当前订单状态', '当前物流状态','当前订单金额','当前物流渠道','创建人','创建时间','挽单金额','币种', '删除人', '删除原因']
            dp.to_excel('F:\\输出文件\\挽单列表-分析{}.xlsx'.format(rq), sheet_name='挽单', index=False, engine='xlsxwriter')
            dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
            sql = '''REPLACE INTO 挽单列表_分析(id, 订单编号,币种, 创建时间, 创建人, 挽单类型, 挽单金额, 当前订单状态, 当前物流状态, 回款状态, 删除人, 删除原因, 统计月份,记录时间) 
                    SELECT id, 订单编号,币种, 创建时间, 创建人, 挽单类型, 挽单金额, 当前订单状态, 当前物流状态,NULL as 回款状态, 删除人, 删除原因, DATE_FORMAT(创建时间,'%Y%m') 统计月份,NOW() 记录时间 
                    FROM customer;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        print('-' * 50)
        print('-' * 50)
    def _getRedeemOrderList(self, timeStart, timeEnd, n, proxy_handle, proxy_id):
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getRedeemOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'order_number': None, 'type': None, 'order_status': None, 'logistics_status': None, 'old_order_status': None, 'old_logistics_status': None, 'operator': None,
                'create_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'is_del': None, 'page': n, 'pageSize': 90, 'area_id': None}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        ordersdict = []
        # print('正在处理json数据转化为dataframe…………')
        try:
            for result in req['data']['list']:
                ordersdict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            sso = Settings_sso()
            sso.send_dingtalk_message("https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e", "挽单列表-更新失败，请检查原因》》》本地数据库：：", ['18538110674'], "是")
        df = pd.json_normalize(ordersdict)
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return df


if __name__ == '__main__':
    start: datetime = datetime.datetime.now()
    # timeStart = datetime.date(2023, 3, 27)
    # timeEnd = datetime.date(2023, 4, 3)
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
    '''
    # -----------------------------------------------自动获取 各问题件 状态运行（二）-----------------------------------------
    '''
    select = 99
    if int(select) == 909:
        handle = "手动0"
        login_TmpCode = 'c584b7efadac33bb94b2e583b28c9514'          # 输入登录口令Tkoen
        proxy_handle = '代理服务器0'
        proxy_id = '192.168.13.89:37469'                            # 输入代理服务器节点和端口

        m = QueryTwo('+86-17596568562', 'xhy123456', login_TmpCode, handle, proxy_handle, proxy_id)
        start: datetime = datetime.datetime.now()
        m.waybill_InfoQuery('2021-12-01', '2021-12-01', proxy_handle, proxy_id)     #   -- 每日启动测试使用

        timeStart, timeEnd = m.readInfo('拒收问题件')                                        # 查询更新-拒收问题件 @--ok
        m.order_js_Query(timeStart, timeEnd, proxy_handle, proxy_id)

        timeStart, timeEnd = m.readInfo('物流问题件')
        m.waybill_InfoQuery(timeStart, timeEnd, proxy_handle, proxy_id)                     # 查询更新-物流问题件
        
        timeStart, timeEnd = m.readInfo('压单表_已核实')
        m.waybill_InfoQuery_yadan(timeStart, timeEnd, proxy_handle, proxy_id)               # 查询更新-物流问题件 - 压单核实
        
        timeStart, timeEnd = m.readInfo('物流客诉件')
        m.waybill_Query(timeStart, timeEnd, proxy_handle, proxy_id)                         # 查询更新-物流客诉件 @--ok
        
        timeStart, timeEnd = m.readInfo('退换货表')
        for team in [1, 2]:
            m.orderReturnList_Query(team, timeStart, timeEnd, proxy_handle, proxy_id)                   # 查询更新-退换货
        
        timeStart, timeEnd = m.readInfo('派送问题件')
        m.waybill_deliveryList(timeStart, timeEnd, proxy_handle, proxy_id)                              # 查询更新-派送问题件、
        
        timeStart, timeEnd = m.readInfo('采购异常')
        user_caigou = "蔡利英|杨嘉仪|蔡贵敏|刘慧霞|张陈平|李晓青"
        m.ssale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'), proxy_handle, proxy_id, user_caigou)  # 查询更新-采购问题件（一、简单查询）

        timeStart, timeEnd = m.readInfo('短信模板')
        m.getMessage_Log(timeStart, timeEnd, proxy_handle, proxy_id)                            # 查询更新-短信模板 @--ok
        
        timeStart, timeEnd = m.readInfo('工单列表')
        m.getOrderCollectionList(timeStart, timeEnd, proxy_handle, proxy_id)                  # 工单列表-物流客诉件

        timeStart, timeEnd = m.readInfo('挽单列表')
        m.getRedeemOrderList(timeStart, timeEnd, proxy_handle, proxy_id)                  # 挽单列表-物流客诉件
        print('查询耗时：', datetime.datetime.now() - start)
    '''
    # -----------------------------------------------自动获取 单点 昨日头程直发渠道 & 天马711的订单明细  | 删单原因 状态运行（二）-----------------------------------------
    '''
    if int(select) == 909:
        proxy_handle = '代理服务器0'
        proxy_id = '192.168.13.89:37467'                            # 输入代理服务器节点和端口
        handle = '手0动'
        login_TmpCode = '517e55c6fb6c34ca99a69874aaf5ec25'          # 输入登录口令Tkoen
        js = QueryOrder('+86-17596568562', 'xhy123456', login_TmpCode, handle, proxy_handle, proxy_id)

        # time_yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        # time_now = (datetime.datetime.now()).strftime('%Y-%m-%d')
        # query = '下单时间'
        # areaId = None                                   # 团队id
        # js.order_Query_Yiwudi(time_yesterday, time_now, areaId, query, proxy_handle, proxy_id)   # 检查 头程直发渠道 & 天马711@--ok


        timeStart = '2023-03-21'
        timeEnd = '2023-03-23'
        query = '下单时间'
        areaId = None                                   # 团队id
        time_handle = '自动'
        js.order_Query_Delete(timeStart, timeEnd, areaId, query, proxy_handle, proxy_id, time_handle)   # 最近三天删单原因@--ok


        time_handle = '自动'
        timeStart = '2023-04-01'
        timeEnd = '2023-04-27'
        js.order_track_Query(time_handle, timeStart, timeEnd, proxy_handle, proxy_id)  # 促单查询；订单检索@--ok

    '''
    # -----------------------------------------------自动获取 数据库 产品明细、产品预估签收率明细 状态运行（三）-----------------------------------------
    '''
    if int(select) == 909:
        my = MysqlControl()
        my.update_gk_product()  # 更新产品id的列表 --- mysqlControl表
        my.update_gk_sign_rate()  # 更新产品预估签收率 --- mysqlControl表

    '''
    # -----------------------------------------------自动获取 仓储 已下架 状态运行（四）-----------------------------------------
    '''
    if int(select) == 99:
        login_TmpCode = '655ffc6d056d37ca92e4398aff91288c'          # 输入登录口令Tkoen
        handle = '手0动'
        proxy_handle = '代理服务器0'
        proxy_id = '192.168.13.89:37469'                            # 输入代理服务器节点和端口
        lw = QueryTwoLower('+86-17596568562', 'xhy123456', login_TmpCode, handle, proxy_id, proxy_handle)
        start: datetime = datetime.datetime.now()

        lw.order_lower('2021-12-31', '2022-01-01', '自动')    # 已下架       更新； 自动时 输入的时间无效；切为不自动时，有效

        lw.readFile(1)                                        # 上传每日压单核实结果
        lw.order_spec()                                       # 压单         更新；压单反馈  （备注（压单核实是否需要））

        handle = '手动'
        timeStart = datetime.date(2023, 3, 1)
        timeEnd = datetime.date(2023, 3, 12)
        for i in range((timeEnd - timeStart).days):                  # 按天循环获取订单状态
            day = timeStart + datetime.timedelta(days=i)
            day_time = str(day)
            # lw.second_sendorder(handle, day_time, day_time)          # 获取 订单列表的 二次改派原运单号 信息

        # lw.stockcompose_upload()                              # 获取 桃园仓重出、
        # lw.get_take_delivery_no()                             # 头程物流跟踪 更新； 获取最近10天的信息
        print('查询耗时：', datetime.datetime.now() - start)







    '''
    # -----------------------------------------------测试部分-----------------------------------------
    '''
    # handle = '手动0'
    # login_TmpCode = 'c584b7efadac33bb94b2e583b28c9514'  # 输入登录口令Tkoen
    # proxy_handle = '代理服务器0'
    # proxy_id = '192.168.13.89:37467'  # 输入代理服务器节点和端口
    # m = QueryTwo('+86-17596568562', 'xhy123456', login_TmpCode, handle, proxy_handle, proxy_id)
    # start: datetime = datetime.datetime.now()


    # timeStart, timeEnd = m.readInfo('压单表_已核实')
    # m.waybill_InfoQuery_yadan(timeStart, timeEnd)  # 查询更新-物流问题件 - 压单核实
    # m.waybill_InfoQuery_yadan('2022-11-09', '2022-11-09')  # 查询更新-物流问题件 - 压单核实

    # timeStart, timeEnd = m.readInfo('短信模板')
    # m.getMessage_Log('2023-01-01', '2023-03-31', proxy_handle, proxy_id)  # 查询更新-短信模板


    # begin = datetime.date(2022, 1, 1)
    # end = datetime.date(2022, 7, 1)
    # for i in range((end - begin).days):                             # 按天循环获取订单状态
    #     day = begin + datetime.timedelta(days=i)
    #     day_time = str(day)
    #     print('****** 更新      起止时间：' + day_time + ' - ' + day_time + ' ******')
    #     m.getMessage_Log(day_time, day_time, proxy_handle, proxy_id)  # 查询更新-短信模板

    #
    # begin = datetime.date(2022, 10, 1)
    # end = datetime.date(2022, 11, 1)
    # m.order_check(begin, end)                 # 检查昨日订单是否有重复的

    # timeStart, timeEnd = m.readInfo('物流问题件')
    # m.waybill_InfoQuery('2022-09-19', '2022-09-22')  # 查询更新-物流问题件
    # m.waybill_InfoQuery('2023-05-23', '2023-05-25', proxy_handle, proxy_id)  # 查询更新-物流问题件


    # timeStart, timeEnd = m.readInfo('派送问题件')
    # m.waybill_deliveryList(timeStart, timeEnd)         # 查询更新-派送问题件


    # m.waybill_Query('2023-03-08', '2023-03-20', proxy_handle, proxy_id)              # 查询更新-物流客诉件

    # m.ssale_Query('2023-03-08', '2023-03-22', proxy_handle, proxy_id)              # 查询更新-采购问题件

    # timeStart, timeEnd = m.readInfo('拒收问题件')
    # m.order_js_Query('2023-02-20', '2023-02-20', proxy_handle, proxy_id)

    # timeStart, timeEnd = m.readInfo('采购异常')
    # m.ssale_Query('2022-02-28', '2022-03-01')                    # 查询更新-采购问题件（一、简单查询）
    # m.sale_Query_info('2021-07-01', '2021-12-01')             # 查询更新-采购问题件 (二、补充查询)

    # m.ssale_Query('2022-04-28', datetime.datetime.now().strftime('%Y-%m-%d'))      # 查询更新-采购问题件（一、简单查询）
    # m.sale_Query(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))          # 查询更新-采购问题件（一、简单查询）
    # m.sale_Query_info(timeStart, datetime.datetime.now().strftime('%Y-%m-%d'))      # 查询更新-采购问题件(二、补充查询)

    # m._sale_Query_info('NR112180927421695')

    # timeStart = '2023-03-01'
    # timeEnd = '2023-06-02'
    # m.getOrderCollectionList(timeStart, timeEnd, proxy_handle, proxy_id)   # 工单列表-物流客诉件

    # timeStart, timeEnd = m.readInfo('挽单列表')
    # m.getRedeemOrderList(timeStart, timeEnd, proxy_handle, proxy_id)  # 挽单列表-物流客诉件

    # for team in [1, 2]:
        # m.orderReturnList_Query(team, '2022-02-15', '2022-02-16')           # 查询更新-退换货

    # timeStart, timeEnd = m.readInfo('拒收问题件')
    # m.order_js_Query('2022-04-28', '2022-05-04')            # 查询更新-拒收问题件



    print('查询耗时：', datetime.datetime.now() - start)