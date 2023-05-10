import pandas as pd
import os
import datetime
import time
import xlwings

import requests
import json
import sys
from queue import Queue
from dateutil.relativedelta import relativedelta
from threading import Thread #  使用 threading 模块创建线程
import pandas.io.formats.excel
import win32api,win32con
import math
from sqlalchemy import create_engine
from settings import Settings
from settings_sso import Settings_sso
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, \
    Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from 查询_产品明细 import QueryTwoT

# -*- coding:utf-8 -*-
class Query_sso_updata(Settings):
    def __init__(self, userMobile, password, userID, login_TmpCode,handle, proxy_handle, proxy_id):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        self.userID = userID
        self.login_TmpCode = login_TmpCode
        # self._online()
        # self._online_Two()
        # self.sso__online_handle()
        # self._online_Threed()
        # self.sso__online_auto()
        if proxy_handle == '代理服务器':
            if handle == '手动':
                self.sso__online_handle_proxy(login_TmpCode, proxy_id)
            else:
                # self.sso__online_auto_host()
                self.sso__online_auto_gp_proxy(proxy_id)
        else:
            if handle == '手动':
                self.sso__online_handle(login_TmpCode)
            else:
                # self.sso__online_auto_host()
                self.sso__online_auto_gp()
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
        self.dk = Settings_sso()    # 钉钉发送
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

    def _online(self):  # 登录系统保持会话状态
        print('正在登录后台系统中......')
        # print('第一阶段获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171637544299948385570258461637545377418833G01447E2DCD07109775CD567044AE05FC09628C',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        req = req.json()
        # print(req)
        # req_url = req['data']
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        # print('+++已获取loginTmpCode值: ' + loginTmpCode)
        print(datetime.datetime.now())
        if 'data' in req.keys():
            try:
                req_url = req['data']
                loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self._online()
        elif 'message' in req.keys():
            info = req['message']
            win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()
        else:
            print('请检查失败原因：', str(req))
            win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()

        time.sleep(1)
        # print('第二阶段请求-登录页面......')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
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
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        gimp = req.headers['Location']
        # print('+++已获取跳转页面： ' + gimp)
        time.sleep(1)
        # print('（二）请求dingtalk_service的cookie值......')
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index = req.headers['Location']

        time.sleep(1)
        # print('（三）请求dingtalk_service的cookie值......')
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=index, headers=r_header, allow_redirects=False)
        # print('+++已获取cookie值+++')

        time.sleep(1)
        # print('第四阶段页面-重定向跳转中......')
        # print('（一）加载chooselogin.html页面......')
        url = r'https://gsso.giikin.com/admin/login_by_dingtalk/chooselogin.html'
        data = {'user_id': self.userID}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': gimp,
                    'Origin': 'http://gsso.giikin.com'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        index = req.headers['Location']
        # print('+++已获取gimp.giikin.com页面')
        time.sleep(1)
        # print('（二）加载gimp.giikin.com页面......')
        url = index
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': index}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index2 = req.headers['Location']
        # print(99)
        # print(index2)
        # print('+++已获取index.html页面')

        # 跳转使用-暂停
        # index2 = index2.replace(':443/', '')
        # print(index2)
        # time.sleep(1)
        # url = index2
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': index2}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # index2 = req.headers['Location']
        # print(index2)
        # 跳转使用-暂停


        time.sleep(1)
        # print('（三）加载index.html页面......')
        url = 'https://gimp.giikin.com' + index2
        # url = 'https://gimp.giikin.com/portal/index/index.html'
        print(url)
        print(8080)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system = req.headers['Location']
        # print('+++已获取index.html?_system=18正式页面')
        # print(990008888888888888)

        time.sleep(1)
        # print('第五阶段正式页面-重定向跳转中......')
        # print('（一）加载index.html?_system页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system2 = req.headers['Location']
        # print('+++已获取index.html?_ticker=页面......')
        time.sleep(1)

        # 跳转使用-暂停
        # print('（二）加载index.html?_ticker=页面......')
        url = index_system2
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)
        index_system3 = req.headers['Location']
        print(808080)
        # print(index_system3)
        index_system3 = index_system3.replace(':443', '')
        print(index_system3)
        # 跳转使用-暂停

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = index_system3
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req)
        print(req.headers)
        print('++++++已成功登录++++++')
        print('*' * 50)
    def _online_Two(self):  # 登录系统保持会话状态
        print('*' * 50)
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171646622570440595157649661651738562272219G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        req = req.json()
        # print(req)
        # req_url = req['data']
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        if 'data' in req.keys():
            try:
                req_url = req['data']
                loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self._online()
        elif 'message' in req.keys():
            info = req['message']
            win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()
        else:
            print('请检查失败原因：', str(req))
            win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()
        # print('******已获取loginTmpCode值: ' + str(loginTmpCode))

        time.sleep(1)
        # print('二、请求-后台登录页面......')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.text)
        # print('******请求登录页面url成功： ' + str(req.text))

        time.sleep(1)
        # print('三、dingtalk_service服务器......')
        # print('（一）加载dingtalk_service跳转页面......')
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        gimp = req.headers['Location']
        # print('******已获取跳转页面： ' + str(gimp))
        time.sleep(1)
        # print('（二）请求dingtalk_service的cookie值......')
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        # index = req.headers['Location']
        # print(index)

        # time.sleep(1)
        # print('（三）请求dingtalk_service的cookie值......')
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=index, headers=r_header, allow_redirects=False)
        # print('+++已获取cookie值+++')

        # time.sleep(1)
        # print('第四阶段页面-重定向跳转中......')
        # print('（一）加载chooselogin.html页面......')
        # url = r'https://gsso.giikin.com/admin/login_by_dingtalk/chooselogin.html'
        # data = {'user_id': self.userID}
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': gimp,
        #             'Origin': 'http://gsso.giikin.com'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        # index = req.headers['Location']
        # print('+++已获取gimp.giikin.com页面')
        # time.sleep(1)
        # print('（二）加载gimp.giikin.com页面......')
        # url = index
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': index}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        # index2 = req.headers['Location']
        # print(99)
        # print(index2)
        # print('+++已获取index.html页面')

        # 跳转使用-暂停
        # index2 = index2.replace(':443/', '')
        # print(index2)
        # time.sleep(1)
        # url = index2
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': index2}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # index2 = req.headers['Location']
        # print(index2)
        # 跳转使用-暂停


        # time.sleep(1)
        # print('（三）加载index.html页面......')
        # url = 'https://gimp.giikin.com' + index
        # # url = 'https://gimp.giikin.com/portal/index/index.html'
        # print(url)
        # print(8080)
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        # index_system = req.headers['Location']
        # print('+++已获取index.html?_system=18正式页面')
        # print(990008888888888888)

        # time.sleep(1)
        # print('第五阶段正式页面-重定向跳转中......')
        # print('（一）加载index.html?_system页面......')
        # url = index_system
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        # index_system2 = req.headers['Location']
        # # print('+++已获取index.html?_ticker=页面......')
        # time.sleep(1)

        # 跳转使用-暂停
        # print('（二）加载index.html?_ticker=页面......')
        # url = index_system2
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)
        index_system3 = req.headers['Location']
        # print(808080)
        # print(index_system3)
        index_system3 = index_system3.replace(':443', '')
        # print(index_system3)
        # 跳转使用-暂停

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = index_system3
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # print(990099900)
        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('（三）加载index.html页面......')
        url = 'https://gimp.giikin.com' + index
        # url = 'https://gimp.giikin.com/portal/index/index.html'
        # print(url)
        # print(8080)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system = req.headers['Location']
        # print('+++已获取index.html?_system=18正式页面')
        # print(7070)

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(6060)
        # print(req)
        # print(req.headers)

        # time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        # url = req.headers['Location']
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++')
        print('*' * 50)
    def _online_Threed(self):  # 登录系统保持会话状态
        print('正在登录后台系统中......')
        # print('第一阶段获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171637544299948385570258461637545377418833G01447E2DCD07109775CD567044AE05FC09628C',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        req = req.json()
        # print(req)
        # req_url = req['data']
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        # print('+++已获取loginTmpCode值: ' + loginTmpCode)
        print(datetime.datetime.now())
        if 'data' in req.keys():
            try:
                req_url = req['data']
                loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self._online()
        elif 'message' in req.keys():
            info = req['message']
            win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()
        else:
            print('请检查失败原因：', str(req))
            win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()

        time.sleep(1)
        # print('第二阶段请求-登录页面......')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        url = req.text
        print(req.text)
        print(101)
        # print('+++请求登录页面url成功+++')

        time.sleep(1)
        # print('第三阶段请求-dingtalk服务器......')
        # print('（一）加载dingtalk_service跳转页面......')
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        gimp = req.headers['Location']
        # print('+++已获取跳转页面： ' + gimp)
        time.sleep(1)
        # print('（二）请求dingtalk_service的cookie值......')
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        index = req.headers['Location']
        # print(index)

        # time.sleep(1)
        # print('（三）请求dingtalk_service的cookie值......')
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=index, headers=r_header, allow_redirects=False)
        # print('+++已获取cookie值+++')

        # time.sleep(1)
        # print('第四阶段页面-重定向跳转中......')
        # print('（一）加载chooselogin.html页面......')
        # url = r'https://gsso.giikin.com/admin/login_by_dingtalk/chooselogin.html'
        # data = {'user_id': self.userID}
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': gimp,
        #             'Origin': 'http://gsso.giikin.com'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req.headers)
        # index = req.headers['Location']
        # print('+++已获取gimp.giikin.com页面')
        # time.sleep(1)
        # print('（二）加载gimp.giikin.com页面......')
        # url = index
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': index}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        # index2 = req.headers['Location']
        # print(99)
        # print(index2)
        # print('+++已获取index.html页面')

        # 跳转使用-暂停
        # index2 = index2.replace(':443/', '')
        # print(index2)
        # time.sleep(1)
        # url = index2
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': index2}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # index2 = req.headers['Location']
        # print(index2)
        # 跳转使用-暂停


        # time.sleep(1)
        # print('（三）加载index.html页面......')
        # url = 'https://gimp.giikin.com' + index
        # # url = 'https://gimp.giikin.com/portal/index/index.html'
        # print(url)
        # print(8080)
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        # index_system = req.headers['Location']
        # print('+++已获取index.html?_system=18正式页面')
        # print(990008888888888888)

        # time.sleep(1)
        # print('第五阶段正式页面-重定向跳转中......')
        # print('（一）加载index.html?_system页面......')
        # url = index_system
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        # index_system2 = req.headers['Location']
        # # print('+++已获取index.html?_ticker=页面......')
        # time.sleep(1)

        # 跳转使用-暂停
        # print('（二）加载index.html?_ticker=页面......')
        # url = index_system2
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)
        index_system3 = req.headers['Location']
        print(808080)
        # print(index_system3)
        index_system3 = index_system3.replace(':443', '')
        print(index_system3)
        # 跳转使用-暂停

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = index_system3
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req)
        print(req.headers)

        print(990099900)
        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index = req.headers['Location']
        print(req)
        print(req.headers)

        time.sleep(1)
        print('（三）加载index.html页面......')
        url = 'https://gimp.giikin.com' + index
        # url = 'https://gimp.giikin.com/portal/index/index.html'
        print(url)
        print(8080)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        index_system = req.headers['Location']
        print('+++已获取index.html?_system=18正式页面')
        print(7070)

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req)
        print(req.headers)

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(6060)
        print(req)
        print(req.headers)

        time.sleep(1)
        # print('（三）加载index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(5050)
        print(req)
        print(req.headers)

        print('++++++已成功登录++++++')
        print('*' * 50)

    # 不使用代理服务器
    def sso__online_handle(self, login_TmpCode):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        # url = r'https://login.dingtalk.com/login/login_with_pwd'
        # data = {'mobile': self.userMobile,
        #         'pwd': self.password,
        #         'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
        #         'pdmToken': '',
        #         'araAppkey': '1917',
        #         'araToken': '0#19171646622570440595157649661651738562272219G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
        #         'araScene': 'login',
        #         'captchaImgCode': '',
        #         'captchaSessionId': '',
        #         'type': 'h5'}
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #     'Origin': 'https://login.dingtalk.com',
        #     'Referer': 'https://login.dingtalk.com/'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # req = req.json()
        # print(req)
        # # req_url = req['data']
        # # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        # if 'data' in req.keys():
        #     try:
        #         req_url = req['data']
        #         loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
        #     except Exception as e:
        #         print('重新启动： 3分钟后', str(Exception) + str(e))
        #         time.sleep(300)
        #         self.sso_online_Two()
        # elif 'message' in req.keys():
        #     info = req['message']
        #     win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
        #     sys.exit()
        # else:
        #     print('请检查失败原因：', str(req))
        #     win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
        #     sys.exit()
        # print('******已获取loginTmpCode值: ' + str(loginTmpCode))

        loginTmpCode = self.login_TmpCode
        print('1、加载： ' + 'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.text)
        # print(req.headers)
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        print('2、加载： ' + 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?')
        url = 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode&loginTmpCode=' + loginTmpCode
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('4.2、加载： ' + 'https://gimp.giikin.com')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        print(req.headers)

        # time.sleep(1)
        # print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
        # url = req.headers['Location']
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)
    # 使用代理服务器
    def sso__online_handle_proxy(self, login_TmpCode, proxy_id):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        # url = r'https://login.dingtalk.com/login/login_with_pwd'
        # data = {'mobile': self.userMobile,
        #         'pwd': self.password,
        #         'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
        #         'pdmToken': '',
        #         'araAppkey': '1917',
        #         'araToken': '0#19171646622570440595157649661651738562272219G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
        #         'araScene': 'login',
        #         'captchaImgCode': '',
        #         'captchaSessionId': '',
        #         'type': 'h5'}
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #     'Origin': 'https://login.dingtalk.com',
        #     'Referer': 'https://login.dingtalk.com/'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # req = req.json()
        # print(req)
        # # req_url = req['data']
        # # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        # if 'data' in req.keys():
        #     try:
        #         req_url = req['data']
        #         loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
        #     except Exception as e:
        #         print('重新启动： 3分钟后', str(Exception) + str(e))
        #         time.sleep(300)
        #         self.sso_online_Two()
        # elif 'message' in req.keys():
        #     info = req['message']
        #     win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
        #     sys.exit()
        # else:
        #     print('请检查失败原因：', str(req))
        #     win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
        #     sys.exit()
        # print('******已获取loginTmpCode值: ' + str(loginTmpCode))

        loginTmpCode = self.login_TmpCode
        print('1、加载： ' + 'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}            # 使用代理服务器
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)

        # print(req)
        # print(req.text)
        # print(req.headers)
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        print('2、加载： ' + 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?')
        url = 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode&loginTmpCode=' + loginTmpCode
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('4.2、加载： ' + 'https://gimp.giikin.com')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        print(req.headers)

        # time.sleep(1)
        # print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
        # url = req.headers['Location']
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

    def sso__online_handle_(self, login_TmpCode):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        # url = r'https://login.dingtalk.com/login/login_with_pwd'
        # data = {'mobile': self.userMobile,
        #         'pwd': self.password,
        #         'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
        #         'pdmToken': '',
        #         'araAppkey': '1917',
        #         'araToken': '0#19171646622570440595157649661651738562272219G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
        #         'araScene': 'login',
        #         'captchaImgCode': '',
        #         'captchaSessionId': '',
        #         'type': 'h5'}
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #     'Origin': 'https://login.dingtalk.com',
        #     'Referer': 'https://login.dingtalk.com/'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # req = req.json()
        # print(req)
        # # req_url = req['data']
        # # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        # if 'data' in req.keys():
        #     try:
        #         req_url = req['data']
        #         loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
        #     except Exception as e:
        #         print('重新启动： 3分钟后', str(Exception) + str(e))
        #         time.sleep(300)
        #         self.sso_online_Two()
        # elif 'message' in req.keys():
        #     info = req['message']
        #     win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
        #     sys.exit()
        # else:
        #     print('请检查失败原因：', str(req))
        #     win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
        #     sys.exit()
        # print('******已获取loginTmpCode值: ' + str(loginTmpCode))

        loginTmpCode = self.login_TmpCode
        print('1、加载： ' + 'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.text)
        # print(req.headers)
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        print('2、加载： ' + 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?')
        url = 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode&loginTmpCode=' + loginTmpCode
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('4.2、加载： ' + 'https://gimp.giikin.com')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        print(req.headers)

        # time.sleep(1)
        # print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
        # url = req.headers['Location']
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

    def sso__online_auto_host(self):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': "+86-18538110674",
                'pwd': "qyz04163510.",
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171653727303200791658056521658928039705627G724AFE4E3392F35159AB9A341B2E56DAA31033',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0',
            'Origin': 'https://login.dingtalk.com',
            'Referer': 'https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        req = req.json()
        # req = {}
        # print(req)
        # req_url = req['data']  0#19171629428116275265671469741656903392035557GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        login_TmpCode = '获取不到参数'
        if 'data' in req.keys():
            try:
                req_url = req['data']
                login_TmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self.sso__online_auto_host()
        elif 'message' in req.keys():
            info = req['message']
            win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
            # sys.exit()
        else:
            print('请检查失败原因：', str(req))
            # win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
            # sys.exit()

        if login_TmpCode == '获取不到参数':
            time.sleep(1)
            # 模拟打开火狐浏览器 获取token
            options = Options()
            options.add_argument('-headless')
            driver = webdriver.Firefox(options=options)
            # driver.get('https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            driver.implicitly_wait(5)
            js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
                        data: { mobile: '+86-18538110674',
                                pwd: 'qyz04163510.',
                                goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                                pdmToken: '',
                                araAppkey: '1917',
                                araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                                araScene: 'login',
                                captchaImgCode: '',
                                captchaSessionId: '',
                                type: 'h5'
                            },
                            type: 'POST',
                            timeout: '10000',
                            async:false,
                            beforeSend(xhr, settings) {
                                xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                            },
                            success: function(data) {
                                if (data.success) {
                                     console.log(data.data);
                                     console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
                                      document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
                                     arguments[0].value=data.data.split('loginTmpCode=')[1];
                                } else {
                                        console.log(data.code);
                                }
                            },
                            error: function(error) {
                                alert("请检查网络");
                            }
                        });
                        '''
            element = driver.find_element('id', 'mobile')
            driver.execute_script(js, element)
            # driver.implicitly_wait(5)
            time.sleep(5)
            login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
            print('loginTmpCode值: ' + login_TmpCode)
            driver.quit()

        elif login_TmpCode == '获取不到参数':
            time.sleep(1)
            # 模拟打开谷歌浏览器 获取token
            options = webdriver.ChromeOptions()
            options.add_argument(r"user-data-dir=C:\Program Files\Google\Chrome\Application\profile")
            driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')

            driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            # driver.implicitly_wait(5)
            time.sleep(5)
            js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
                        data: { mobile: '+86-18538110674',
                                pwd: 'qyz04163510.',
                                goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                                pdmToken: '',
                                araAppkey: '1917',
                                araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                                araScene: 'login',
                                captchaImgCode: '',
                                captchaSessionId: '',
                                type: 'h5'
                            },
                            type: 'POST',
                            timeout: '10000',
                            async:false,
                            beforeSend(xhr, settings) {
                                xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                            },
                            success: function(data) {
                                if (data.success) {
                                     console.log(data.data);
                                     console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
                                      document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
                                     arguments[0].value=data.data.split('loginTmpCode=')[1];
                                } else {
                                        console.log(data.code);
                                }
                            },
                            error: function(error) {
                                alert("请检查网络");
                            }
                        });
                        '''
            element = driver.find_element('id', 'mobile')
            driver.execute_script(js, element)
            # driver.implicitly_wait(5)
            time.sleep(5)
            login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
            print('loginTmpCode值: ' + login_TmpCode)
            driver.quit()



        print('******已获取loginTmpCode值: ' + str(login_TmpCode))
        loginTmpCode = login_TmpCode
        # loginTmpCode = 'af8203b900ce347287492b0051fe1e11'
        # print('1、加载： ' + 'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.text)
        # print(req.headers)
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('2、加载： ' + 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?')
        url = 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode&loginTmpCode=' + loginTmpCode
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)
        url = req.headers['Location']

        time.sleep(1)
        if url != '/portal/index/index.html':
            print('4.2、加载： ' + 'https://gimp.giikin.com')
            url = req.headers['Location']
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            index = req.headers['Location']
            print(req)
            print(req.headers)
        else:
            index = req.headers['Location']

        time.sleep(1)
        # print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # time.sleep(1)
        # print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
        # url = req.headers['Location']
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++++++ ' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

    # 不使用代理服务器
    def sso__online_auto_gp(self):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': '+86-18538110674',
                'pwd': 'qyz04163510.',
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171629428116275265671469741658739612489317GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:94.0) Gecko/20100101 Firefox/94.0',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # req = req.json()
        req = {}
        # print(req)
        # req_url = req['data']  0#19171629428116275265671469741656903392035557GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        login_TmpCode = '获取不到参数'
        if 'data' in req.keys():
            try:
                req_url = req['data']
                login_TmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self.sso__online_auto_gp()
        elif 'message' in req.keys():
            info = req['message']
            win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
            # sys.exit()
        else:
            print('请检查失败原因：', str(req))
            # win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
            # sys.exit()
        if login_TmpCode == '获取不到参数':
            time.sleep(1)
            # 模拟打开火狐浏览器 获取token
            options = Options()
            options.add_argument('-headless')
            driver = webdriver.Firefox(options=options)
            # driver.get('https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            driver.implicitly_wait(5)
            js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
                        data: { mobile: '+86-18538110674',
                                pwd: 'qyz04163510.',
                                goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                                pdmToken: '',
                                araAppkey: '1917',
                                araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                                araScene: 'login',
                                captchaImgCode: '',
                                captchaSessionId: '',
                                type: 'h5'
                            },
                            type: 'POST',
                            timeout: '10000',
                            async:false,
                            beforeSend(xhr, settings) {
                                xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                            },
                            success: function(data) {
                                if (data.success) {
                                     console.log(data.data);
                                     console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
                                      document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
                                     arguments[0].value=data.data.split('loginTmpCode=')[1];
                                } else {
                                        console.log(data.code);
                                }
                            },
                            error: function(error) {
                                alert("请检查网络");
                            }
                        });
                        '''
            element = driver.find_element('id', 'mobile')
            driver.execute_script(js, element)
            # driver.implicitly_wait(5)
            time.sleep(5)
            login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
            print('loginTmpCode值: ' + login_TmpCode)
            driver.quit()

        elif login_TmpCode == '获取不到参数':
            time.sleep(1)
            # 模拟打开谷歌浏览器 获取token
            options = webdriver.ChromeOptions()
            options.add_argument(r"user-data-dir=C:\Program Files\Google\Chrome\Application\profile")
            driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')

            driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            # driver.implicitly_wait(5)
            time.sleep(5)
            js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
                        data: { mobile: '+86-18538110674',
                                pwd: 'qyz04163510.',
                                goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                                pdmToken: '',
                                araAppkey: '1917',
                                araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                                araScene: 'login',
                                captchaImgCode: '',
                                captchaSessionId: '',
                                type: 'h5'
                            },
                            type: 'POST',
                            timeout: '10000',
                            async:false,
                            beforeSend(xhr, settings) {
                                xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                            },
                            success: function(data) {
                                if (data.success) {
                                     console.log(data.data);
                                     console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
                                      document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
                                     arguments[0].value=data.data.split('loginTmpCode=')[1];
                                } else {
                                        console.log(data.code);
                                }
                            },
                            error: function(error) {
                                alert("请检查网络");
                            }
                        });
                        '''
            element = driver.find_element('id', 'mobile')
            driver.execute_script(js, element)
            # driver.implicitly_wait(5)
            time.sleep(5)
            login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
            print('loginTmpCode值: ' + login_TmpCode)
            driver.quit()

        # print('******已获取loginTmpCode值: ' + str(login_TmpCode))
        loginTmpCode = login_TmpCode
        # loginTmpCode = 'af8203b900ce347287492b0051fe1e11'
        # print('1、加载： ' + 'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.text)
        # print(req.headers)
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('2、加载： ' + 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?')
        url = 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode&loginTmpCode=' + loginTmpCode
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)
        url = req.headers['Location']

        '''
        time.sleep(1)
        if url != '/portal/index/index.html':
            # print('4.2、加载： ' + 'https://gimp.giikin.com')
            url = req.headers['Location']
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            index = req.headers['Location']
            # print(req)
            # print(req.headers)
        else:
            index = req.headers['Location']
        time.sleep(1)
        # print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        '''

        # 此处跳转换新的地址了
        # print('5.1、加载： ' + 'https://gimp.giikin.com//admin/login_by_dingtalk/finishLoginJump?jump_url=https://gimp.giikin.com')
        time.sleep(1)
        if '/admin/login_by_dingtalk/finishLoginJump?jump_url=https://gimp.giikin.com' in url:
            index = req.headers['Location']
            url = 'https://gsso.giikin.com' + index
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)

        elif url == '/portal/index/index.html':
            index = req.headers['Location']
            url = 'https://gimp.giikin.com' + index
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)

        else:
            url = req.headers['Location']
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            index = req.headers['Location']
            url = 'https://gimp.giikin.com' + index
            # print(req)
            # print(req.headers)

        time.sleep(1)
        # print('5.2、加载： ' + 'https://gimp.giikin.com')
        # print(url)
        url = 'https://gimp.giikin.com/'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('5.3、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        # print(url)
        index = req.headers['Location']
        url = 'https://gimp.giikin.com/' + index
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # time.sleep(1)
        # print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
        # url = req.headers['Location']
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++++++ ' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)
    # 使用代理服务器
    def sso__online_auto_gp_proxy(self, proxy_id):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': '+86-18538110674',
                'pwd': 'qyz04163510.',
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171629428116275265671469741658739612489317GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:94.0) Gecko/20100101 Firefox/94.0',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # req = req.json()
        req = {}
        # print(req)
        # req_url = req['data']  0#19171629428116275265671469741656903392035557GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        login_TmpCode = '获取不到参数'
        if 'data' in req.keys():
            try:
                req_url = req['data']
                login_TmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self.sso__online_auto_gp_proxy(proxy_id)
        elif 'message' in req.keys():
            info = req['message']
            win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
            # sys.exit()
        else:
            print('请检查失败原因：', str(req))
            # win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
            # sys.exit()
        if login_TmpCode == '获取不到参数':
            time.sleep(1)
            # 模拟打开火狐浏览器 获取token
            options = Options()
            options.add_argument('-headless')
            driver = webdriver.Firefox(options=options)
            # driver.get('https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            driver.implicitly_wait(5)
            js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
                        data: { mobile: '+86-18538110674',
                                pwd: 'qyz04163510.',
                                goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                                pdmToken: '',
                                araAppkey: '1917',
                                araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                                araScene: 'login',
                                captchaImgCode: '',
                                captchaSessionId: '',
                                type: 'h5'
                            },
                            type: 'POST',
                            timeout: '10000',
                            async:false,
                            beforeSend(xhr, settings) {
                                xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                            },
                            success: function(data) {
                                if (data.success) {
                                     console.log(data.data);
                                     console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
                                      document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
                                     arguments[0].value=data.data.split('loginTmpCode=')[1];
                                } else {
                                        console.log(data.code);
                                }
                            },
                            error: function(error) {
                                alert("请检查网络");
                            }
                        });
                        '''
            element = driver.find_element('id', 'mobile')
            driver.execute_script(js, element)
            # driver.implicitly_wait(5)
            time.sleep(5)
            login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
            print('loginTmpCode值: ' + login_TmpCode)
            driver.quit()

        elif login_TmpCode == '获取不到参数':
            time.sleep(1)
            # 模拟打开谷歌浏览器 获取token
            options = webdriver.ChromeOptions()
            options.add_argument(r"user-data-dir=C:\Program Files\Google\Chrome\Application\profile")
            driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')

            driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
            # driver.implicitly_wait(5)
            time.sleep(5)
            js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
                        data: { mobile: '+86-18538110674',
                                pwd: 'qyz04163510.',
                                goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                                pdmToken: '',
                                araAppkey: '1917',
                                araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                                araScene: 'login',
                                captchaImgCode: '',
                                captchaSessionId: '',
                                type: 'h5'
                            },
                            type: 'POST',
                            timeout: '10000',
                            async:false,
                            beforeSend(xhr, settings) {
                                xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                            },
                            success: function(data) {
                                if (data.success) {
                                     console.log(data.data);
                                     console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
                                      document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
                                     arguments[0].value=data.data.split('loginTmpCode=')[1];
                                } else {
                                        console.log(data.code);
                                }
                            },
                            error: function(error) {
                                alert("请检查网络");
                            }
                        });
                        '''
            element = driver.find_element('id', 'mobile')
            driver.execute_script(js, element)
            # driver.implicitly_wait(5)
            time.sleep(5)
            login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
            print('loginTmpCode值: ' + login_TmpCode)
            driver.quit()

        # print('******已获取loginTmpCode值: ' + str(login_TmpCode))
        loginTmpCode = login_TmpCode
        # loginTmpCode = 'af8203b900ce347287492b0051fe1e11'
        # print('1、加载： ' + 'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html')
        url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
        data = {'tmpCode': loginTmpCode,
                'system': 18,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.text)
        # print(req.headers)
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('2、加载： ' + 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?')
        url = 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode&loginTmpCode=' + loginTmpCode
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        # print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)
        url = req.headers['Location']


        '''
        time.sleep(1)
        if url != '/portal/index/index.html':
            # print('4.2、加载： ' + 'https://gimp.giikin.com')
            url = req.headers['Location']
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            index = req.headers['Location']
            # print(req)
            # print(req.headers)
        else:
            index = req.headers['Location']
        time.sleep(1)
        # print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        '''

        # 此处跳转换新的地址了
        # print('5.1、加载： ' + 'https://gimp.giikin.com//admin/login_by_dingtalk/finishLoginJump?jump_url=https://gimp.giikin.com')
        time.sleep(1)
        if '/admin/login_by_dingtalk/finishLoginJump?jump_url=https://gimp.giikin.com' in url:
            index = req.headers['Location']
            url = 'https://gsso.giikin.com' + index
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)

        elif url == '/portal/index/index.html':
            index = req.headers['Location']
            url = 'https://gimp.giikin.com' + index
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)

        else:
            url = req.headers['Location']
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            index = req.headers['Location']
            url = 'https://gimp.giikin.com' + index
            # print(req)
            # print(req.headers)

        time.sleep(1)
        # print('5.2、加载： ' + 'https://gimp.giikin.com')
        # print(url)
        url = 'https://gimp.giikin.com/'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('5.3、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        # print(url)
        index = req.headers['Location']
        url = 'https://gimp.giikin.com/' + index
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)


        time.sleep(1)
        # print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        # time.sleep(1)
        # print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
        # url = req.headers['Location']
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #             'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++++++ ' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)


    # 获取签收表内容
    def readFormHost(self, team, query):
        match3 = {'新加坡': 'slxmt',
                  '马来西亚': 'slxmt',
                  '菲律宾': 'slxmt',
                  '新马': 'slxmt',
                  '日本': 'slrb',
                  '香港': 'slgat',
                  '台湾': 'slgat',
                  '港台': 'slgat',
                  '泰国': 'sltg'}
        start = datetime.datetime.now()
        if team == 'slsc':
            path = r'F:\需要用到的文件\品牌数据源'
        else:
            if query == '导入':
                path = r'F:\需要用到的文件\数据库导入'
            else:
                path = r'F:\需要用到的文件\数据库'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath, team, query)
                # os.remove(filePath)
                print('已清除上传文件…………')
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath, team, query):
        match2 = {'slgat': '神龙港台',
                  'slgat_hfh': '火凤凰港台',
                  'slgat_hs': '红杉港台',
                  'slsc': '品牌',
                  'gat': '港台',
                  'sltg': '泰国',
                  'slxmt': '新马',
                  'slxmt_t': 'T新马',
                  'slxmt_hfh': '火凤凰新马',
                  'slrb': '日本',
                  'slrb_js': '金狮-日本',
                  'slrb_hs': '红杉-日本',
                  'slrb_jl': '精灵-日本'}
        print('---正在获取 ' + match2[team] + ' 签收表的详情++++++')
        fileType = os.path.splitext(filePath)[1]
        app = xlwings.App(visible=False, add_book=False)
        app.display_alerts = False
        if 'xls' in fileType:
            wb = app.books.open(filePath, update_links=False, read_only=True)
            for sht in wb.sheets:
                try:
                    db = None
                    db = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                    db.rename(columns={'规格(中文)': '规格中文'}, inplace=True)
                    columns = list(db.columns)  # 获取数据的标题名，转为列表
                    # print(columns)
                    columns_value = ['商品链接', '收货人', '电话长度', '邮编长度', '配送地址', '地址翻译',
                                     '邮箱', '留言', '', '预选物流公司(新)', '是否api下单', '特价', '市 / 区', '是否发送邮件',
                                     '备注', '是否发送短信', '商品合计', '超商店铺ID', '超商店铺名称', '超商网点地址', '超商类型',
                                     '市/区', '优化师', '纬经度', '通关码', '站点域名', '品名英文',
                                     '简站', '地址类型']
                    for column_val in columns_value:
                        if column_val in columns:
                            db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                    # db = db[['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '联系电话', '拉黑率', '邮编','是否低价',
                    #          '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型', '商品分类', '商品ID',
                    #          '订单类型', '物流状态', '异常提示', '重量', '删除原因', '父级分类', '转采购时间', '发货时间', '收货时间', '上线时间', '完成时间',
                    #          '删除人', 'IP', '体积', '省洲', '运输方式', '优化师', '问题原因', '审单类型', '代下单客服', '改派的下架时间', '克隆人','站点ID',
                    #          '选品人', '设计师', '运费', '服务费', '拒收原因', '克隆时间']]
                    db['运单号'] = db['运单号'].str.strip()                     # 去掉运单号中的前后空字符串
                    db['物流渠道'] = db['物流渠道'].str.strip()
                    db['产品名称'] = db['产品名称'].str.split('#', expand=True)[1]
                    db['站点ID'] = db['站点ID'].str.strip()
                    print(db.columns)
                except Exception as e:
                    print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                if db is not None and len(db) > 0:
                    print('++++正在导入更新：' + sht.name + ' 共：' + str(len(db)) + '行',
                          'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                    # 将返回的dateFrame导入数据库的临时表
                    self.writeCacheHost(db)
                    print('++++正在更新：' + sht.name + '--->>>到总订单')
                    # 将数据库的临时表替换进指定的总表
                    self.replaceSqlHost(team, query)
                    print('++++----->>>' + sht.name + '：订单更新完成++++')
                else:
                    print('----------数据为空导入失败：' + sht.name)
            wb.close()
        app.quit()

    # 写入临时缓存表
    def writeCacheHost(self, dataFrame):
        dataFrame.to_sql('d1_host', con=self.engine1, index=False, if_exists='replace')
    # 写入总表
    def replaceSqlHost(self, team, query):
        if team in ('gat', 'slgat', 'slgat_hfh', 'slgat_hs'):
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
            				        IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '11', '1', '2')) 旬,
            			            DATE(h.下单时间) 日期,
            				        h.运营团队 团队,
            				        IF(h.`币种` = '台币', 'TW', IF(h.`币种` = '港币', 'HK', h.`币种`)) 区域,
            				        IF(h.`币种` = '台币', '台湾', IF(h.`币种` = '港币', '香港', h.`币种`)) 币种,
            				        h.平台 订单来源,
            				        订单编号,
            				        数量,
            				        h.联系电话 电话号码,
            				        h.运单号 运单编号,
            				        IF(h.物流渠道 LIKE "台湾-天马-711" AND LENGTH(h.运单号)=20, CONCAT(861,RIGHT(h.运单号,8)), IF((h.物流渠道 LIKE "台湾-速派-新竹改派" or h.物流渠道 LIKE "台湾-易速配-新竹改派") AND (h.运单号 LIKE "A%" OR h.运单号 LIKE "B%"),RIGHT(h.运单号,LENGTH(h.运单号)-1),UPPER(h.运单号))) 查件单号,
            				        h.订单状态 系统订单状态,
            				        IF(h.`物流状态` in ('发货中'), null, h.`物流状态`) 系统物流状态,
            				        IF(h.`订单类型` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
            				        h.物流渠道 物流方式,
            				        dim_trans_way.simple_name 物流名称,
            				        dim_trans_way.remark 运输方式,
            				        IF(h.`货物类型` = 'P 普通货', 'P', IF(h.`货物类型` = 'T 特殊货', 'T', h.`货物类型`)) 货物类型,
            				        是否低价,
            				        商品ID,
            				        产品id,
            				        产品名称,
            				        dim_cate.ppname 父级分类,
            				        dim_cate.pname 二级分类,
                		            dim_cate.`name` 三级分类,
            				        h.支付方式 付款方式,
            				        h.应付金额 价格,
            				        下单时间,
            				        审核时间,
            				        h.发货时间 仓储扫描时间,
            				        null 完结状态,
            				        h.完成时间 完结状态时间,
            				        null 价格RMB,
            				        null 价格区间,
            				        null 成本价,
            				        null 物流花费,
            				        null 打包花费,
            				        null 其它花费,
            				        h.重量 包裹重量,
            				        h.体积 包裹体积,
            				        邮编,
            				        h.转采购时间 添加物流单号时间,
            				        h.规格中文,
            				        h.省洲 省洲,
            				        IF(h.审单类型 like '%自动审单%','是','否') 审单类型,
            				        h.审单类型 审单类型明细,
            				        h.拉黑率,
            				        null 订单配送总量,
            				        null 签收量,
            				        null 拒收量,
            				        h.删除原因,
            				        null 删除时间,
            				        h.问题原因,
            				        null 问题时间,
            				        h.代下单客服 下单人,
            				        h.克隆人,
            				        null 下架类型,
            				        h.改派的下架时间 下架时间,
            				        h.收货时间 物流提货时间,
            				        null 物流发货时间,
            				        h.上线时间 上线时间,
            				        null 国内清关时间,
            				        null 目的清关时间,
            				        null 回款时间,
                                    h.IP,
            				        h.选品人
                            FROM d1_host h 
                            LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                            LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                            LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`; '''.format(team)
        elif team in ('slsc'):
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
			                    IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '11', '1', '2')) 旬,
			                    DATE(h.下单时间) 日期,
				                h.运营团队 团队,
				                IF(h.`币种` = '日币', 'JP', IF(h.`币种` = '菲律宾', 'PH', IF(h.`币种` = '新加坡', 'SG', IF(h.`币种` = '马来西亚', 'MY', IF(h.`币种` = '台币', 'TW', h.`币种`))))) 区域,
				                IF(h.`币种` = '日币', '日本', IF(h.`币种` = '菲律宾', '菲律宾', IF(h.`币种` = '新加坡', '新加坡', IF(h.`币种` = '马来西亚', '马来西亚', IF(h.`币种` = '台币', '台湾', h.`币种`))))) 币种,
				                h.平台 订单来源,
				                订单编号,
				                数量,
				                h.联系电话 电话号码,
				                h.运单号 运单编号,
				                IF(h.`订单类型` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
				                h.物流渠道 物流方式,
			--	                IF(h.`物流渠道` LIKE '%捷浩通%', '捷浩通', IF(h.`物流渠道` LIKE '%翼通达%','翼通达', IF(h.`物流渠道` LIKE '%博佳图%', '博佳图', IF(h.`物流渠道` LIKE '%保辉达%', '保辉达物流', IF(h.`物流渠道` LIKE '%万立德%','万立德', h.`物流渠道`))))) 物流名称,
				                dim_trans_way.simple_name 物流名称,
				                dim_trans_way.remark 运输方式,
				                IF(h.`货物类型` = 'P 普通货', 'P', IF(h.`货物类型` = 'T 特殊货', 'T', h.`货物类型`)) 货物类型,
				                是否低价,
				                产品id,
				                产品名称,
				                dim_cate.ppname 父级分类,
				                dim_cate.pname 二级分类,
    		                    dim_cate.`name` 三级分类,
				                IF(h.支付方式 = '货到付款' ,'货到付款' , '在线') 付款方式,
				                h.应付金额 价格,
				                下单时间,
				                审核时间,
				                h.发货时间 仓储扫描时间,
				                null 完结状态,
				                h.完成时间 完结状态时间,
				                null 价格RMB,
				                null 价格区间,
				                null 成本价,
				                null 物流花费,
				                null 打包花费,
				                null 其它花费,
				                h.重量 包裹重量,
				                h.体积 包裹体积,
				                邮编,
				                h.转采购时间 添加物流单号时间,
				                IF(h.运营团队 = '精灵家族-品牌',IF(h.站点ID=1000000269,'饰品','内衣'),h.站点ID) 站点ID,
				                null 订单删除原因,
				                h.订单状态 系统订单状态,
				                IF(h.`物流状态` in ('发货中'), null, h.`物流状态`) 系统物流状态,
            				    h.上线时间 上线时间
                    FROM d1_host h 
                    LEFT JOIN dim_product_slsc ON  dim_product_slsc.id = h.产品id
            --        LEFT JOIN (SELECT * FROM dim_product WHERE id IN (SELECT MAX(id) FROM dim_product GROUP BY id ) ORDER BY id) e on e.id = h.产品id
                    LEFT JOIN dim_cate ON  dim_cate.id = dim_product_slsc.third_cate_id
                    LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team in ('slrb_jl', 'slrb_js', 'slrb_hs'):
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
			                    IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '11', '1', '2')) 旬,
			                    DATE(h.下单时间) 日期,
				                h.运营团队 团队,
				                IF(h.`币种` = '日币', 'JP', h.`币种`) 区域,
				                IF(h.`币种` = '日币', '日本', h.`币种`) 币种,
				                h.平台 订单来源,
				                订单编号,
				                数量,
				                h.联系电话 电话号码,
				                h.运单号 运单编号,
				                IF(h.`订单类型` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
				                h.物流渠道 物流方式,
			--	                IF(h.`物流渠道` LIKE '%捷浩通%', '捷浩通', IF(h.`物流渠道` LIKE '%翼通达%','翼通达', IF(h.`物流渠道` LIKE '%博佳图%', '博佳图', IF(h.`物流渠道` LIKE '%保辉达%', '保辉达物流', IF(h.`物流渠道` LIKE '%万立德%','万立德', h.`物流渠道`))))) 物流名称,
				                dim_trans_way.simple_name 物流名称,
				                dim_trans_way.remark 运输方式,
				                IF(h.`货物类型` = 'P 普通货', 'P', IF(h.`货物类型` = 'T 特殊货', 'T', h.`货物类型`)) 货物类型,
				                是否低价,
				                产品id,
				                产品名称,
				                dim_cate.ppname 父级分类,
				                dim_cate.pname 二级分类,
    		                    dim_cate.`name` 三级分类,
				                h.支付方式 付款方式,
				                h.应付金额 价格,
				                下单时间,
				                审核时间,
				                h.发货时间 仓储扫描时间,
				                null 完结状态,
				                h.完成时间 完结状态时间,
				                null 价格RMB,
				                null 价格区间,
				                null 成本价,
				                null 物流花费,
				                null 打包花费,
				                null 其它花费,
				                h.重量 包裹重量,
				                h.体积 包裹体积,
				                邮编,
				                h.转采购时间 添加物流单号时间,
				                IF(h.运营团队 = '精灵家族-品牌',IF(h.站点ID=1000000269,'饰品','内衣'),h.站点ID) 站点ID,
				                null 订单删除原因,
				                h.订单状态 系统订单状态,
				                IF(h.`物流状态` in ('发货中'), null, h.`物流状态`) 系统物流状态,
            				    h.上线时间 上线时间
                    FROM d1_host h 
                    LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                    LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                    LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team == 'slrb':
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
        			                    IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '11', '1', '2')) 旬,
        			                    DATE(h.下单时间) 日期,
        				                h.运营团队 团队,
        				                IF(h.`币种` = '日币', 'JP', h.`币种`) 区域,
        				                IF(h.`币种` = '日币', '日本', h.`币种`) 币种,
        				                h.平台 订单来源,
        				                订单编号,
        				                数量,
        				                h.联系电话 电话号码,
        				                h.运单号 运单编号,
        				                IF(h.`订单类型` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
        				                h.物流渠道 物流方式,
        			--	                IF(h.`物流渠道` LIKE '%捷浩通%', '捷浩通', IF(h.`物流渠道` LIKE '%翼通达%','翼通达', IF(h.`物流渠道` LIKE '%博佳图%', '博佳图', IF(h.`物流渠道` LIKE '%保辉达%', '保辉达物流', IF(h.`物流渠道` LIKE '%万立德%','万立德', h.`物流渠道`))))) 物流名称,
        				                dim_trans_way.simple_name 物流名称,
        				                dim_trans_way.remark 运输方式,
        				                IF(h.`货物类型` = 'P 普通货', 'P', IF(h.`货物类型` = 'T 特殊货', 'T', h.`货物类型`)) 货物类型,
        				                是否低价,
        				                产品id,
        				                产品名称,
        				                dim_cate.ppname 父级分类,
        				                dim_cate.pname 二级分类,
            		                    dim_cate.`name` 三级分类,
        				                h.支付方式 付款方式,
        				                h.应付金额 价格,
        				                下单时间,
        				                审核时间,
        				                h.发货时间 仓储扫描时间,
        				                null 完结状态,
        				                h.完成时间 完结状态时间,
        				                null 价格RMB,
        				                null 价格区间,
        				                null 成本价,
        				                null 物流花费,
        				                null 打包花费,
        				                null 其它花费,
        				                h.重量 包裹重量,
        				                h.体积 包裹体积,
        				                邮编,
        				                h.转采购时间 添加物流单号时间,
        				                null 订单删除原因,
        				                h.订单状态 系统订单状态,
        				                IF(h.`物流状态` in ('发货中'), null, h.`物流状态`) 系统物流状态,
            				            h.上线时间 上线时间 
                            FROM d1_host h 
                            LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                            LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                            LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team == 'sltg':
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
                                IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '11', '1', '2')) 旬,
			                    DATE(h.下单时间) 日期,
				                h.运营团队 团队,
				                IF(h.`币种` = '泰铢', 'TH', h.`币种`) 区域,
				                IF(h.`币种` = '泰铢', '泰国', h.`币种`) 币种,
				                h.平台 订单来源,
				                订单编号,
				                数量,
				                h.联系电话 电话号码,
				                h.运单号 运单编号,
				                IF(h.`订单类型` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
				                h.物流渠道 物流方式,
                                dim_trans_way.simple_name 物流名称,
				                dim_trans_way.remark 运输方式,
				                IF(h.`货物类型` = 'P 普通货', 'P', IF(h.`货物类型` = 'T 特殊货', 'T', h.`货物类型`)) 货物类型,
				                是否低价,
				                产品id,
				                产品名称,
				                dim_cate.ppname 父级分类,
				                dim_cate.pname 二级分类,
    		                    dim_cate.`name` 三级分类,
				                h.支付方式 付款方式,
				                h.应付金额 价格,
				                下单时间,
				                审核时间,
				                h.发货时间 仓储扫描时间,
				                null 完结状态,
				                h.完成时间 完结状态时间,
				                null 价格RMB,
				                null 价格区间,
				                null 成本价,
				                null 物流花费,
				                null 打包花费,
				                null 其它花费,
				                h.重量 包裹重量,
				                h.体积 包裹体积,
				                邮编,
				                h.转采购时间 添加物流单号时间,
				                null 订单删除原因,
				                h.订单状态 系统订单状态,
				                IF(h.`物流状态` in ('发货中'), null, h.`物流状态`) 系统物流状态,
            				    h.上线时间 上线时间
                    FROM d1_host h 
                    LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                    LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                    LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team in ('slxmt', 'slxmt_t', 'slxmt_hfh'):
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
                            IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '11', '1', '2')) 旬,
                            DATE(h.下单时间) 日期,
                            h.运营团队 团队,
                            IF(h.`币种` = '马来西亚', 'MY', IF(h.`币种` ='菲律宾', 'PH', IF(h.`币种` = '新加坡', 'SG', null))) 区域,
                            币种,
                            h.平台 订单来源,
                            订单编号,
                            数量,
                            h.联系电话 电话号码,
                            h.运单号 运单编号,
                            IF(h.`订单类型` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
                            h.物流渠道 物流方式,
                            dim_trans_way.simple_name 物流名称,
                            dim_trans_way.remark 运输方式,
                            IF(h.`货物类型` = 'P 普通货', 'P', IF(h.`货物类型` = 'T 特殊货', 'T', h.`货物类型`)) 货物类型,
                            是否低价,
                            产品id,
                            产品名称,
                            dim_cate.ppname 父级分类,
                            dim_cate.pname 二级分类,
                            dim_cate.`name` 三级分类,
                            h.支付方式 付款方式,
                            h.应付金额 价格,
                            下单时间,
                            审核时间,
                            h.发货时间 仓储扫描时间,
                            null 完结状态,
                            h.完成时间 完结状态时间,
                            null 价格RMB,
                            null 价格区间,
                            null 成本价,
                            null 物流花费,
                            null 打包花费,
                            null 其它花费,
                            h.重量 包裹重量,
                            h.体积 包裹体积,
                            邮编,
                            h.转采购时间 添加物流单号时间,
                            null 订单删除原因,
                            h.省洲 省洲,
                            h.订单状态 系统订单状态,
                            IF(h.`物流状态` in ('发货中'), null, h.`物流状态`) 系统物流状态,
            				h.上线时间 上线时间
                        FROM d1_host h 
                            LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                            LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                            LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        if query == '导入':
            try:
                print('正在导入临时表中......')
                df = pd.read_sql_query(sql=sql, con=self.engine1)
                columns = list(df.columns)
                columns = ', '.join(columns)
                df.to_sql('d1_host_cp', con=self.engine1, index=False, if_exists='replace')
                print('正在导入表总表中......')
                sql = '''REPLACE INTO {}_order_list({}, 记录时间) SELECT *, CURDATE() 记录时间 FROM d1_host_cp; '''.format(team,columns)
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            except Exception as e:
                print('插入失败：', str(Exception) + str(e))
            print('导入成功…………')
        elif query == '更新':
            try:
                print('正在更新临时表中......')
                df = pd.read_sql_query(sql=sql, con=self.engine1)
                df.to_sql('d1_host_cp', con=self.engine1, index=False, if_exists='replace')
                print('正在更新总表中......')
                sql = '''update {0}_order_list a, d1_host_cp b
                            set a.`币种`= b.`币种`,
                                a.`数量`= b.`数量`,
                                a.`电话号码`= b.`电话号码` ,
                                a.`运单编号`= IF(b.`运单编号` = '', NULL, b.`运单编号`),
                                a.`查件单号`= IF(b.`查件单号` = '', NULL, b.`查件单号`),
                                a.`系统订单状态`= IF(b.`系统订单状态` = '', NULL, b.`系统订单状态`),
                                a.`系统物流状态`= IF(b.`系统物流状态` = '' or b.`系统物流状态` = '发货中', NULL, b.`系统物流状态`),
                                a.`是否改派`= b.`是否改派`,
                                a.`物流方式`= IF(b.`物流方式` = '',NULL, b.`物流方式`),
                                a.`物流名称`= IF(b.`物流名称` = '', NULL, b.`物流名称`),
                                a.`货物类型`= IF(b.`货物类型` = '', NULL, b.`货物类型`),
                                a.`商品id`= IF(b.`商品id` = '', a.`商品id`, b.`商品id`),
                                a.`产品id`= IF(b.`产品id` = '', a.`产品id`, b.`产品id`),
                                a.`产品名称`= IF(b.`产品名称` = '', a.`产品名称`, b.`产品名称`),
                                a.`价格`= IF(b.`价格` = '', a.`价格`, b.`价格`),
                                a.`审核时间`= IF(b.`审核时间` = '', NULL, b.`审核时间`),
                                a.`上线时间`= IF(b.`上线时间` = '' or b.`上线时间` = '0000-00-00 00:00:00' , NULL, b.`上线时间`),
                                a.`仓储扫描时间`= IF(b.`仓储扫描时间` = '', NULL, b.`仓储扫描时间`),
                                a.`完结状态时间`= IF(b.`完结状态时间` = '', NULL, b.`完结状态时间`),
                                a.`包裹重量`= IF(b.`包裹重量` = '', NULL, b.`包裹重量`),
                                a.`省洲`= IF(b.`省洲` = '', NULL, b.`省洲`),
                                a.`规格中文`= IF(b.`规格中文` = '', NULL, b.`规格中文`),
                                a.`审单类型`= IF(b.`审单类型` = '', NULL, IF(b.`审单类型` like '%自动审单%','是','否')),
                                a.`审单类型明细`= IF(b.`审单类型` = '', NULL, b.`审单类型`),
                                a.`拉黑率`= IF(b.`拉黑率` = '', '0.00%', b.`拉黑率`),
                                a.`订单配送总量`= IF(b.`订单配送总量` = '', NULL, b.`订单配送总量`),
                                a.`签收量`= IF(b.`签收量` = '', NULL, b.`签收量`),
                                a.`拒收量`= IF(b.`拒收量` = '', NULL, b.`拒收量`),
                                a.`删除原因`= IF(b.`删除原因` = '', NULL,  b.`删除原因`),
                                a.`删除时间`= IF(b.`删除时间` = '', NULL,  b.`删除时间`),
                                a.`问题原因`= IF(b.`问题原因` = '', NULL,  b.`问题原因`),
                                a.`问题时间`= IF(b.`问题时间` = '', NULL,  b.`问题时间`),
                                a.`下单人`= IF(b.`下单人` = '', NULL,  b.`下单人`),
                                a.`克隆人`= IF(b.`克隆人` = '', NULL,  b.`克隆人`),
                                a.`选品人`= IF(b.`选品人` = '', NULL,  b.`选品人`)
                    where a.`订单编号`=b.`订单编号`;'''.format(team)
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            except Exception as e:
                print('更新失败：', str(Exception) + str(e))
            print('更新成功…………')



    # 查询更新（新后台的获取）
    def dayQuery(self, team):  # 进入订单检索界面，
        print('>>>>>>正式查询中<<<<<<')
        print('正在获取需要订单信息......')
        start = datetime.datetime.now()
        sql = '''SELECT `订单编号`  FROM {0};'''.format(team)
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(ordersDict)
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
            try:
                self.orderInfoQuery(ord)
            except Exception as e:
                print('获取失败： 30秒后重新获取', str(Exception) + str(e))
                time.sleep(30)
                self.orderInfoQuery(ord)
        print('单日查询耗时：', datetime.datetime.now() - start)


    # 更新团队订单明细（新后台的获取  方法一（1）的全部更新）
    def orderInfo(self, team, updata, begin, end):  # 进入订单检索界面
        # print('正在获取需要订单信息......')
        match1 = {'gat': '港台',
                  'slsc': '品牌'}
        if updata != '全部':
            # 获取日期时间
            sql = 'SELECT MAX(`日期`) 日期 FROM {0}_order_list;'.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['日期'][0])
            yy = int((rq - datetime.timedelta(days=9)).strftime('%Y'))
            mm = int((rq - datetime.timedelta(days=9)).strftime('%m'))
            dd = int((rq - datetime.timedelta(days=9)).strftime('%d'))
            begin = datetime.date(yy, mm, dd)
            print(begin)
            yy2 = int(datetime.datetime.now().strftime('%Y'))
            mm2 = int(datetime.datetime.now().strftime('%m'))
            dd2 = int(datetime.datetime.now().strftime('%d'))
            end = datetime.date(yy2, mm2, dd2)
            print(end)
        for i in range((end - begin).days):             # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            last_month = str(day)
            print('正在更新 ' + match1[team] + last_month + ' 号订单信息…………')
            start = datetime.datetime.now()
            sql = '''SELECT id,`订单编号`  FROM {0} sl WHERE sl.`日期` = '{1}';'''.format('gat_order_list', last_month)
            ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
            if ordersDict.empty:
                print('无需要更新订单信息！！！')
                return
            print(ordersDict['订单编号'][0])
            orderId = list(ordersDict['订单编号'])
            max_count = len(orderId)    # 使用len()获取列表的长度，上节学的
            n = 0
            while n < max_count:        # 这里用到了一个while循环，穿越过来的
                ord = ', '.join(orderId[n:n + 500])
                n = n + 500
                self.orderInfoQuery(ord)
            print('单日查询耗时：', datetime.datetime.now() - start)
    def orderInfoQuery(self, ord):  # 进入订单检索界面
        print('+++正在查询订单信息中')
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
        ordersDict = []
        count = 0
        try:
            for result in req['data']['list']:
                # print(result['orderNumber'])
                # print(result['specs'])
                # 添加新的字典键-值对，为下面的重新赋值用
                # 添加新的字典键-值对，为下面的重新赋值用
                if result['specs'] != []:
                    result['saleId'] = 0
                    result['saleProduct'] = 0
                    result['productId'] = 0
                    result['spec'] = 0
                    result['chooser'] = 0
                    result['saleId'] = result['specs'][0]['saleId']
                    result['saleProduct'] = (result['specs'][0]['saleProduct']).split('#')[2]
                    result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
                    result['spec'] = result['specs'][0]['spec']
                    result['chooser'] = result['specs'][0]['chooser']
                else:
                    result['saleId'] = ''
                    result['saleProduct'] = ''
                    result['productId'] = ''
                    result['spec'] = ''
                    result['chooser'] = ''
                result['auto_VerifyTip'] = ''
                result['order_count'] = ''
                result['order_qs'] = ''
                result['order_js'] = ''
                if result['autoVerifyTip'] == "":
                    result['auto_VerifyTip'] = '0.00%'
                else:
                    if '未读到拉黑表记录' in result['autoVerifyTip']:
                        result['auto_VerifyTip'] = '0.00%'
                    else:
                        if '拉黑率问题' not in result['autoVerifyTip']:
                            t2 = result['autoVerifyTip'].split('拉黑率')[1]
                            for tt2 in t2:
                                if '%' in tt2:
                                    result['auto_VerifyTip'] = t2.split('%')[0] + '%'
                            # t2 = result['autoVerifyTip'].split(',拉黑率')[1]
                            # result['auto_VerifyTip'] = t2.split('%;')[0] + '%'  ：26,：0,
                        elif '拉黑率问题' in result['autoVerifyTip']:
                            t2 = result['autoVerifyTip'].split('拒收订单量：')[1]
                            t2 = t2.split('%')[0]
                            result['auto_VerifyTip'] = t2.split('拉黑率')[1] + '%'
                    if '订单配送总量：' in result['autoVerifyTip']:
                        t4 = result['autoVerifyTip'].split('订单配送总量：')[1]
                        result['order_count'] = t4.split(',')[0]
                    if '送达订单量：' in result['autoVerifyTip']:
                        t4 = result['autoVerifyTip'].split('送达订单量：')[1]
                        result['order_qs'] = t4.split(',')[0]
                    if '拒收订单量：' in result['autoVerifyTip']:
                        t4 = result['autoVerifyTip'].split('拒收订单量：')[1]
                        result['order_js'] = t4.split(',')[0]
                quest = ''
                for re in result['questionReason']:
                    quest = quest + ';' + re
                result['questionReason'] = quest
                delr = ''
                for re in result['delReason']:
                    delr = delr + ';' + re
                result['delReason'] = delr
                auto = ''
                for re in result['autoVerify']:
                    auto = auto + ';' + re
                result['autoVerify'] = auto
                ordersDict.append(result)
            data = pd.json_normalize(ordersDict)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            count = count + 1
            time.sleep(10)
            python = sys.executable
            os.execl(python, python, 't2.py', * sys.argv)
            if count == 3:
                print('--->>>重启失败： 需手动重新启动！！！')
                pass
        # print('正在写入缓存中......')
        try:
            df = data[['orderNumber', 'currency', 'area', 'shipInfo.shipPhone', 'shipInfo.shipState', 'shipInfo.shipCity', 'wayBillNumber', 'saleId', 'saleProduct', 'productId',
                       'spec', 'quantity', 'orderStatus', 'logisticsStatus', 'logisticsName', 'addTime', 'verifyTime', 'transferTime', 'onlineTime', 'deliveryTime',
                       'finishTime', 'stateTime', 'logisticsUpdateTime', 'cloneUser', 'logisticsUpdateTime', 'reassignmentTypeName', 'dpeStyle', 'amount', 'payType',
                       'weight', 'autoVerify', 'delReason', 'delTime', 'questionReason', 'questionTime', 'service', 'chooser', 'logisticsRemarks', 'auto_VerifyTip',
                       'order_count', 'order_qs', 'order_js', 'composite_amount']]
            print(df)
            # print('正在更新临时表中......')
            df.to_sql('d1_cpy', con=self.engine1, index=False, if_exists='replace')
            sql = '''SELECT DATE(h.addTime) 日期,
            				    IF(h.`currency` = '日币', '日本', IF(h.`currency` = '泰铢', '泰国', IF(h.`currency` = '港币', '香港', IF(h.`currency` = '台币', '台湾', IF(h.`currency` = '韩元', '韩国', h.`currency`))))) 币种,
            				    h.orderNumber 订单编号,
            				    h.quantity 数量,
            				    h.`shipInfo.shipPhone` 电话号码,
            				    UPPER(h.wayBillNumber) 运单编号,
            				    IF(h.logisticsName LIKE "台湾-天马-711" AND LENGTH(h.wayBillNumber)=20, CONCAT(861,RIGHT(h.wayBillNumber,8)), IF((h.logisticsName LIKE "台湾-速派-新竹改派" or h.logisticsName LIKE "台湾-易速配-新竹改派") AND (h.wayBillNumber LIKE "A%" OR h.wayBillNumber LIKE "B%"),RIGHT(h.wayBillNumber,LENGTH(h.wayBillNumber)-1),UPPER(h.wayBillNumber))) 查件单号,
            				    h.orderStatus 系统订单状态,
            				    IF(h.`logisticsStatus` in ('发货中'), null, h.`logisticsStatus`) 系统物流状态,
            				    IF(h.`reassignmentTypeName` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
            				    TRIM(h.logisticsName) 物流方式,
            				    dim_trans_way.simple_name 物流名称,
            				    IF(h.`dpeStyle` = 'P 普通货', 'P', IF(h.`dpeStyle` = 'T 特殊货', 'T', h.`dpeStyle`)) 货物类型,
            				    h.`saleId` 商品id,
            				    h.`productId` 产品id,
            		            h.`saleProduct` 产品名称,
            		            h.amount 价格,
            				    h.verifyTime 审核时间,
            				    h.transferTime 转采购时间,
            				    h.onlineTime 上线时间,
            				    h.deliveryTime 仓储扫描时间,
            				    h.finishTime 完结状态时间,
            				    h.logisticsUpdateTime 物流更新时间,
            				    h.stateTime 状态时间,
            				    h.`weight` 包裹重量,
                                h.`shipInfo.shipState` 省洲,
                                h.`shipInfo.shipCity` 市区,
            				    h.`spec` 规格中文,
            				    h.`autoVerify` 审单类型,
            				    h.`auto_VerifyTip` 拉黑率,
            				    h.`order_count` 订单配送总量,
            				    h.`order_qs` 签收量,
            				    h.`order_js` 拒收量,
            				    IF(h.`delReason` LIKE ';%',RIGHT(h.`delReason`,LENGTH(h.`delReason`)-1),h.`delReason`) as 删除原因,
            				    h.`delTime` 删除时间,
            				    h.`questionReason` 问题原因,
            				    h.`questionTime` 问题时间,
            				    h.`service` 下单人,
            				    h.`cloneUser` 克隆人,
            				    h.`chooser` 选品人,
            				    h.`composite_amount` 组合销售金额,
            				    h.`logisticsRemarks` 物流状态注释
                            FROM d1_cpy h
                                LEFT JOIN dim_product ON  dim_product.sale_id = h.saleId
                                LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                                LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = TRIM(h.logisticsName);'''.format('gat_order_list')
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            df.to_sql('d1_cpy_cp', con=self.engine1, index=False, if_exists='replace')
            print('正在更新表总表中......')
            sql = '''update {0} a, d1_cpy_cp b
                            set a.`币种`= b.`币种`,
                                a.`数量`= b.`数量`,
                                a.`电话号码`= b.`电话号码` ,
                                a.`运单编号`= IF(b.`运单编号` = '', NULL, b.`运单编号`),
                                a.`查件单号`= IF(b.`查件单号` = '', NULL, b.`查件单号`),
                                a.`系统订单状态`= IF(b.`系统订单状态` = '', NULL, b.`系统订单状态`),
                                a.`系统物流状态`= IF(b.`系统物流状态` = '', NULL, b.`系统物流状态`),
                                a.`是否改派`= b.`是否改派`,
                                a.`物流方式`= IF(b.`物流方式` = '',NULL, b.`物流方式`),
                                a.`物流名称`= IF(b.`物流名称` = '', NULL, b.`物流名称`),
                                a.`货物类型`= IF(b.`货物类型` = '', NULL, b.`货物类型`),
                                a.`商品id`= IF(b.`商品id` = '', a.`商品id`, b.`商品id`),
                                a.`产品id`= IF(b.`产品id` = '', a.`产品id`, b.`产品id`),
                                a.`产品名称`= IF(b.`产品名称` = '', a.`产品名称`, b.`产品名称`),
                                a.`价格`= IF(b.`价格` = '', a.`价格`, b.`价格`),
                                a.`审核时间`= IF(b.`审核时间` = '', NULL, b.`审核时间`),
                                a.`上线时间`= IF(b.`上线时间` = '' or b.`上线时间` = '0000-00-00 00:00:00' , NULL, b.`上线时间`),
                                a.`仓储扫描时间`= IF(b.`仓储扫描时间` = '', NULL, b.`仓储扫描时间`),
                                a.`完结状态时间`= IF(b.`状态时间` = '', IF(b.`物流更新时间` = '', IF(b.`完结状态时间` = '', NULL, b.`完结状态时间`), b.`物流更新时间`), b.`状态时间`),
                                a.`包裹重量`= IF(b.`包裹重量` = '', NULL, b.`包裹重量`),
                                a.`省洲`= IF(b.`省洲` = '', NULL, b.`省洲`),
                                a.`市区`= IF(b.`市区` = '', NULL, b.`市区`),
                                a.`规格中文`= IF(b.`规格中文` = '', NULL, b.`规格中文`),
                                a.`审单类型`= IF(b.`审单类型` = '', NULL, IF(b.`审单类型` like '%自动审单%','是','否')),
                                a.`审单类型明细`= IF(b.`审单类型` = '', NULL, b.`审单类型`),
                                a.`拉黑率`= IF(b.`拉黑率` = '', '0.00%', b.`拉黑率`),
                                a.`订单配送总量`= IF(b.`订单配送总量` = '', NULL, b.`订单配送总量`),
                                a.`签收量`= IF(b.`签收量` = '', NULL, b.`签收量`),
                                a.`拒收量`= IF(b.`拒收量` = '', NULL, b.`拒收量`),
                                a.`删除原因`= IF(b.`删除原因` = '', NULL,  b.`删除原因`),
                                a.`删除时间`= IF(b.`删除时间` = '', NULL,  b.`删除时间`),
                                a.`问题原因`= IF(b.`问题原因` = '', NULL,  b.`问题原因`),
                                a.`问题时间`= IF(b.`问题时间` = '', NULL,  b.`问题时间`),
                                a.`下单人`= IF(b.`下单人` = '', NULL,  b.`下单人`),
                                a.`克隆人`= IF(b.`克隆人` = '', NULL,  b.`克隆人`),
                                a.`选品人`= IF(b.`选品人` = '', NULL,  b.`选品人`),
                                a.`组合销售金额`= IF(b.`组合销售金额` = '', NULL,  b.`组合销售金额`)
                    where a.`订单编号`=b.`订单编号`;'''.format('gat_order_list')
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        except Exception as e:
            print('更新失败：', str(Exception) + str(e))
        print('*************************本批次更新成功***********************************')
    # 更新团队订单明细（新后台的获取  方法一（2）的全部更新）
    def order_getList(self, team, updata, begin, end, proxy_handle, proxy_id):  # 进入订单检索界面
        # print('正在获取需要订单信息......')
        match1 = {'gat': '港台', 'slsc': '品牌'}
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500, 'orderPrefix': None, 'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                'emailStatus': None, 'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': None, 'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None,
                'orderStatus': None, 'tuan': None, 'tuanStatus': None, 'hasChangeSale': None, 'isComposite': None, 'optimizer': None, 'volumeEnd': None, 'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None, 'weightStart': None, 'weightEnd': None, 'estimateWeightStart': None,
                'estimateWeightEnd': None, 'order': None, 'sortField': None, 'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None, 'isChangeMark': None, 'percentStart': None,
                'percentEnd': None, 'userid': None, 'questionId': None, 'delUserId': None, 'transferNumber': None, 'smsStatus': None, 'designer_id': None, 'logistics_remarks': None, 'clone_type': None,
                'categoryId': None, 'addressType': None, 'timeStart': begin + ' 00:00:00', 'timeEnd': end + ' 23:59:59'}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)          # json类型数据转换为dict字典
        # print(req)
        max_count = req['data']['count']    # 获取 请求订单量
        # print(max_count)
        if max_count != 0 and max_count != []:
            print('++++++' + begin + ' 总计： ' + str(max_count) + ' 条信息+++++++')  # 获取总单量
            print('*' * 50)
            in_count = math.ceil(max_count / 500)
            df = pd.DataFrame([])
            dlist = []
            n = 1
            while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                data = self._order_getList(n, begin, end, proxy_handle, proxy_id)
                dlist.append(data)
                print('剩余查询次数' + str(in_count - n))
                n = n + 1
            dp = df.append(dlist, ignore_index=True)
            # print('正在写入缓存中......')
            dp = dp[['orderNumber', 'currency', 'area', 'shipInfo.shipPhone', 'shipInfo.shipState', 'shipInfo.shipCity','shipInfo.shipName', 'shipInfo.shipAddress','wayBillNumber','saleId', 'saleProduct', 'productId','spec','quantity', 'orderStatus',
                     'logisticsStatus', 'logisticsName', 'addTime', 'verifyTime','transferTime', 'onlineTime', 'deliveryTime','finishTime','stateTime', 'logisticsUpdateTime', 'cloneUser', 'logisticsUpdateTime', 'reassignmentTypeName',
                     'dpeStyle', 'amount', 'payType', 'weight', 'autoVerify', 'delReason', 'delTime', 'questionReason', 'questionTime', 'service', 'chooser', 'logisticsRemarks', 'auto_VerifyTip',
                     'percentInfo.arriveCount', 'percentInfo.orderCount', 'percentInfo.rejectCount', 'tel_phone', 'percent','warehouse','cloneTypeName', 'isBlindBox', 'mainOrderNumber', 'pre_second_numbers','abbreviation']]
            print(dp)
            # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
            # dp.to_excel('H:\\桌面\\需要用到的文件\\\输出文件\\派送问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            print('正在更新临时表中......')
            dp.to_sql('d1_cpy', con=self.engine1, index=False, if_exists='replace')
            sql = '''SELECT DATE(h.addTime) 日期,
                                    IF(h.`currency` = '日币', '日本', IF(h.`currency` = '泰铢', '泰国', IF(h.`currency` = '港币', '香港', IF(h.`currency` = '台币', '台湾', IF(h.`currency` = '韩元', '韩国', h.`currency`))))) 币种,
                                    h.orderNumber 订单编号,
                                    h.quantity 数量,
                                    h.`shipInfo.shipPhone` 电话号码,
                                    UPPER(h.wayBillNumber) 运单编号,
                                    IF(h.logisticsName LIKE "台湾-天马-711" AND LENGTH(h.wayBillNumber)=20, CONCAT(861,RIGHT(h.wayBillNumber,8)), IF((h.logisticsName LIKE "台湾-速派-新竹改派" or h.logisticsName LIKE "台湾-易速配-新竹改派") AND h.wayBillNumber LIKE "A%",RIGHT(h.wayBillNumber,LENGTH(h.wayBillNumber)-1),UPPER(h.wayBillNumber))) 查件单号,
                                    h.orderStatus 系统订单状态,
                                    IF(h.`logisticsStatus` in ('发货中'), null, h.`logisticsStatus`) 系统物流状态,
                                    IF(h.`reassignmentTypeName` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
                                    TRIM(h.payType) 付款方式,
                                    IF(TRIM(h.payType) NOT LIKE '%货到付款%','在线付款','货到付款') AS 支付类型,
                                    TRIM(h.logisticsName) 物流方式,
                                    dim_trans_way.simple_name 物流名称,
                                    IF(h.`dpeStyle` = 'P 普通货', 'P', IF(h.`dpeStyle` = 'T 特殊货', 'T', h.`dpeStyle`)) 货物类型,
                                    h.`saleId` 商品id,
                                    h.`productId` 产品id,
                                    h.`saleProduct` 产品名称,
                                    h.amount 价格,
                                    h.verifyTime 审核时间,
                                    h.transferTime 转采购时间,
                                    h.onlineTime 上线时间,
                                    h.deliveryTime 仓储扫描时间,
                                    h.finishTime 完结状态时间,
                                    h.logisticsUpdateTime 物流更新时间,
                                    h.stateTime 状态时间,
                                    h.`weight` 包裹重量,
                                    h.`shipInfo.shipState` 省洲,
                                    h.`shipInfo.shipCity` 市区,
                                    h.`spec` 规格中文,
                                    h.`autoVerify` 审单类型,
                                    h.`auto_VerifyTip` 拉黑率,
                                    h.`percentInfo.orderCount` 订单配送总量,
                                    h.`percentInfo.arriveCount` 签收量,
                                    h.`percentInfo.rejectCount` 拒收量,
                                    h.`delReason` 删除原因,
                                    h.`delTime` 删除时间,
                                    h.`questionReason` 问题原因,
                                    h.`questionTime` 问题时间,
                                    h.`service` 下单人,
                                    h.`cloneUser` 克隆人,
                                    h.`chooser` 选品人,
                                    h.`logisticsRemarks` 物流状态注释,
                                    dim_cate.`ppname` 父级分类,
                                    dim_cate.`pname` 二级分类,
                                    dim_cate.`name` 三级分类,
                                    h.`shipInfo.shipName` 姓名,
                                    h.`shipInfo.shipAddress` 地址,
                                    h.`tel_phone` 标准电话,
                                    h.`percent` 下单拒收率,
                                    h.`warehouse` 发货仓库,
                                    h.`cloneTypeName` 克隆类型,
                                    h.`isBlindBox` 是否盲盒,
                                    h.`mainOrderNumber` 主订单,
                                    h.`pre_second_numbers` 改派原运单号
                                   FROM d1_cpy h
                                       LEFT JOIN dim_product ON  dim_product.sale_id = h.saleId
                                       LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                                       LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = TRIM(h.logisticsName);'''.format('gat_order_list')
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            df.to_sql('d1_cpy_cp', con=self.engine1, index=False, if_exists='replace')
            print('正在更新表总表中......')
            sql = '''update {0} a, d1_cpy_cp b
                                   set a.`币种`= b.`币种`,
                                       a.`数量`= b.`数量`,
                                       a.`电话号码`= b.`电话号码` ,
                                       a.`运单编号`= IF(b.`运单编号` = '', NULL, b.`运单编号`),
                                       a.`查件单号`= IF(b.`查件单号` = '', NULL, b.`查件单号`),
                                       a.`系统订单状态`= IF(b.`系统订单状态` = '', NULL, b.`系统订单状态`),
                                       a.`系统物流状态`= IF(b.`系统物流状态` = '', NULL, b.`系统物流状态`),
                                       a.`是否改派`= b.`是否改派`,
                                       a.`付款方式`= IF(b.`付款方式` = '',NULL, b.`付款方式`),
                                       a.`支付类型`= IF(b.`支付类型` = '',NULL, b.`支付类型`),
                                       a.`物流方式`= IF(b.`物流方式` = '',NULL, b.`物流方式`),
                                       a.`物流渠道`= IF(b.`是否改派` ='直发',
                                                        IF(b.`物流方式` LIKE '香港-易速配-顺丰%','香港-易速配-顺丰', 
                                                            IF(b.`物流方式` LIKE '台湾-天马-711%' or b.`物流方式` LIKE '台湾-天马-新竹%','台湾-天马-新竹', 
                                                            IF(b.`物流方式` LIKE '台湾-铱熙无敌-新竹%' or b.`物流方式` LIKE '%优美宇通-新竹%','台湾-铱熙无敌-新竹', 
                                                            IF(b.`物流方式` LIKE '台湾-铱熙无敌-黑猫%','台湾-铱熙无敌-黑猫', 
                                                            IF(b.`物流方式` LIKE '台湾-铱熙无敌-711%','台湾-铱熙无敌-711超商', 
                                                            IF(b.`物流方式` LIKE '台湾-铱熙无敌-宅配通%','台湾-铱熙无敌-宅配通', 
                                                            IF(b.`物流方式` LIKE '台湾-速派-新竹%','台湾-速派-新竹', 
                                                            IF(b.`物流方式` LIKE '香港-立邦-改派','香港-立邦-顺丰', 
                                                            IF(b.`物流方式` LIKE '香港-圆通-改派','香港-圆通', b.`物流方式`)))))) ))),
                                                        IF(b.`物流方式` LIKE '香港-森鸿%','香港-森鸿-改派',
                                                            IF(b.`物流方式` LIKE '香港-立邦-顺丰%','香港-立邦-改派',
                                                            IF(b.`物流方式` LIKE '香港-易速配%','香港-易速配-改派',
                                                            IF(b.`物流方式` LIKE '台湾-立邦普货头程-森鸿尾程%' OR b.`物流方式` LIKE '台湾-大黄蜂普货头程-森鸿尾程%' OR b.`物流方式` LIKE '台湾-森鸿-新竹%','森鸿',
                                                            IF(b.`物流方式` LIKE '台湾-立邦普货头程-易速配尾程%' OR b.`物流方式` LIKE '台湾-大黄蜂普货头程-易速配尾程%','龟山',
                                                            IF(b.`物流方式` LIKE '台湾-易速配-龟山%' OR b.`物流方式` LIKE '台湾-易速配-新竹%' OR b.`物流方式` LIKE '新易速配-台湾-改派%' OR b.`物流方式` = '易速配','龟山',
                                                            IF(b.`物流方式` LIKE '台湾-天马-顺丰%','天马顺丰',
                                                            IF(b.`物流方式` LIKE '台湾-天马-新竹%' OR b.`物流方式` LIKE '台湾-天马-711%','天马新竹',
                                                            IF(b.`物流方式` LIKE '台湾-天马-黑猫%','天马黑猫',
                                                            IF(b.`物流方式` LIKE '台湾-速派-新竹%' OR b.`物流方式` LIKE '台湾-速派-711超商%','速派新竹',
                                                            IF(b.`物流方式` LIKE '台湾-速派宅配通%','速派宅配通',
                                                            IF(b.`物流方式` LIKE '台湾-速派-黑猫%','速派黑猫',
                                                            IF(b.`物流方式` LIKE '香港-圆通%','香港-圆通-改派',
                                                            IF(b.`物流方式` LIKE '台湾-优美宇通-新竹%','台湾-铱熙无敌-新竹改派',
                                                            IF(b.`物流方式` LIKE '台湾-铱熙无敌-黑猫普货' or b.`物流方式` LIKE '台湾-铱熙无敌-黑猫特货','台湾-铱熙无敌-黑猫改派',b.`物流方式`)))))))))))))))),
                                       a.`物流名称`= IF(b.`物流名称` = '', NULL, b.`物流名称`),
                                       a.`货物类型`= IF(b.`货物类型` = '', NULL, b.`货物类型`),
                                       a.`商品id`= IF(b.`商品id` = '', a.`商品id`, b.`商品id`),
                                       a.`产品id`= IF(b.`产品id` = '', a.`产品id`, b.`产品id`),
                                       a.`产品名称`= IF(b.`产品名称` = '', a.`产品名称`, b.`产品名称`),
                                       a.`价格`= IF(b.`价格` = '', a.`价格`, b.`价格`),
                                       a.`审核时间`= IF(b.`审核时间` = '', NULL, b.`审核时间`),
                                       a.`上线时间`= IF(b.`上线时间` = '' or b.`上线时间` = '0000-00-00 00:00:00' , NULL, b.`上线时间`),
                                       a.`仓储扫描时间`= IF(b.`仓储扫描时间` = '', NULL, b.`仓储扫描时间`),
                                       a.`完结状态时间`= IF(b.`状态时间` = '', IF(b.`物流更新时间` = '', IF(b.`完结状态时间` = '', NULL, b.`完结状态时间`), b.`物流更新时间`), b.`状态时间`),
                                       a.`包裹重量`= IF(b.`包裹重量` = '', NULL, b.`包裹重量`),
                                       a.`省洲`= IF(b.`省洲` = '', NULL, b.`省洲`),
                                       a.`市区`= IF(b.`市区` = '', NULL, b.`市区`),
                                       a.`规格中文`= IF(b.`规格中文` = '', NULL, b.`规格中文`),
                                       a.`审单类型`= IF(b.`审单类型` = '', NULL, IF(b.`审单类型` like '%自动审单%','是','否')),
                                       a.`审单类型明细`= IF(b.`审单类型` = '', NULL, b.`审单类型`),
                                       a.`拉黑率`= IF(b.`拉黑率` = '', '0.00%', b.`拉黑率`),
                                       a.`订单配送总量`= IF(b.`订单配送总量` = '', NULL, b.`订单配送总量`),
                                       a.`签收量`= IF(b.`签收量` = '', NULL, b.`签收量`),
                                       a.`拒收量`= IF(b.`拒收量` = '', NULL, b.`拒收量`),
                                       a.`删除原因`= IF(b.`删除原因` = '', NULL,  b.`删除原因`),
                                       a.`删除时间`= IF(b.`删除时间` = '', NULL,  b.`删除时间`),
                                       a.`问题原因`= IF(b.`问题原因` = '', NULL,  b.`问题原因`),
                                       a.`问题时间`= IF(b.`问题时间` = '', NULL,  b.`问题时间`),
                                       a.`下单人`= IF(b.`下单人` = '', NULL,  b.`下单人`),
                                       a.`克隆人`= IF(b.`克隆人` = '', NULL,  b.`克隆人`),
                                       a.`选品人`= IF(b.`选品人` = '', NULL,  b.`选品人`),
                                       a.`父级分类`= IF(a.`父级分类` IS NULL, IF(b.`父级分类` = '', NULL,  b.`父级分类`),  a.`父级分类`),
                                       a.`二级分类`= IF(a.`二级分类` IS NULL, IF(b.`二级分类` = '', NULL,  b.`二级分类`),  a.`二级分类`),
                                       a.`三级分类`= IF(a.`三级分类` IS NULL, IF(b.`三级分类` = '', NULL,  b.`三级分类`),  a.`三级分类`),
                                       a.`姓名`= IF(b.`姓名` = '', NULL,  b.`姓名`),
                                       a.`地址`= IF(b.`地址` = '', NULL,  b.`地址`),
                                       a.`标准电话`= IF(b.`标准电话` = '', NULL,  b.`标准电话`),
                                       a.`下单拒收率`= IF(b.`下单拒收率` = '', NULL,  b.`下单拒收率`),
                                       a.`发货仓库`= IF(b.`发货仓库` = '', NULL,  b.`发货仓库`),
                                       a.`克隆类型`= IF(b.`克隆类型` = '', NULL,  b.`克隆类型`),
                                       a.`是否盲盒`= IF(b.`是否盲盒` = '', NULL,  b.`是否盲盒`),
                                       a.`主订单`= IF(b.`主订单` = '', NULL,  b.`主订单`),
                                       a.`改派原运单号`= IF(b.`改派原运单号` = '', NULL,  b.`改派原运单号`)
                           where a.`订单编号`=b.`订单编号`;'''.format('gat_order_list')
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _order_getList(self, n, begin, end, proxy_handle, proxy_id):  # 进入订单检索界面
        # print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': n, 'pageSize': 500, 'orderPrefix': None, 'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                'emailStatus': None, 'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': None, 'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None,
                'orderStatus': None, 'tuan': None, 'tuanStatus': None, 'hasChangeSale': None, 'isComposite': None, 'optimizer': None, 'volumeEnd': None, 'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None, 'weightStart': None, 'weightEnd': None, 'estimateWeightStart': None,
                'estimateWeightEnd': None, 'order': None, 'sortField': None, 'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None, 'isChangeMark': None, 'percentStart': None,
                'percentEnd': None, 'userid': None, 'questionId': None, 'delUserId': None, 'transferNumber': None, 'smsStatus': None, 'designer_id': None, 'logistics_remarks': None, 'clone_type': None,
                'categoryId': None, 'addressType': None, 'timeStart': begin + ' 00:00:00', 'timeEnd': end + ' 23:59:59'}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        ordersDict = []
        count = 0
        try:
            for result in req['data']['list']:
                # print(result['orderNumber'])
                # if result['orderNumber'] == 'NR209160833054866' or result['orderNumber'] == 'GT209252247507204':
                    # print(result)
                # 添加新的字典键-值对，为下面的重新赋值用
                # 添加新的字典键-值对，为下面的重新赋值用
                if result['specs'] != []:
                    result['saleId'] = 0
                    result['saleProduct'] = 0
                    result['productId'] = 0
                    result['spec'] = 0
                    result['chooser'] = 0
                    result['saleId'] = result['specs'][0]['saleId']
                    result['saleProduct'] = (result['specs'][0]['saleProduct']).split('#')[2]
                    result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
                    result['spec'] = result['specs'][0]['spec']
                    result['chooser'] = result['specs'][0]['chooser']
                else:
                    result['saleId'] = ''
                    result['saleProduct'] = ''
                    result['productId'] = ''
                    result['spec'] = ''
                    result['chooser'] = ''
                result['auto_VerifyTip'] = ''
                if result['autoVerifyTip'] == "":
                    result['auto_VerifyTip'] = '0.00%'
                else:
                    if '未读到拉黑表记录' in result['autoVerifyTip']:
                        result['auto_VerifyTip'] = '0.00%'
                    else:
                        if '拉黑率问题' not in result['autoVerifyTip']:
                            if '拉黑率' not in result['autoVerifyTip']:
                                result['auto_VerifyTip'] = '0.00%'
                            else:
                                t2 = result['autoVerifyTip'].split('拉黑率')[1]
                                for tt2 in t2:
                                    if '%' in tt2:
                                        result['auto_VerifyTip'] = t2.split('%')[0] + '%'
                                # t2 = result['autoVerifyTip'].split(',拉黑率')[1]
                                # result['auto_VerifyTip'] = t2.split('%;')[0] + '%'  ：26,：0,
                        elif '拉黑率问题' in result['autoVerifyTip']:
                            t2 = result['autoVerifyTip'].split('拒收订单量：')[1]
                            t2 = t2.split('%')[0]
                            result['auto_VerifyTip'] = t2.split('拉黑率')[1] + '%'
                quest = ''
                for re in result['questionReason']:
                    quest = quest + ';' + re
                result['questionReason'] = quest
                delr = ''
                for re in result['delReason']:
                    delr = delr + ';' + re
                result['delReason'] = delr
                auto = ''
                for re in result['autoVerify']:
                    auto = auto + ';' + re
                result['autoVerify'] = auto
                ordersDict.append(result)
            data = pd.json_normalize(ordersDict)
            # print(data)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e) + begin)
        # print('*************************查询成功***********************************')
        return data


    # 更新团队订单明细（新后台的获取  方法二的部分更新）
    def orderInfo_th(self, searchType, team, team2, last_month, now_month):  # 进入订单检索界面，
        # print('正在获取需要订单信息......')
        start = datetime.datetime.now()
        sql = '''SELECT id,`订单编号`  FROM {0} sl WHERE sl.`日期` >= '{1}' and  sl.`日期` <'{2}';'''.format(team, last_month, now_month)
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        if ordersDict.empty:
            print('无需要更新订单信息！！！')
            # sys.exit()
            return
        print(ordersDict)
        print(ordersDict['订单编号'][0])
        orderId = list(ordersDict['订单编号'])
        # print('获取耗时：', datetime.datetime.now() - start)
        max_count = len(orderId)    # 使用len()获取列表的长度，上节学的
        if max_count > 500:
            in_count = math.ceil(max_count / 500)
            ord = ', '.join(orderId[0:500])
            df = self.orderInfoQuery_th(ord, searchType)
            dlist = []
            n = 0
            t = 1
            while n < max_count - 500:  # 这里用到了一个while循环，穿越过来的
                print('剩余查询次数 : ' + str(in_count - t))
                n = n + 500
                t = t + 1
                ord = ','.join(orderId[n:n + 500])
                data = self.orderInfoQuery_th(ord, searchType)
                dlist.append(data)
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
        else:
            ord = ','.join(orderId[0:max_count])
            dp = self.orderInfoQuery_th(ord, searchType)
        print('正在更新临时表中......')
        dp.to_sql('d1_cpy', con=self.engine1, index=False, if_exists='replace')
        sql = '''SELECT DATE(h.addTime) 日期,
        				    IF(h.`currency` = '日币', '日本', IF(h.`currency` = '泰铢', '泰国', IF(h.`currency` = '港币', '香港', IF(h.`currency` = '台币', '台湾', IF(h.`currency` = '韩元', '韩国', h.`currency`))))) 币种,
        				    h.orderNumber 订单编号,
        				    h.quantity 数量,
        				    h.`shipInfo.shipPhone` 电话号码,
        				    h.wayBillNumber 运单编号,
        				    h.orderStatus 系统订单状态,
        				    IF(h.`logisticsStatus` in ('发货中'), null, h.`logisticsStatus`) 系统物流状态,
        				    IF(h.`reassignmentTypeName` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
        				    TRIM(h.logisticsName) 物流方式,
        				    dim_trans_way.simple_name 物流名称,
        				    IF(h.`dpeStyle` = 'P 普通货', 'P', IF(h.`dpeStyle` = 'T 特殊货', 'T', h.`dpeStyle`)) 货物类型,
        				    h.`saleId` 商品id,
        				    h.`productId` 产品id,
        		            h.`saleProduct` 产品名称,
        				    h.verifyTime 审核时间,
        				    h.transferTime 转采购时间,
        				    h.onlineTime 上线时间,
        				    h.deliveryTime 仓储扫描时间,
        				    h.finishTime 完结状态时间,
        				    h.`weight` 包裹重量,
                            h.`shipInfo.shipState` 省洲,
        				    h.`spec` 规格中文,
        				    h.`autoVerify` 审单类型,
        				    h.`delReason` 删除原因,
        				    h.`delTime` 删除时间,
        				    h.`questionReason` 问题原因,
        				    h.`questionTime` 问题时间,
        				    h.`service` 下单人,
        				    h.`cloneUser` 克隆人
                        FROM d1_cpy h
                            LEFT JOIN dim_product ON  dim_product.sale_id = h.saleId
                            LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                            LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = TRIM(h.logisticsName);'''.format(team)
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df.to_sql('d1_cpy_cp', con=self.engine1, index=False, if_exists='replace')
        print('正在更新表总表中......')
        sql = '''update {0} a, d1_cpy_cp b
                        set a.`币种`= b.`币种`,
                            a.`数量`= b.`数量`,
                            a.`电话号码`= b.`电话号码` ,
                            a.`运单编号`= IF(b.`运单编号` = '', NULL, b.`运单编号`),
                            a.`系统订单状态`= IF(b.`系统订单状态` = '', NULL, b.`系统订单状态`),
                            a.`系统物流状态`= IF(b.`系统物流状态` = '', NULL, b.`系统物流状态`),
                            a.`是否改派`= b.`是否改派`,
                            a.`物流方式`= IF(b.`物流方式` = '',NULL, b.`物流方式`),
                            a.`物流名称`= IF(b.`物流名称` = '', NULL, b.`物流名称`),
                            a.`货物类型`= IF(b.`货物类型` = '', NULL, b.`货物类型`),
                            a.`商品id`= IF(b.`商品id` = '', a.`商品id`, b.`商品id`),
                            a.`产品id`= IF(b.`产品id` = '', a.`产品id`, b.`产品id`),
                            a.`产品名称`= IF(b.`产品名称` = '', a.`产品名称`, b.`产品名称`),
                            a.`审核时间`= IF(b.`审核时间` = '', NULL, b.`审核时间`),
                            a.`上线时间`= IF(b.`上线时间` = '' or b.`上线时间` = '0000-00-00 00:00:00' , a.`上线时间`, b.`上线时间`),
                            a.`仓储扫描时间`= IF(b.`仓储扫描时间` = '', NULL, b.`仓储扫描时间`),
                            a.`完结状态时间`= IF(b.`完结状态时间` = '', NULL, b.`完结状态时间`),
                            a.`包裹重量`= IF(b.`包裹重量` = '', NULL, b.`包裹重量`),
                            a.`省洲`= IF(b.`省洲` = '', NULL, b.`省洲`),
                            a.`规格中文`= IF(b.`规格中文` = '', NULL, b.`规格中文`),
                            a.`审单类型`= IF(b.`审单类型` = '', NULL, IF(b.`审单类型` like '%自动审单%','是','否')),
                            a.`删除原因`= IF(b.`删除原因` = '', NULL,  b.`删除原因`),
                            a.`删除时间`= IF(b.`删除时间` = '', NULL,  b.`删除时间`),
                            a.`问题原因`= IF(b.`问题原因` = '', NULL,  b.`问题原因`),
                            a.`问题时间`= IF(b.`问题时间` = '', NULL,  b.`问题时间`),
                            a.`下单人`= IF(b.`下单人` = '', NULL,  b.`下单人`),
                            a.`克隆人`= IF(b.`克隆人` = '', NULL,  b.`克隆人`)
                where a.`订单编号`=b.`订单编号`;'''.format(team2)
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=20000)
        print('单日查询耗时：', datetime.datetime.now() - start)
        print('*************************本批次查询成功***********************************')
    def orderInfoQuery_th(self, ord, searchType):  # 进入订单检索界面
        print('+++正在查询订单信息中')
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
        if searchType == '订单号':
            data.update({'orderPrefix': ord,
                         'shippingNumber': None})
        elif searchType == '运单号':
            data.update({'order_number': None,
                         'shippingNumber': ord})
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
        ordersDict = []
        count = 0
        try:
            for result in req['data']['list']:
                # print(result)
                # print(result['orderNumber'])
                # 添加新的字典键-值对，为下面的重新赋值用
                if result['specs'] != []:
                    result['saleId'] = 0
                    result['saleProduct'] = 0
                    result['productId'] = 0
                    result['spec'] = 0
                    result['saleId'] = result['specs'][0]['saleId']
                    result['saleProduct'] = (result['specs'][0]['saleProduct']).split('#')[2]
                    result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
                    result['spec'] = result['specs'][0]['spec']
                else:
                    result['saleId'] = ''
                    result['saleProduct'] = ''
                    result['productId'] = ''
                    result['spec'] = ''
                quest = ''
                for re in result['questionReason']:
                    quest = quest + ';' + re
                result['questionReason'] = quest
                delr = ''
                for re in result['delReason']:
                    delr = delr + ';' + re
                result['delReason'] = delr
                auto = ''
                for re in result['autoVerify']:
                    auto = auto + ';' + re
                result['autoVerify'] = auto
                ordersDict.append(result)
            data = pd.json_normalize(ordersDict)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            count = count + 1
            time.sleep(10)
            python = sys.executable
            os.execl(python, python, 't2.py', * sys.argv)
            if count == 3:
                print('--->>>重启失败： 需手动重新启动！！！')
                pass
        # print('正在写入缓存中......')
        # df = None
        try:
            df = data[['orderNumber', 'currency', 'area', 'shipInfo.shipPhone', 'shipInfo.shipState', 'wayBillNumber', 'saleId', 'saleProduct', 'productId',
                       'spec', 'quantity', 'orderStatus', 'logisticsStatus', 'logisticsName', 'addTime', 'verifyTime', 'transferTime', 'onlineTime', 'deliveryTime',
                       'finishTime', 'stateTime', 'logisticsUpdateTime', 'cloneUser', 'logisticsUpdateTime', 'reassignmentTypeName', 'dpeStyle', 'amount', 'payType',
                       'weight', 'autoVerify', 'delReason', 'delTime', 'questionReason', 'questionTime', 'service']]
            print(df)
        except Exception as e:
            print('查询失败：', str(Exception) + str(e))
        print('*************************单次查询成功***********************************')
        return df

    # 更新团队订单明细（新后台的获取  方法三的每天新增的订单更新）
    def orderInfo_append(self, timeStart, timeEnd, areaId, token,handle):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500, 'order_number': None, 'shippingNumber': None,
                'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None,
                'type': None, 'collId': None, 'isClone': None,
                'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': areaId, 'reassignmentType': None, 'lowerstatus': '',
                'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None, 'tuan': None,
                'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None, 'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None, 'weightStart': None,
                'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None, 'sortField': None,
                'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None,
                'timeStart': timeStart + '00:00:00', 'timeEnd': timeEnd + '23:59:59'}
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        req = json.loads(req.text)  # json类型数据转换为dict字典
        print('******首次查询成功******')
        max_count = req['data']['count']
        in_count = math.ceil(max_count/500)
        print('共 ' + str(max_count) + ' 条; 需查询' + str(in_count) + '次')
        df = pd.DataFrame([])               # 创建空的dataframe数据框
        dlist = []
        n = 1
        while n < in_count + 1:  # 这里用到了一个while循环，穿越过来的
            print('第' + str(n) + '次 查询')
            data = self._orderInfo_append(timeStart, timeEnd, n, areaId)
            n = n + 1
            dlist.append(data)
        dp = df.append(dlist, ignore_index=True)
        print('正在导入临时表中......')
        dp.to_sql('d1_host', con=self.engine1, index=False, if_exists='replace')
        sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
                    				        IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '11', '1', '2')) 旬,
                    			            DATE(h.下单时间) 日期,
                    				        h.运营团队 团队,
                    				        IF(h.`币种` = '台币', 'TW', IF(h.`币种` = '港币', 'HK', h.`币种`)) 区域,
                    				        IF(h.`币种` = '台币', '台湾', IF(h.`币种` = '港币', '香港', h.`币种`)) 币种,
                    				        h.平台 订单来源,
                    				        订单编号,
                    				        数量,
                    				        h.联系电话 电话号码,
                    				        h.运单号 运单编号,
                    				        null 查件单号,
                    				        h.订单状态 系统订单状态,
                    				        IF(h.`物流状态` = '发货中', null, h.`物流状态`) 系统物流状态,
                    				        IF(h.`订单类型` in ('未下架未改派','直发下架'), '直发', '改派') 是否改派,
                    				        h.物流渠道 物流方式,
                    				        dim_trans_way.simple_name 物流名称,
                    				        dim_trans_way.remark 运输方式,
                    				        IF(h.`货物类型` = 'P 普通货', 'P', IF(h.`货物类型` = 'T 特殊货', 'T', h.`货物类型`)) 货物类型,
                    				        IF(h.`是否低价` = 0, '否', '是') 是否低价,
                    				        商品ID,
                    				        产品id,
                    				        产品名称,
                    				        dim_cate.ppname 父级分类,
                    				        dim_cate.pname 二级分类,
                        		            dim_cate.`name` 三级分类,
                    				        h.支付方式 付款方式,
                    				        h.应付金额 价格,
                    				        IF(下单时间 = '',NULL,下单时间) 下单时间,
                    				        IF(审核时间 = '',NULL,审核时间) 审核时间,
                    				        IF(h.发货时间 = '',NULL,h.发货时间) 仓储扫描时间,
                    				        null 完结状态,
                    				        IF(h.完成时间 = '',NULL,h.完成时间) 完结状态时间,
                    				        null 价格RMB,
                    				        null 价格区间,
                    				        null 成本价,
                    				        null 物流花费,
                    				        null 打包花费,
                    				        null 其它花费,
                    				        h.重量 包裹重量,
                    				        h.体积 包裹体积,
                    				        邮编,
                    				        IF(h.转采购时间 = '',NULL,h.转采购时间) 添加物流单号时间,
                    				        null 规格中文,
                    				        h.省洲 省洲,
                    				        null 审单类型,
                    				        null 拉黑率,
                    				        null 删除原因,
                    				        null 删除时间,
                    				        null 问题原因,
                    				        null 问题时间,
                    				        null 下单人,
                    				        null 克隆人,
                    				        null 下架类型,
                    				        null 下架时间,
                    				        null 物流提货时间,
                    				        null 物流发货时间,
                    				        IF(h.上线时间 = '',NULL,h.上线时间) 上线时间,
                    				        null 国内清关时间,
                    				        null 目的清关时间,
                    				        null 回款时间,
                                            null IP,
                    				        null 选品人
                                    FROM d1_host h
                                    LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                                    LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                                    LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`
                                    WHERE h.下单时间 < TIMESTAMP(CURDATE()); '''.format('gat')
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df.to_sql('d1_host_cp', con=self.engine1, index=False, if_exists='replace')
        columns = list(df.columns)
        columns = ', '.join(columns)

        print('正在综合检查 父级分类、产品id 为空的信息---')
        sql = '''SELECT 日期,订单编号,商品id,产品id
                FROM {0} sl
                WHERE (sl.`父级分类` IS NULL or sl.`父级分类`= '' OR sl.`产品名称` IS NULL or sl.`产品名称`= '') AND ( NOT sl.`系统订单状态` IN ('已删除', '问题订单', '支付失败', '未支付'));'''.format('d1_host_cp')
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        if ordersDict.empty:
            print(' ****** 没有要补充的信息; ****** ')
        else:
            print('！！！ 请再次补充缺少的数据中！！！')
            lw = QueryTwoT('+86-18538110674', 'qyz04163510.', token, handle)
            lw.productInfo('d1_host_cp', ordersDict)

        print('正在导入 总表中......')
        sql = '''REPLACE INTO {0}_order_list({1}, 记录时间) SELECT *, CURDATE() 记录时间 FROM d1_host_cp; '''.format('gat', columns)
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        # print('正在写入Execl......')
        # print('正在写入Execl......')
        # dp.to_excel('F:\\输出文件\\订单检索-时间查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
        print('查询已导出+++')
    def _orderInfo_append(self, timeStart, timeEnd, n, areaId):  # 进入订单检索界面
        # print('......正在查询信息中......')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': n, 'pageSize': 500, 'order_number': None, 'shippingNumber': None,
                'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None,
                'type': None, 'collId': None, 'isClone': None,
                'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': areaId, 'reassignmentType': None, 'lowerstatus': '',
                'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None, 'tuan': None,
                'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None, 'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None, 'weightStart': None,
                'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None, 'sortField': None,
                'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None,
                'timeStart': timeStart + ' 00:00:00', 'timeEnd': timeEnd + ' 23:59:59'}
        # print(data)
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('......已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        ordersdict = []
        try:
            for result in req['data']['list']:
                if result['specs'] != '':
                    result['saleId'] = 0        # 添加新的字典键-值对，为下面的重新赋值用
                    result['saleName'] = 0
                    result['productId'] = 0
                    result['saleProduct'] = 0
                    result['spec'] = 0
                    result['chooser'] = 0
                    result['saleId'] = result['specs'][0]['saleId']
                    result['saleName'] = result['specs'][0]['saleName']
                    result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
                    result['saleProduct'] = (result['specs'][0]['saleProduct']).split('#')[2]
                    result['spec'] = result['specs'][0]['spec']
                    result['chooser'] = result['specs'][0]['chooser']
                else:
                    result['saleId'] = ''
                    result['saleProduct'] = ''
                    result['productId'] = ''
                    result['spec'] = ''
                quest = ''
                for re in result['questionReason']:
                    quest = quest + ';' + re
                result['questionReason'] = quest
                delr = ''
                for re in result['delReason']:
                    delr = delr + ';' + re
                result['delReason'] = delr
                auto = ''
                for re in result['autoVerify']:
                    auto = auto + ';' + re
                result['autoVerify'] = auto
                ordersdict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersdict)
        df = None
        try:
            df = data[['orderNumber', 'befrom', 'currency', 'area', 'productId', 'saleProduct', 'saleName', 'spec', 'shipInfo.shipName', 'shipInfo.shipPhone', 'percent', 'phoneLength',
                       'shipInfo.shipAddress', 'shipInfo.shipZip',  'amount', 'quantity', 'orderStatus', 'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime','logisticsName', 'dpeStyle',
                       'hasLowPrice', 'collId', 'saleId', 'reassignmentTypeName', 'logisticsStatus', 'weight', 'delReason', 'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime',
                        'finishTime', 'refundTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer','autoVerify', 'autoVerifyTip', 'cloneUser', 'isClone', 'warehouse',
                       'smsStatus', 'logisticsControl', 'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName', 'update_time']]
            df.columns = ['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格中文', '收货人', '联系电话', '拉黑率', '电话长度',
                          '配送地址', '邮编', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                          '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                          '完成时间', '销售退货时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '异常提示', '克隆人', '克隆ID', '发货仓库',
                          '是否发送短信', '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
        except Exception as e:
            print('------查询为空')
        print('......本批次查询成功......')
        print(df)
        return df


    # 改派-查询未发货的订单
    def gp_order(self, proxy_handle, proxy_id):
        print('正在获取 改派未发货…………')
        today = datetime.date.today().strftime('%Y.%m.%d')
        sql = '''SELECT xj.订单编号, xj.下单时间, xj.新运单号 运单编号, xj.查件单号, xj.产品id, xj.商品名称, xj.下架时间, xj.仓库, xj.物流渠道, xj.币种, xj.统计时间, xj.记录时间, b.物流状态, c.标准物流状态,b.状态时间, NULL 系统订单状态, NULL 系统物流状态, 
                        IF(统计时间 >=CURDATE() ,'未发货',NULL) AS 状态, g.工单类型, g.工单是否完成,g.提交形式, g.提交时间, g.登记人, g.运单编号 AS 运单号, g.最新处理人, g.最新处理时间, g.最新处理结果, g.同步操作记录, IF(LENGTH(xj.新运单号) >=20,'是',NULL) 导新运单号
                FROM ( SELECT *
                       FROM 已下架表 x
                       WHERE x.记录时间 >= TIMESTAMP ( CURDATE( ) ) AND x.币种 = '台币'
                ) xj
                LEFT JOIN gat_wl_data b ON xj.`查件单号` = b.`运单编号`
                LEFT JOIN gat_logisitis_match c ON b.物流状态 = c.签收表物流状态      
                LEFT JOIN 
                (   SELECT gd.*, GROUP_CONCAT(gd2.是否完成) as 工单是否完成
                    FROM 工单列表 gd
                    LEFT JOIN 工单列表 gd2 ON gd.`订单编号` = gd2.`订单编号`
                    WHERE DATE_FORMAT(gd.提交时间, '%Y-%m-%d') >= DATE_FORMAT(DATE_SUB(curdate(), INTERVAL 2 MONTH),'%Y-%m-%d')
                    GROUP BY 订单编号
                ) g ON xj.`订单编号` = g.`订单编号`;'''
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df = df.loc[df["币种"] == "台币"]
        df.to_sql('cache', con=self.engine1, index=False, if_exists='replace')

        print('正在查询 改派未发货…………')
        sql = '''SELECT 订单编号 FROM {0};'''.format('cache')
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        if ordersDict.empty:
            print('无需要更新订单信息！！！')
            # sys.exit()
            return
        orderId = list(ordersDict['订单编号'])
        max_count = len(orderId)    # 使用len()获取列表的长度，上节学的
        n = 0
        df = pd.DataFrame([])  # 创建空的dataframe数据框
        dlist = []
        while n < max_count:        # 这里用到了一个while循环，穿越过来的
            ord = ','.join(orderId[n:n + 500])
            n = n + 500
            data =self._gp_order(ord, proxy_handle, proxy_id)
            if data is not None and len(data) > 0:
                dlist.append(data)
        dp = df.append(dlist, ignore_index=True)
        dp.to_sql('cache_cp', con=self.engine1, index=False, if_exists='replace')

        print('正在更新 改派未发货......')
        sql = '''update `cache` a, `cache_cp` b
                        set a.`系统订单状态`= b.`orderStatus`,
                            a.`系统物流状态`= b.`logisticsStatus`,
                            a.`运单编号`= b.`wayBillNumber`
                where a.`订单编号`=b.`orderNumber`;'''.format('cache')
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)

        print('正在导出 改派未发货…………')
        sql = '''SELECT * FROM cache;'''.format('team')
        dt = pd.read_sql_query(sql=sql, con=self.engine1)
        file_path = 'F:\\神龙签收率\\(未发货) 改派-物流\\{} 改派未发货.xlsx'.format(today)
        dt.to_excel(file_path, sheet_name='台湾', index=False, engine='xlsxwriter')
        print('----已写入excel ')
    # 改派-查询未发货的订单（新后台的获取）
    def _gp_order(self, ord, proxy_handle, proxy_id):  # 进入订单检索界面
        print('+++正在查询 信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500, 'orderPrefix': ord, 'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'shippingNumber': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None,'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                'emailStatus': None, 'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': '','warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None,
                'orderStatus': None, 'tuan': None,  'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None, 'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None, 'weightStart': None, 'weightEnd': None, 'estimateWeightStart': None,
                'estimateWeightEnd': None, 'order': None, 'sortField': None, 'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None}
        if proxy_handle == '代理服务器':
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}
            req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        else:
            req = self.session.post(url=url, headers=r_header, data=data)
        # req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        ordersdict = []
        try:
            for result in req['data']['list']:
                # print(result)
                ordersdict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersdict)
        df = data[['orderNumber', 'orderStatus', 'wayBillNumber', 'logisticsName', 'logisticsStatus', 'warehouse', 'update_time']]
        # print('++++++本批次查询成功+++++++')
        print('*' * 50)
        # print(df)
        return df



if __name__ == '__main__':
    m = Query_sso_updata('+86-18538110674', 'qyz04163510.', 1343,'',"")
    start: datetime = datetime.datetime.now()
    match1 = {'gat': '港台',
              'gat_order_list': '港台',
              'slsc': '品牌'}
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    # for team in ['slsc', 'gat','slgat', 'slgat_hfh', 'slgat_hs', slrb', 'slrb_jl', 'slrb_js']:
    # 1、手动导入状态
    # for team in ['gat']:
    #     query = '导入'         # 导入；，更新--->>数据更新切换
    #     m.readFormHost(team, query)
    # 2、手动更新状态

    select = 3

    if int(select) == 1:
        for team in ['gat']:
            query = '导入'         # 导入；，更新--->>数据更新切换
            m.readFormHost(team, query)


    elif int(select) == 2:
        # -----------------------------------------------系统导入状态运行（二）-----------------------------------------
        #   台湾token, 日本token, 新马token：  f5dc2a3134c17a2e970977232e1aae9b
        #   泰国token： 83583b29fc24ec0529082ff7928246a6

        begin = datetime.date(2021, 10, 1)       # 1、手动设置时间；若无法查询，切换代理和直连的网络
        print(begin)
        end = datetime.date(2021, 10, 31)
        print(end)
        # yy = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y'))  # 2、自动设置时间
        # mm = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%m'))
        # begin = datetime.date(yy, mm, 1)
        # print(begin)
        # yy2 = int(datetime.datetime.now().strftime('%Y'))
        # mm2 = int(datetime.datetime.now().strftime('%m'))
        # dd2 = int(datetime.datetime.now().strftime('%d'))
        # end = datetime.date(yy2, mm2, dd2)
        # print(end)
        print(datetime.datetime.now())

        team = 'gat_order_list'     # 获取单号表
        team2 = 'gat_order_list'    # 更新单号表
        searchType = '订单号'  # 运单号，订单号   查询切换
        print('++++++正在获取 ' + match1[team] + ' 信息++++++')
        for i in range((end - begin).days):  # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            yesterday = str(day) + ' 23:59:59'
            last_month = str(day)
            now_month = str(day)
            print('正在更新 ' + match1[team] + str(last_month) + ' 号 --- ' + str(now_month) + ' 号信息…………')
            # m.orderInfo(searchType, team, team2, last_month)


        for i in range((end - begin).days):  # 按天循环获取订单状态
            print(i)
            last_month = begin + datetime.timedelta(days=5 * i)
            now_month = begin + datetime.timedelta(days=(i+1) * 5)
            if end >= now_month:
                print('正在更新 ' + str(last_month) + ' 号 --- ' + str(now_month) + ' 号信息…………')
                # m.orderInfo_th(searchType, team, team2, last_month, now_month)
            else:
                now_month = last_month + datetime.timedelta(days=(end - last_month).days)
                print('正在更新 ' + str(last_month) + ' 号 --- ' + str(now_month) + ' 号信息…………')
                # m.orderInfo_th(searchType, team, team2, last_month, now_month)
                break

        # m.orderInfoQuery('NR112151454534728', '订单号', 'gat_order_list', 'gat_order_list')  # 进入订单检索界面

        m.orderInfoQuery('GT203090849067593')  # 进入订单检索界面

        # print('更新耗时：', datetime.datetime.now() - start)

    elif int(select) == 3:
        areaId = 179                           # Line运营的areaId=179
        begin = datetime.date(2022, 4, 15)
        end = datetime.date(2022, 4, 15)
        m.orderInfo_append(str(begin), str(end), areaId)