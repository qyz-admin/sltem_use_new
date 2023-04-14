import os
import win32api,win32con
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from excelControl import ExcelControl
from mysqlControl import MysqlControl
from wlMysql import WlMysql
from wlExecl import WlExecl
from sso_updata import Query_sso_updata
from gat_update2临时 import QueryUpdate
import datetime
import time
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

import xlwings as xl
from settings import Settings
import pandas as pd
from sqlalchemy import create_engine
import math
import json
import requests
from settings_sso import Settings_sso

class Updata_gat(Settings):
    def __init__(self):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
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
        self.engine4 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql4['user'],
                                                                                    self.mysql4['password'],
                                                                                    self.mysql4['host'],
                                                                                    self.mysql4['port'],
                                                                                    self.mysql4['datebase']))

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

        time.sleep(1)
        if url != '/portal/index/index.html':
            # print('4.2、加载： ' + 'https://gimp.giikin.com')
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

        time.sleep(1)
        if url != '/portal/index/index.html':
            # print('4.2、加载： ' + 'https://gimp.giikin.com')
            # print('4.2、加载： ' + 'https://gimp.giikin.com')
            url = req.headers['Location']
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
            req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
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
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}  # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
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


    def creatMyOrderSl(self, begin, end) :  # 最近五天的全部订单信息
            match = {'gat': '"神龙家族-台湾", "红杉家族-港澳台2", "金狮-港澳台", "红杉家族-港澳台", "火凤凰-台湾", "火凤凰-香港", "金鹏家族-4组", "神龙-香港", "奥创队", "客服中心-港台","研发部-研发团队","Line运营", "神龙-主页运营", "翼虎家族-mercadolibre","金蝉家族公共团队","金蝉家族优化组","金蝉项目组","APP运营","郑州-北美"',
                     'slsc': '"金鹏家族-品牌", "金鹏家族-品牌1组", "金鹏家族-品牌2组", "金鹏家族-品牌3组"',
                     'sl_rb': '"神龙家族-日本团队", "金狮-日本", "红杉家族-日本", "红杉家族-日本666", "精灵家族-日本", "精灵家族-韩国", "精灵家族-品牌", "火凤凰-日本", "金牛家族-日本", "金鹏家族-小虎队", "奎蛇-日本", "奎蛇-韩国", "神龙-韩国"'
                     }
            match2 = {'gat': '17, 24, 26, 78, 118, 132, 135, 138, 156, 161, 173, 179, 182, 209, 225, 226, 234, 45, 41',
                     'slsc': '"金鹏家族-品牌", "金鹏家族-品牌1组", "金鹏家族-品牌2组", "金鹏家族-品牌3组"',
                     'sl_rb': '"神龙家族-日本团队", "金狮-日本", "红杉家族-日本", "红杉家族-日本666", "精灵家族-日本", "精灵家族-韩国", "精灵家族-品牌", "火凤凰-日本", "金牛家族-日本", "金鹏家族-小虎队", "奎蛇-日本", "奎蛇-韩国", "神龙-韩国"'
                     }
            for i in range((end - begin).days):  # 按天循环获取订单状态
                day = begin + datetime.timedelta(days=i)
                yesterday = str(day) + ' 23:59:59'
                last_month = str(day)
                print('正在获取 ' + match[team] + last_month[5:7] + '-' + yesterday[8:10] + ' 号订单…………')
                sql = '''SELECT a.id,
                            a.month 年月,
                            a.month_mid 旬,
                            a.rq 日期,
                --            IF(dim_area.name = "红杉家族-港澳台2","红杉家族-港澳台",IF(dim_area.name = "神龙家族-台湾" and dim_currency_lang.pname = '香港',"神龙-香港",IF(dim_area.name = "神龙-香港" and dim_currency_lang.pname = '台湾',"神龙家族-台湾",IF(dim_area.name = "火凤凰-台湾" and dim_currency_lang.pname = '香港',"火凤凰-香港" ,IF(dim_area.name = "火凤凰-香港" and dim_currency_lang.pname = '台湾',"火凤凰-台湾" ,dim_area.name))))) 团队,
                --            IF(dim_area.name in ('火凤凰-台湾','火凤凰-香港'),'火凤凰港台',IF(dim_area.name in ('神龙家族-台湾','神龙-香港'),'神龙港台',IF(dim_area.name = '客服中心-港台','客服中心港台',IF(dim_area.name = '研发部-研发团队','研发部港台',IF(dim_area.name = '神龙-主页运营','神龙主页运营',IF(dim_area.name = '红杉家族-港澳台','红杉港台',IF(dim_area.name = '郑州-北美','郑州北美',IF(dim_area.name = '金狮-港澳台','金狮港台',IF(dim_area.name = '金鹏家族-4组','金鹏港台',dim_area.name))))))))) 所属团队,
                            IF(a.area_id IN (24,78),'红杉家族-港澳台',IF(a.area_id = 17 AND a.currency_id = 6,"神龙-香港",IF(a.area_id = 138 AND a.currency_id = 13,"神龙家族-台湾",IF(a.area_id = 118 AND a.currency_id = 6,"火凤凰-香港",IF(a.area_id = 132 AND a.currency_id = 13,"火凤凰-台湾",dim_area.name)))))  团队,
                            IF(a.area_id in (118,132),'火凤凰港台',IF(a.area_id in (17,138),'神龙港台',IF(a.area_id = 161,'客服中心港台', IF(a.area_id = 173,'研发部港台',IF(a.area_id = 182,'神龙主页运营',IF(a.area_id in (24,78),'红杉港台',IF(a.area_id = 41,'郑州北美',IF(a.area_id = 26,'金狮港台',IF(a.area_id = 135,'金鹏港台',IF(a.area_id = 209,'翼虎港台',dim_area.name)))))))))) 所属团队,
                            a.region_code 区域,
                            dim_currency_lang.pname 币种,
                            a.beform 订单来源,
                            a.order_number 订单编号,
                            a.qty 数量,
                            a.ship_phone 电话号码,
                            UPPER(a.waybill_number) 运单编号,
                            IF(dim_trans_way.all_name LIKE "台湾-天马-711" AND LENGTH(a.waybill_number)=20, CONCAT(861,RIGHT(a.waybill_number,8)), UPPER(a.waybill_number)) 查件单号,
                --            a.order_status 系统订单状态id,
                --            IF(a.logistics_status = 1, 0, a.logistics_status) 系统物流状态id,
                            os.name 系统订单状态,
                            IF(ls.name ='发货中', null, ls.name) 系统物流状态,
                            IF(a.second=0,'直发','改派') 是否改派,
                            dim_trans_way.all_name 物流方式,
                            null 物流渠道,
                            dim_trans_way.simple_name 物流名称,
                            dim_trans_way.remark 运输方式,
                            a.logistics_type 货物类型,
                            IF(a.low_price=0,'否','是') 是否低价,
                            a.sale_id 商品id,
                            gk_sale.product_id 产品id,
                            gk_sale.product_name 产品名称,
                            dim_cate.ppname 父级分类,
                            dim_cate.pname 二级分类,
                            dim_cate.name 三级分类,
                            dim_payment.pay_name 付款方式,
                            IF(dim_payment.pay_name NOT LIKE '%货到付款%','在线付款','货到付款') AS 支付类型,
                            a.amount 价格,
                            a.addtime 下单时间,
                            a.verity_time 审核时间,
                            a.delivery_time 仓储扫描时间,
                            IF(a.finish_status=0,'未收款',IF(a.finish_status=2,'收款',IF(a.finish_status=3,'拒收',IF(a.finish_status=4,'退款',IF(a.finish_status=5,'售后订单',a.finish_status))))) 完结状态,
                            a.endtime 完结状态时间,   
                            a.salesRMB 价格RMB,
                            intervals.intervals 价格区间,
                            null 成本价,
                            a.logistics_cost 物流花费,
                            null 打包花费,
                            a.other_fee 其它花费,
                            a.weight 包裹重量,
                            null 包裹体积,
                            a.ship_zip 邮编,
                            a.turn_purchase_time 添加物流单号时间,
                            null 规格中文,
                            a.ship_state 省洲,
                            null 市区,
                            null 审单类型,
                            null 审单类型明细,
                            null 拉黑率,
                            null 订单配送总量,
                            null 拒收量,
                            null 签收量,
                            a.del_reason 删除原因,
                            null 删除时间,
                            a.question_reason 问题原因,
                            null 问题时间,
                            null 下单人,
                            null 克隆人,
                            a.stock_type 下架类型,
                            a.lower_time 下架时间,
                            a.tihuo_time 物流提货时间,
                            a.fahuo_time 物流发货时间,
                            a.online_time 上线时间,
                            a.guonei_time 国内清关时间,
                            a.mudidi_time 目的清关时间,
                            a.receipt_time 回款时间,
                            a.ip IP,
                            null 选品人,
                            null 组合销售金额,
                            null 姓名,
                            null 地址,
                            null 取货方式,
                            null 标准电话,
                            null 下单拒收率,
                            null 发货仓库,
                            null 克隆类型,
                            null 是否盲盒,
                            null 主订单,
                            null 改派原运单号
                    FROM gk_order a
                            LEFT JOIN dim_area ON dim_area.id = a.area_id
                            LEFT JOIN dim_payment ON dim_payment.id = a.payment_id
                            LEFT JOIN gk_sale ON gk_sale.id = a.sale_id
                            LEFT JOIN dim_trans_way ON dim_trans_way.id = a.logistics_id
                            LEFT JOIN dim_cate ON dim_cate.id = gk_sale.third_cate_id
                            LEFT JOIN intervals ON intervals.id = a.intervals
                            LEFT JOIN dim_currency_lang ON dim_currency_lang.id = a.currency_lang_id
                            LEFT JOIN dim_order_status os ON os.id = a.order_status
                            LEFT JOIN dim_logistics_status ls ON ls.id = a.logistics_status
                    WHERE  a.rq = '{0}' AND dim_area.id IN ({1});'''.format(last_month, match2[team])
                df = pd.read_sql_query(sql=sql, con=self.engine2)
                print('++++++正在将 ' + yesterday[8:10] + ' 号订单写入数据库++++++')
                # 这一句会报错,需要修改my.ini文件中的[mysqld]段中的"max_allowed_packet = 1024M"
                try:
                    df.to_sql('sl_order_t22', con=self.engine1, index=False, if_exists='replace')
                    sql = 'REPLACE INTO {0}_order_list SELECT *, NOW() 记录时间 FROM sl_order_t22; '.format('gat')
                    pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
                except Exception as e:
                    print('插入失败：', str(Exception) + str(e))
                print('-' * 20 + '写入完成' + '-' * 20)
            return '写入完成'

    # 更新团队订单明细（新后台的获取  方法一（2）的全部更新）
    def order_getList(self, handle, login_TmpCode, begin, end, proxy_handle, proxy_id):  # 进入订单检索界面
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
            print('正在更新临时表中......')
            dp.to_sql('d1_cpy_t22', con=self.engine1, index=False, if_exists='replace')
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
                                   FROM d1_cpy_t22 h
                                       LEFT JOIN dim_product ON  dim_product.sale_id = h.saleId
                                       LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                                       LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = TRIM(h.logisticsName);'''.format('gat_order_list')
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            df.to_sql('d1_cpy_cp_t22', con=self.engine1, index=False, if_exists='replace')

            print('正在更新表总表中......')
            sql = '''update {0} a, d1_cpy_cp_t22 b
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
                                                            IF(b.`物流方式` LIKE '台湾-速派-新竹%','台湾-速派-新竹', 
                                                            IF(b.`物流方式` LIKE '香港-立邦-改派','香港-立邦-顺丰', 
                                                            IF(b.`物流方式` LIKE '香港-圆通-改派','香港-圆通', b.`物流方式`)))))) )),
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
        # dp.to_excel('G:\\输出文件\\订单检索-时间查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
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


if __name__ == '__main__':
    u = Updata_gat()
    # 初始化时间设置
    proxy_handle = '代理服务器0'
    proxy_id = '192.168.13.89:37466'  # 输入代理服务器节点和端口
    handle = '手动0'
    login_TmpCode = '0bd57ce215513982b1a984d363469e30'  # 输入登录口令Tkoen
    team = 'gat'

    if team == 'gat0':
        # 更新时间
        timeStart = (datetime.datetime.now() - relativedelta(months=1)).strftime('%Y-%m') + '-01'
        data_begin = datetime.datetime.strptime(timeStart, '%Y-%m-%d').date()
        begin = data_begin
        end = datetime.datetime.now().date()
    else:
        # 更新时间
        data_begin = datetime.date(2023, 4, 10)  # 数据库更新
        begin = datetime.date(2023, 4, 10)  # 单点更新
        end = datetime.date(2023, 4, 14)

    print('****** 数据库更新起止时间：' + data_begin.strftime('%Y-%m-%d') + ' - ' + end.strftime('%Y-%m-%d') + ' ******')
    print('****** 单点  更新起止时间：' + begin.strftime('%Y-%m-%d') + ' - ' + end.strftime('%Y-%m-%d') + ' ******')

    # TODO------------------------------------更新数据  数据库分段读取------------------------------------
    # u.creatMyOrderSl(begin, end)

    # TODO------------------------------------更新数据  单点检索读取------------------------------------
    if proxy_handle == '代理服务器':
        if handle == '手动':
            u.sso__online_handle_proxy(login_TmpCode, proxy_id)
        else:
            u.sso__online_auto_gp_proxy(proxy_id)
    else:
        if handle == '手动':
            u.sso__online_handle(login_TmpCode)
        else:
            u.sso__online_auto_gp()
    for i in range((end - begin).days):                             # 按天循环获取订单状态
        day = begin + datetime.timedelta(days=i)
        day_time = str(day)
        u.order_getList(handle, login_TmpCode, day_time, day_time, proxy_handle, proxy_id)




    # # TODO------------------------------------新增数据  单点检索读取------------------------------------
    # print('---------------------------------- 单点导入更新部分：--------------------------------')
    # u.order_getList(begin, end, proxy_handle, proxy_id)
    #
    # print('---------------------------------- 手动导入更新部分：--------------------------------')
    # handle = '手动'
    # sso = Query_sso_updata('+86-18538110674', 'qyz04163510.', '1343','',handle, proxy_handle, proxy_id)
    # sso.readFormHost('gat', '导入')                                   # 导入新增的订单 line运营  手动导入
    # sso.readFormHost('gat', '更新')                                   # 更新新增的订单 手动导入


