import datetime
import time
import win32api,win32con
import sys
import json
from sqlalchemy import create_engine
from settings import Settings
import pandas as pd
from dateutil.relativedelta import relativedelta
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
import requests

class Settings_sso():
    def __init__(self):
        self.SS = Settings()
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.userMobile = '+86-17596568562'
        self.password = 'xhy123456'
        self.userID = '1343'
        # self.userMobile = '+86-15565053520'
        # self.password = 'sunan1022wang.@&'
        # self.userID = '168'
        self.engine1 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.SS.mysql1['user'],
                                                                                    self.SS.mysql1['password'],
                                                                                    self.SS.mysql1['host'],
                                                                                    self.SS.mysql1['port'],
                                                                                    self.SS.mysql1['datebase']))
        self.engine2 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.SS.mysql2['user'],
                                                                                    self.SS.mysql2['password'],
                                                                                    self.SS.mysql2['host'],
                                                                                    self.SS.mysql2['port'],
                                                                                    self.SS.mysql2['datebase']))
        self.engine20 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.SS.mysql20['user'],
                                                                                    self.SS.mysql20['password'],
                                                                                    self.SS.mysql20['host'],
                                                                                    self.SS.mysql20['port'],
                                                                                    self.SS.mysql20['datebase']))
        self.engine3 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.SS.mysql3['user'],
                                                                                    self.SS.mysql3['password'],
                                                                                    self.SS.mysql3['host'],
                                                                                    self.SS.mysql3['port'],
                                                                                    self.SS.mysql3['datebase']))
        # 单点系统登录使用
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

    def sso_online_Two_Five(self):  # 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171651910731826177971417081651918686746378G76942D6B6E83AC559B7B9F797D5850AF4E933E',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
            'Origin': 'https://login.dingtalk.com',
            'Accept': '*/*',
            'Content-Type': 'application/x-www-form-urlencoded',
            'Sec-Fetch-Mode': 'cors',
            'Referer': 'https://login.dingtalk.com/'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        print(req)
        print(req.headers)
        print(req.text)
        print(req.content)
        print(req.cookies)
        print(req.url)
        print(req.apparent_encoding)
        print(req.history)
        print(req.links)
        print(req.next)
        req = req.json()
        print(req)
        # req_url = req['data']
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        if 'data' in req.keys():
            try:
                req_url = req['data']
                loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self.sso_online_Two()
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
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('三、dingtalk_service服务器......')
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        gimp = req.headers['Location']

        time.sleep(1)
        # print('（3.1）加载： ' + str(gimp))
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index_system3 = req.headers['Location']
        # print(808080)
        index_system3 = index_system3.replace(':443', '')
        # print(index_system3)

        time.sleep(1)
        # print('（3.2）再次加载/dingtalk_service/getunionidbytempcode?code=页面......')
        url = index_system3
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # print('（3.3）加载//gimp.giikin.com 页面......')
        # print(990099900)
        time.sleep(1)
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('四、加载/gimp.giikin.com/portal/index/index.html 页面......')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        # print(8080)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system = req.headers['Location']
        # print(7070)

        time.sleep(1)
        # print('（4.1）加载index.html?_system=18页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('（4.2）加载/gimp.giikin.com/portal/index/index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(6060)
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

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

        # 仓储系统登录使用
    def sso_online_Two(self):  # 登录系统保持会话状态
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
        print(req)
        req = req.json()
        print(req)
        # req_url = req['data']
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        if 'data' in req.keys():
            try:
                req_url = req['data']
                loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self.sso_online_Two()
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
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('三、dingtalk_service服务器......')
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        gimp = req.headers['Location']

        time.sleep(1)
        # print('（3.1）加载： ' + str(gimp))
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index_system3 = req.headers['Location']
        # print(808080)
        index_system3 = index_system3.replace(':443', '')
        # print(index_system3)

        time.sleep(1)
        # print('（3.2）再次加载/dingtalk_service/getunionidbytempcode?code=页面......')
        url = index_system3
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # print('（3.3）加载//gimp.giikin.com 页面......')
        # print(990099900)
        time.sleep(1)
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('四、加载/gimp.giikin.com/portal/index/index.html 页面......')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        # print(8080)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system = req.headers['Location']
        # print(7070)

        time.sleep(1)
        # print('（4.1）加载index.html?_system=18页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('（4.2）加载/gimp.giikin.com/portal/index/index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(6060)
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

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

        # 仓储系统登录使用
    def sso_online_Two_T(self):  # 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171642209116632183727649041642209149705969GCB15B029EA5D5E340FD6CEF95DA55D48563DD7',
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
                self.sso_online_Two()
        elif 'message' in req.keys():
            info = req['message']
            win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()
        else:
            print('请检查失败原因：', str(req))
            win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
            sys.exit()
        print('******已获取loginTmpCode值: ' + str(loginTmpCode))

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
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('三、dingtalk_service服务器......')
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        gimp = req.headers['Location']

        time.sleep(1)
        print('（3.1）加载： ' + str(gimp))
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index_system3 = req.headers['Location']
        # print(808080)
        index_system3 = index_system3.replace(':443', '')
        # print(index_system3)

        time.sleep(1)
        # print('（3.2）再次加载/dingtalk_service/getunionidbytempcode?code=页面......')
        url = index_system3
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        # print('（3.3）加载//gimp.giikin.com 页面......')
        # print(990099900)
        time.sleep(1)
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('四、加载/gimp.giikin.com/portal/index/index.html 页面......')
        url = 'https://gimp.giikin.com' + index
        # print(url)
        # print(8080)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        index_system = req.headers['Location']
        # print(7070)

        time.sleep(1)
        # print('（4.1）加载index.html?_system=18页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('（4.2）加载/gimp.giikin.com/portal/index/index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(6060)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(5050)
        # print(req)
        # print(req.headers)

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

        # 仓储系统登录使用
    def sso_online_cang(self):  # 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171640662225970131824980691640846029429745GC1F1BF386B34F4C680DD7B7D2938FA61F3FF27',
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
        req_url = req['data']
        loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        if 'data' in req.keys():
            try:
                req_url = req['data']
                loginTmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self.sso_online_Two()
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
        url = r'http://gwms-v3.giikin.cn/tool/dingtalk_service/gettempcodebylogin'
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
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('三、dingtalk_service服务器......')
        # print('（3.0）加载： ' + str(req.text))
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        gimp = req.headers['Location']

        time.sleep(1)
        # print('（3.1）加载： ' + str(gimp))
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)

        # print('（3.2）加载： http://gwms-v3.giikin.cn/admin/index/index')
        url = 'http://gwms-v3.giikin.cn/admin/index/index'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': gimp}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

        # 查询订单更新 以订单编号 （单点系统）

    # 发送钉钉消息
    def send_dingtalk_message(self, url, content, mobile_list, isAtAll):
        headers = {'Content-Type': 'application/json', "Charset": "UTF-8"}
        if isAtAll == '是':
            data = {"msgtype": "text",
                    "text": {  # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                        "content": content
                    },
                    "at": {
                        # "atMobiles": mobile_list, @所有人写True并且将上面atMobiles注释掉
                        # 是否@所有人
                        "isAtAll": True
                    }
            }
        if isAtAll == '单个':
            data = {"msgtype": "text",
                    "text": {  # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                        "content": content
                    },
                    "at": {
                        # "atMobiles": mobile_list, @所有人写True并且将上面atMobiles注释掉
                        # 是否@所有人
                        "isAtAll": False
                    }
            }
        else:
            data = {
                "msgtype": "text",
                "text": {
                    # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                    "content": content
                },
                "at": {
                    # 要@的人
                    "atMobiles": mobile_list,
                    # 是否@所有人
                    "isAtAll": False  # @全体成员（在此可设置@特定某人）
                }
            }

        # 4、对请求的数据进行json封装
        sendData = json.dumps(data)  # 将字典类型数据转化为json格式
        sendData = sendData.encode("utf-8")  # python3的Request要求data为byte类型

        r = requests.post(url, headers=headers, data=json.dumps(data))
        req = json.loads(r.text)  # json类型数据转换为dict字典
        print(req['errmsg'])
        # return req['errmsg']

    # 自动输入token
    def sso_online_cang_auto(self):  # 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录仓储系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': self.userMobile,
                'pwd': self.password,
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171629428116275265671469741658739291139085GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:94.0) Gecko/20100101 Firefox/94.0',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0',
                    'Host': 'login.dingtalk.com',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # req = req.json()
        req = {}
        # print(req)
        # req_url = req['data']
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        login_TmpCode = '获取不到参数'
        if 'data' in req.keys():
            try:
                req_url = req['data']
                login_TmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
            except Exception as e:
                print('重新启动： 3分钟后', str(Exception) + str(e))
                time.sleep(300)
                self.sso_online_Two()
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
                                    data: { mobile: '+86-17596568562',
                                            pwd: 'xhy123456',
                                            goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode',
                                            pdmToken: '',
                                            araAppkey: '1917',
                                            araToken: '0#19171646622570440595157649661653036928565594G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
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
            options = webdriver.ChromeOptions()
            options.add_argument(r"user-data-dir=C:\Program Files\Google\Chrome\Application\profile")
            driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')
            driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode')
            time.sleep(3)
            js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
                        data: { mobile: '+86-17596568562',
                                pwd: 'xhy123456',
                                goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode',
                                pdmToken: '',
                                araAppkey: '1917',
                                araToken: '0#19171646622570440595157649661653036928565594G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
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
            time.sleep(3)
            login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
            print('loginTmpCode值: ' + login_TmpCode)
            driver.quit()

        print('******已获取loginTmpCode值: ' + str(login_TmpCode))
        loginTmpCode = login_TmpCode
        # print('二、请求-后台登录页面......')
        url = r'http://gwms-v3.giikin.cn/tool/dingtalk_service/gettempcodebylogin'
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
        # print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('三、dingtalk_service服务器......')
        # print('（3.0）加载： ' + str(req.text))
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        gimp = req.headers['Location']

        time.sleep(1)
        # print('（3.1）加载： ' + str(gimp))
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)

        # time.sleep(1)
        # print('（3.2）加载： http://gwms-v3.giikin.cn/admin/index/index')
        url = 'http://gwms-v3.giikin.cn/admin/index/index'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': gimp}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req.headers)
        print('++++++已成功登录++++++++++++ ' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

        # 查询订单更新 以订单编号 （单点系统）
    # 不使用代理服务器
    def sso__online_auto(self):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': '+86-17596568562',
                'pwd': 'xhy123456',
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
                self.sso_online_Two()
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
                        data: { mobile: '+86-17596568562',
                                pwd: 'xhy123456',
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
                        data: { mobile: '+86-17596568562',
                                pwd: 'xhy123456',
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
        index_system3 = req.headers['Location']
        # print(88)
        # print(index_system3)
        # 此处暂停使用443
        if ':443' in index_system3:
            # print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
            # print(index_system3)
            url = index_system3.replace(':443', '')
            # print(url)
            r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                        'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            # print(req)
            # print(req.headers)
            url = req.headers['Location']
        else:
            url = index_system3

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
        if '/admin/login_by_dingtalk/finishLoginJump?jump_url=' in url:
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
    def sso__online_auto_proxy(self, proxy_id):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        # print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': '+86-17596568562',
                'pwd': 'xhy123456',
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
                self.sso_online_Two()
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
                        data: { mobile: '+86-17596568562',
                                pwd: 'xhy123456',
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
                        data: { mobile: '+86-17596568562',
                                pwd: 'xhy123456',
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
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}            # 使用代理服务器
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
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}            # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}            # 使用代理服务器
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
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}            # 使用代理服务器
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
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}            # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        # print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}            # 使用代理服务器
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

    def sso__online_auto22(self):  # 手动输入token 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录后台系统中......')
        url = r'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:94.0) Gecko/20100101 Firefox/94.0',
                    'Host': 'oapi.dingtalk.com',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        # print(req.text)
        print(12121)

        url = r'https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:94.0) Gecko/20100101 Firefox/94.0',
                    'Host': 'login.dingtalk.com',
                    'Referer': 'https://oapi.dingtalk.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        # print(req.text)
        print(33333)

        print('一、获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': "+86-17596568562",
                'pwd': 'xhy123456',
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171629428116275265671469741658810338032256GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:94.0) Gecko/20100101 Firefox/94.0',
                    'Host': 'login.dingtalk.com',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode'}
        print(data)
        req = self.session.post(url=url, headers=r_header, data=data)
        print(req.headers)
        print(req.text)
        req = req.json()
        # req = {}
        print(req)
        print(888)
        req_url = req['data']  # 0#19171629428116275265671469741656903392035557GC87818BBCC3CCDF73DCA3659F13FFA069CD0EA
        login_TmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        # login_TmpCode = '获取不到参数'
        # if 'data' in req.keys():
        #     try:
        #         req_url = req['data']
        #         login_TmpCode = req_url.split('loginTmpCode=')[1]  # 获取loginTmpCode值
        #     except Exception as e:
        #         print('重新启动： 3分钟后', str(Exception) + str(e))
        #         time.sleep(300)
        #         self.sso_online_Two()
        # elif 'message' in req.keys():
        #     info = req['message']
        #     win32api.MessageBox(0, "登录失败: " + info, "错误 提醒", win32con.MB_ICONSTOP)
        #     # sys.exit()
        # else:
        #     print('请检查失败原因：', str(req))
        #     win32api.MessageBox(0, "请检查失败原因: 是否触发了验证码； 或者3分钟后再尝试登录！！！", "错误 提醒", win32con.MB_ICONSTOP)
        #     # sys.exit()
        # if login_TmpCode == '获取不到参数':
        #     time.sleep(1)
        #     # 模拟打开火狐浏览器 获取token
        #     options = Options()
        #     # options.add_argument('-headless')
        #     driver = webdriver.Firefox(options=options)
        #     # driver.get('https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
        #     driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
        #     driver.implicitly_wait(5)
        #     js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
        #                 data: { mobile: '+86-17596568562',
        #                         pwd: 'xhy123456',
        #                         goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
        #                         pdmToken: '',
        #                         araAppkey: '1917',
        #                         araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
        #                         araScene: 'login',
        #                         captchaImgCode: '',
        #                         captchaSessionId: '',
        #                         type: 'h5'
        #                     },
        #                     type: 'POST',
        #                     timeout: '10000',
        #                     async:false,
        #                     beforeSend(xhr, settings) {
        #                         xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
        #                     },
        #                     success: function(data) {
        #                         if (data.success) {
        #                              console.log(data.data);
        #                              console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
        #                               document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
        #                              arguments[0].value=data.data.split('loginTmpCode=')[1];
        #                         } else {
        #                                 console.log(data.code);
        #                         }
        #                     },
        #                     error: function(error) {
        #                         alert("请检查网络");
        #                     }
        #                 });
        #                 '''
        #     element = driver.find_element('id', 'mobile')
        #     driver.execute_script(js, element)
        #     # driver.implicitly_wait(5)
        #     time.sleep(5)
        #     login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
        #     print('loginTmpCode值: ' + login_TmpCode)
        #     driver.quit()

        # elif login_TmpCode == '获取不到参数':
        #     time.sleep(1)
        #     # 模拟打开谷歌浏览器 获取token
        #     options = webdriver.ChromeOptions()
        #     options.add_argument(r"user-data-dir=C:\Program Files\Google\Chrome\Application\profile")
        #     driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')

        #     driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
        #     # driver.implicitly_wait(5)
        #     time.sleep(5)
        #     js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
        #                 data: { mobile: '+86-17596568562',
        #                         pwd: 'xhy123456',
        #                         goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
        #                         pdmToken: '',
        #                         araAppkey: '1917',
        #                         araToken: '0#19171646622570440595157649661658739404065586G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
        #                         araScene: 'login',
        #                         captchaImgCode: '',
        #                         captchaSessionId: '',
        #                         type: 'h5'
        #                     },
        #                     type: 'POST',
        #                     timeout: '10000',
        #                     async:false,
        #                     beforeSend(xhr, settings) {
        #                         xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
        #                     },
        #                     success: function(data) {
        #                         if (data.success) {
        #                              console.log(data.data);
        #                              console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
        #                               document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
        #                              arguments[0].value=data.data.split('loginTmpCode=')[1];
        #                         } else {
        #                                 console.log(data.code);
        #                         }
        #                     },
        #                     error: function(error) {
        #                         alert("请检查网络");
        #                     }
        #                 });
        #                 '''
        #     element = driver.find_element('id', 'mobile')
        #     driver.execute_script(js, element)
        #     # driver.implicitly_wait(5)
        #     time.sleep(5)
        #     login_TmpCode = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
        #     print('loginTmpCode值: ' + login_TmpCode)
        #     driver.quit()



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

    # 手动输入token
    def sso_online_cang_handle(self, login_TmpCode):  # 登录系统保持会话状态
        print(datetime.datetime.now())
        print('正在登录仓储系统中......')
        # # print('一、获取-钉钉用户信息......')
        # url = r'https://login.dingtalk.com/login/login_with_pwd'
        # data = {'mobile': self.userMobile,
        #         'pwd': self.password,
        #         'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode',
        #         'pdmToken': '',
        #         'araAppkey': '1917',
        #         'araToken': '0#19171640662225970131824980691640846029429745GC1F1BF386B34F4C680DD7B7D2938FA61F3FF27',
        #         'araScene': 'login',
        #         'captchaImgCode': '',
        #         'captchaSessionId': '',
        #         'type': 'h5'}
        # r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
        #     'Origin': 'https://login.dingtalk.com',
        #     'Referer': 'https://login.dingtalk.com/'}
        # req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        # req = req.json()
        # # print(req)
        # req_url = req['data']
        # loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
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
        # # print('******已获取loginTmpCode值: ' + str(loginTmpCode))
        # time.sleep(1)
        #"https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode&loginTmpCode=59d3a6ee423937ebab33d44b476007a4"

        loginTmpCode = login_TmpCode
        # print('二、请求-后台登录页面......')
        url = r'http://gwms-v3.giikin.cn/tool/dingtalk_service/gettempcodebylogin'
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
        print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        # print('三、dingtalk_service服务器......')
        # print('（3.0）加载： ' + str(req.text))
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        gimp = req.headers['Location']

        time.sleep(1)
        print('（3.1）加载： ' + str(gimp))
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        # print(req)

        print('（3.2）加载： http://gwms-v3.giikin.cn/admin/index/index')
        url = 'http://gwms-v3.giikin.cn/admin/index/index'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': gimp}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

        # 查询订单更新 以订单编号 （单点系统）
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

        loginTmpCode = login_TmpCode
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
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
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
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)

        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)

        # print(req)
        # print(req.headers)

        print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)

        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('4.2、加载： ' + 'https://gimp.giikin.com')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        # url = 'https://gimp.giikin.com' + req.headers['Location']
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
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

        loginTmpCode = login_TmpCode
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
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
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
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)

        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('3、加载： ' + 'http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode?')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)

        # print(req)
        # print(req.headers)

        print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)

        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('4.2、加载： ' + 'https://gimp.giikin.com')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        index = req.headers['Location']
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('5、加载： ' + 'https://gimp.giikin.com/portal/index/index.html')
        # url = 'https://gimp.giikin.com' + req.headers['Location']
        url = 'https://gimp.giikin.com' + index
        # print(url)
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('6、加载： ' + 'https://gsso.giikin.com/admin/login/index.html')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False, proxies=proxies)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('7、加载： ' + 'https://gimp.giikin.com/portal/index/index.html?_ticker')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        # req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        proxies = {'http': 'socks5://' + proxy_id, 'https': 'socks5://' + proxy_id}         # 使用代理服务器
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

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)


    def sso__online_handle2(self, login_TmpCode):  # 手动输入token 登录系统保持会话状态
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

        loginTmpCode = login_TmpCode
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
        print(req)
        print(req.headers)

        print('4.1、加载： ' + 'https://gsso.giikin.com:443/admin/dingtalk_service/getunionidbytempcode?')
        index_system3 = req.headers['Location']
        # print(index_system3)
        url = index_system3.replace(':443', '')
        print(url)
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
        # url = 'https://gimp.giikin.com' + req.headers['Location']
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

        print('++++++已成功登录++++++' + str(req))
        print(datetime.datetime.now())
        print('*' * 100)

    # 查询订单更新 以订单编号（单点系统）
    def updata(self, sql, sql2, team,data_df,data_df2, login_TmpCode , handle):
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # self.sso_online_Two()
        # self.sso__online_handle(login_TmpCode)
        if handle == '手动':
            self.sso__online_handle(login_TmpCode)
        else:
            self.sso__online_auto()
        print('正在获取需 更新订单信息…………')
        start = datetime.datetime.now()
        db_data = pd.read_sql_query(sql=sql, con=self.engine1)
        if db_data.empty:
            print('无需要更新订单信息！！！')
            return
        print(db_data['订单编号'][0])
        orderId = list(db_data['订单编号'])
        max_count = len(orderId)  # 使用len()获取列表的长度，上节学的
        if max_count > 500:
            ord = ', '.join(orderId[0:500])
            df = self._updata(ord,data_df,data_df2)
            dlist = []
            n = 0
            while n < max_count - 500:  # 这里用到了一个while循环，穿越过来的
                n = n + 500
                ord = ','.join(orderId[n:n + 500])
                data = self._updata(ord,data_df,data_df2)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        else:
            ord = ','.join(orderId[0:max_count])
            dp = self._updata(ord,data_df,data_df2)
        print('正在写入临时缓存表......')
        dp.to_sql('cache', con=self.engine1, index=False, if_exists='replace')
        dp.to_excel('F:\\输出文件\\订单-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        print('查询已导出+++')
        print('正在更新订单跟进表中......')
        pd.read_sql_query(sql=sql2, con=self.engine1, chunksize=10000)
        print('更新耗时：', datetime.datetime.now() - start)
        # 更新订单跟进 的状态信息
    def _updata(self, ord,data_df,data_df2):
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500, 'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None,'email': None,
                'ip': None, 'productIds': None, 'saleIds': None, 'payType': None, 'logisticsId': None,'logisticsStyle':None,
                'logisticsMode': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None, 'emailStatus':None,
                'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': '','warehouse': None,
                'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None,'tuan': None, 'tuanStatus': None,
                'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None,'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None,
                'weightStart': None, 'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None,
                'sortField': None, 'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None}
        data.update({'orderPrefix': ord,
                     'shippingNumber': None})
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersdict = []
        # print('正在处理json数据转化为dataframe…………')
        try:
            for result in req['data']['list']:
                result['saleId'] = 0  # 添加新的字典键-值对，为下面的重新赋值用
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
            df = data[data_df]
            df.columns = data_df2
        except Exception as e:
            print('------查询为空')
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return df

    # 查询压单更新 以订单编号（仓储的获取）
    def updata_yadan(self, sql, sql2, team, data_df, data_df2,login_TmpCode,handle):  # 进入压单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # self.sso_online_cang()
        # self.sso_online_cang_handle(login_TmpCode)
        if handle == '手动':
            self.sso_online_cang_handle(login_TmpCode)
        else:
            self.sso_online_cang_auto()

        print('正在获取需 更新订单信息…………')
        start = datetime.datetime.now()
        db_data = pd.read_sql_query(sql=sql, con=self.engine1)
        if db_data.empty:
            print('无需要更新订单信息！！！')
            return
        print(db_data['订单编号'][0])
        orderId = list(db_data['订单编号'])
        max_count = len(orderId)  # 使用len()获取列表的长度，上节学的
        print('查询压单:' + str(max_count))
        if max_count > 500:
            ord = "', '".join(orderId[0:500])
            df = self._updata_yadan(ord, data_df, data_df2)
            dlist = []
            n = 0
            while n < max_count - 500:  # 这里用到了一个while循环，穿越过来的
                n = n + 500
                ord = "', '".join(orderId[n:n + 500])
                data = self._updata_yadan(ord, data_df, data_df2)
                print(data)
                if data is not None and len(data) > 0:
                    dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        else:
            ord = "', '".join(orderId[0:max_count])
            dp = self._updata_yadan(ord, data_df, data_df2)
        print(dp)
        if dp is None or len(dp) == 0:
            print('查询为空，不需更新+++')
        else:
            print('正在写入临时缓存表......')
            dp.to_sql('cache', con=self.engine1, index=False, if_exists='replace')
            print('查询已导出+++')
            sql = '''SELECT DISTINCT c.*,DATEDIFF(curdate(),入库时间) 压单天数,IF(DATEDIFF(curdate(),入库时间) > 5,'5天以前',null) AS 5天以前,
                                IF(物流 LIKE '%速派%','台湾-速派-新竹&711超商',
                                IF(物流 LIKE '%天马%','台湾-天马-新竹&711',
                                IF(物流 LIKE '%优美宇通%' or 物流 LIKE '%铱熙无敌%','台湾-铱熙无敌-新竹普货&特货',物流))) AS 物流方式
                    FROM `cache` c
                    LEFT JOIN gat_waybill_list g ON c.订单编号 = g.订单编号;'''
            df1 = pd.read_sql_query(sql=sql, con=self.engine1)
            df1.to_excel('F:\\输出文件\\压单反馈-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            print('正在更新订单跟进表中......')
            pd.read_sql_query(sql=sql2, con=self.engine1, chunksize=10000)
            print('更新耗时：', datetime.datetime.now() - start)
    def _updata_yadan(self, ord,data_df,data_df2):  # 进入压单检索界面
        print('+++正在查询订单信息中')
        timeStart = ((datetime.datetime.now() + datetime.timedelta(days=1)) - relativedelta(months=2)).strftime('%Y-%m-%d')
        timeEnd = (datetime.datetime.now()).strftime('%Y-%m-%d')
        url = r'http://gwms-v3.giikin.cn/order/pressure/index2'
        r_header = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
            'origin': 'http://gwms-v3.giikin.cn',
            'Referer': 'http://gwms-v3.giikin.cn/order/pressure/index2'}
        data = {'page': 1,
                'limit': 500,
                'startDate': timeStart + ' 00:00:00',
                'endDate': timeEnd + ' 23:59:59',
                'selectStr': "1=1  and a.reason in (1,2,3,4) and b.order_number in ('" + ord + "')"
                }
        # print(data)
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型 或者 str字符串  数据转换为dict字典
        max_count = req['count']
        print(max_count)
        if max_count != [] and max_count != 0:
            ordersdict = []
            try:
                for result in req['data']:
                    ordersdict.append(result)
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            data = pd.json_normalize(ordersdict)
            data = data[data_df]
            data.columns = data_df2
        else:
            data = None
            print('****** 没有信息！！！')
        return data

    # 查询出库更新 以订单编号（仓储的获取）
    def updata_chuku(self, sql, sql2, team, data_df, data_df2,login_TmpCode,handle):  # 进入压单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # self.sso_online_cang()
        # self.sso_online_cang_handle(login_TmpCode)
        if handle == '手动':
            self.sso_online_cang_handle(login_TmpCode)
        else:
            self.sso_online_cang_auto()

        print('正在获取需 更新订单信息…………')
        start = datetime.datetime.now()
        db_data = pd.read_sql_query(sql=sql, con=self.engine1)
        if db_data.empty:
            print('无需要更新订单信息！！！')
            return
        print(db_data['订单编号'][0])
        orderId = list(db_data['订单编号'])
        max_count = len(orderId)  # 使用len()获取列表的长度，上节学的
        if max_count > 500:
            ord = "', '".join(orderId[0:500])
            df = self._updata_chuku(ord, data_df, data_df2)
            dlist = []
            n = 0
            while n < max_count - 500:  # 这里用到了一个while循环，穿越过来的
                n = n + 500
                ord = "', '".join(orderId[n:n + 500])
                data = self._updata_chuku(ord, data_df, data_df2)
                print(data)
                if data is not None and len(data) > 0:
                    dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        else:
            print(22)
            ord = "','".join(orderId[0:max_count])
            dp = self._updata_chuku(ord, data_df, data_df2)
        if dp is None or len(dp) == 0:
            print('查询为空，不需更新+++')
        else:
            print('正在写入临时缓存表......')
            print(dp)
            print('查询已导出+++')
            dp.to_excel('F:\\输出文件\\出库-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            dp.to_sql('cache', con=self.engine1, index=False, if_exists='replace')
            print('正在更新订单跟进表中......')
            pd.read_sql_query(sql=sql2, con=self.engine1, chunksize=10000)
            print('更新耗时：', datetime.datetime.now() - start)
        # 进入运单扫描导出 界面
    def _updata_chuku(self, ord,data_df,data_df2):
        print('+++正在查询订单信息中')
        timeStart = ((datetime.datetime.now() + datetime.timedelta(days=1)) - relativedelta(months=2)).strftime('%Y-%m-%d')
        timeEnd = (datetime.datetime.now()).strftime('%Y-%m-%d')
        url = r'http://gwms-v3.giikin.cn/order/delivery/deliverylog'
        r_header = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
            'origin': 'http://gwms-v3.giikin.cn',
            'Referer': 'http://gwms-v3.giikin.cn/order/order/shelves'}
        data = {'page': 1,
                'limit': 500,
                'startDate': timeStart + ' 00:00:00',
                'endDate': timeEnd + ' 23:59:59',
                'selectStr': "1=1 and bs.order_number in ('" + ord + "')"
                }
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型 或者 str字符串  数据转换为dict字典
        max_count = req['count']
        print(max_count)
        if max_count != [] and max_count != 0:
            ordersdict = []
            try:
                for result in req['data']:
                    ordersdict.append(result)
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            data = pd.json_normalize(ordersdict)
            data = data[data_df]
            data.columns = data_df2
        else:
            data = None
            print('****** 没有信息！！！')
        return data

    # 查询提货更新 以订单编号（仓储的获取）
    def updata_tihuo(self, sql, sql2, team, data_df, data_df2):  # 进入压单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        self.sso_online_cang()
        print('正在获取需 更新订单信息…………')
        start = datetime.datetime.now()
        db = pd.read_sql_query(sql=sql, con=self.engine1)
        if db.empty:
            print('无需要更新订单信息！！！')
            return
        print(db['订单编号'][0])
        orderId = list(db['订单编号'])
        max_count = len(orderId)  # 使用len()获取列表的长度，上节学的
        if max_count > 0:
            df = pd.DataFrame([['','','','','','','']], columns=data_df2)
            dlist = []
            for ord in orderId:
                print(ord)
                data = self._updata_tihuo(ord, data_df, data_df2)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
        else:
            dp = None
        print('正在写入临时缓存表......')
        print(dp)
        dp.to_excel('F:\\输出文件\\提货-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
        if dp.empty:
            print('查询为空，不需更新+++')
        else:
            # dp.to_sql('cache', con=self.engine1, index=False, if_exists='replace')
            print('查询已导出+++')
            # dp.to_excel('F:\\输出文件\\提货-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            # print('正在更新订单跟进表中......')
            # pd.read_sql_query(sql=sql2, con=self.engine1, chunksize=10000)
        print('更新耗时：', datetime.datetime.now() - start)
        # 进入运单扫描导出 界面
    def _updata_tihuo(self, ord, data_df, data_df2):
        print('+++正在查询订单信息中')
        timeStart = ((datetime.datetime.now() + datetime.timedelta(days=1)) - relativedelta(months=2)).strftime('%Y-%m-%d')
        timeEnd = (datetime.datetime.now()).strftime('%Y-%m-%d')
        # timeStart = '2022-04-22'
        # timeEnd = '2022-04-22'
        url = r'http://gwms-v3.giikin.cn/roo/meta/page?'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gwms-v3.giikin.cn/roo/meta/index/listId/46'}
        data = {'sEcho': 12, 'iColumns': 11, 'sColumns': ',, , , , , , , , ,', 'iDisplayStart': 0, 'iDisplayLength': 10, 'mDataProp_0': 'id', 'sSearch_0': None, 'bRegex_0': False,
                'bSearchable_0': True, 'bSortable_0': False,'mDataProp_1': 'billno', 'sSearch_1': None, 'bRegex_1': False, 'bSearchable_1': True, 'bSortable_1': True, 'mDataProp_2': 'order_number',
                'sSearch_2': None, 'bRegex_2': False, 'bSearchable_2': True, 'bSortable_2': False, 'mDataProp_3': 'result', 'sSearch_3': None, 'bRegex_3': False, 'bSearchable_3': True,
                'bSortable_3': False, 'mDataProp_4': 'uid', 'sSearch_4': None, 'bRegex_4': False, 'bSearchable_4': True, 'bSortable_4': False, 'mDataProp_5': 'country_code',
                'sSearch_5': None,'bRegex_5': False, 'bSearchable_5': True, 'bSortable_5': False, 'mDataProp_6': 'intime', 'sSearch_6': None, 'bRegex_6': False, 'bSearchable_6': True,
                'bSortable_6': False, 'mDataProp_7': 'logistics_id',  'sSearch_7':None, 'bRegex_7': False,  'bSearchable_7': True, 'bSortable_7': False, 'mDataProp_8': 'country',
                'sSearch_8': None, 'bRegex_8': False, 'bSearchable_8': True,  'bSortable_8': False, 'mDataProp_9': 'is_exception',  'sSearch_9': None, 'bRegex_9': False,
                'bSearchable_9': True, 'bSortable_9': False, 'mDataProp_10': 'is_deal', 'sSearch_10': None, 'bRegex_10': False,  'bSearchable_10': True,  'bSortable_10': False,
                'sSearch': None, 'bRegex': False, 'iSortCol_0': 0, 'sSortDir_0': 'desc','iSortingCols': 1, 'listId': 46,
                'startDate': timeStart + ' 00:00:00',
                'endDate': timeEnd + ' 23:59:59',
                'queryStr': 'a.order_number=' + "'" + ord + "'"
                # '_': 1650620225353
                }
        # print(data)
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        print(req)
        req = json.loads(req.text)  # json类型 或者 str字符串  数据转换为dict字典
        print(req)
        max_count = req['iTotalRecords']
        print(max_count)
        if max_count != [] and max_count != '0' and max_count != 0:
            ordersdict = []
            try:
                req_data = req['aaData']
                result = next(reversed(req_data))
                # for result in req['aaData']:
                # 方法三（最佳方法）
                # next(reversed(od))  # get the last key
                # next(reversed(od.items()))  # get the last item
                # next(iter(od))  # get the first key
                # next(iter(od.items()))  # get the first item
                ordersdict.append(result)
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
            data = pd.json_normalize(ordersdict)
            data = data[data_df]
            data.columns = data_df2
            print(data)
        else:
            data = None
            print('****** 没有信息！！！')
        return data



    def test(self):
        sql = '''SELECT 订单编号 FROM customer;'''
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        print(df)
        df.to_sql('cache', con=self.engine1, index=False, if_exists='replace')
if __name__ == '__main__':
    m = Settings_sso()
    # m.sso_online_Two_Five()
    # m.sso__online_auto22()
    m.sso__online_auto()
    # m.test()
    # m._updata_tihuo('GT203302314025681','','')


    '''
    1、单点后台请求网站：
    https://login.dingtalk.com/login/index.htm?goto=https%3A%2F%2Foapi.dingtalk.com%2Fconnect%2Foauth2%2Fsns_authorize%3Fappid%3Ddingoajqpi5bp2kfhekcqm%26response_type%3Dcode%26scope%3Dsnsapi_login%26state%3DSTATE%26redirect_uri%3Dhttps%3A%2F%2Fgsso.giikin.com%2Fadmin%2Fdingtalk_service%2Fgetunionidbytempcode
    2、后台请求获取值：loginTmpCode=971701a5cf0230e9a685e4a651cc82e1
      $.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
            data: { mobile: '+86-17596568562',
                    pwd: 'xhy123456',
                    goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                    pdmToken: '',
                    araAppkey: '1917',
                    araToken: '0#19171651897715811055201302751651976157916999GD771245699468C2D36034C0D1CB3A896998EA5',
                    araScene: 'login',
                    captchaImgCode: '',
                    captchaSessionId: '',
                    type: 'h5'
                },
                type: 'POST',
                timeout: '10000',
                beforeSend(xhr, settings) {
                    xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                },
                success: function(data) {
                    if (data.success) {
                         console.log(data.data)
                         console.log(data)
                         console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1])
                    } else {
                            console.log(data.code)
                    }
                },
                error: function(error) {
                    alert("请检查网络");
                }
            })
            
            
            
    3、仓储后台请求网站：
    https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode
    4、后台请求获取值：loginTmpCode=971701a5cf0230e9a685e4a651cc82e1
      $.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
            data: { mobile: '+86-17596568562',
                    pwd: 'xhy123456',
                    goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoag6pwcnuxvwto821j&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gwms-v3.giikin.cn/tool/dingtalk_service/getunionidbytempcode',
                    pdmToken: '',
                    araAppkey: '1917',
                    araToken: '0#19171651897715811055201302751651976157916999GD771245699468C2D36034C0D1CB3A896998EA5',
                    araScene: 'login',
                    captchaImgCode: '',
                    captchaSessionId: '',
                    type: 'h5'
                },
                type: 'POST',
                timeout: '10000',
                beforeSend(xhr, settings) {
                    xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                },
                success: function(data) {
                    if (data.success) {
                         console.log(data.data)
                         console.log(data)
                         console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1])
                    } else {
                            console.log(data.code)
                    }
                },
                error: function(error) {
                    alert("请检查网络");
                }
            })
            
    '''