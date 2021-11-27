import datetime
import time
import win32api,win32con
import sys


class Settings():
    def __init__(self):
        self.excelPath = r'D:\Users\Administrator\Desktop\直发表'
        self.mysql1 = {'host': 'localhost',      #数据库地址
                      'user': 'root',           #数据库账户
                      'port': '3306',
                      'password': '123456',     #数据库密码   654321
                      'datebase': 'logistics_status',   #数据库名称
                      'charset': 'utf8'         #数据库编码
                       }
        self.mysql200 = {'host': 'tidb.giikin.com',  # 数据库地址
                       'user': 'shenlongkf',  # 数据库账户
                       'port': '4000',
                       'password': 'SIK87&67asd',  # 数据库密码
                       'datebase': 'gdqs_shenlong',  # 数据库名称
                       'charset': 'utf8'  # 数据库编码
                       }

        def _online_Two(self):  # 登录系统保持会话状态
            print(datetime.datetime.now())
            print('正在登录后台系统中......')
            print('一、获取-钉钉用户信息......')
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
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
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
            print('******已获取loginTmpCode值: ' + str(loginTmpCode))

            time.sleep(1)
            print('二、请求-后台登录页面......')
            url = r'https://gsso.giikin.com/admin/dingtalk_service/gettempcodebylogin.html'
            data = {'tmpCode': loginTmpCode,
                    'system': 18,
                    'url': '',
                    'ticker': '',
                    'companyId': 1}
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                'Origin': 'https://login.dingtalk.com',
                'Referer': 'http://gsso.giikin.com/admin/login/logout.html'}
            req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
            # print(req.text)
            print('******请求登录页面url成功： ' + str(req.text))

            time.sleep(1)
            print('三、dingtalk_service服务器......')
            # print('（一）加载dingtalk_service跳转页面......')
            url = req.text
            data = {'tmpCode': loginTmpCode,
                    'system': 1,
                    'url': '',
                    'ticker': '',
                    'companyId': 1}
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
            # print(req.headers)
            gimp = req.headers['Location']
            print('******已获取跳转页面： ' + str(gimp))
            time.sleep(1)
            # print('（二）请求dingtalk_service的cookie值......')
            url = gimp
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
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
            print(808080)
            # print(index_system3)
            index_system3 = index_system3.replace(':443', '')
            print(index_system3)
            # 跳转使用-暂停

            time.sleep(1)
            # print('（三）加载index.html?_ticker=页面......')
            url = index_system3
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            print(req)
            print(req.headers)

            print(990099900)
            time.sleep(1)
            # print('（三）加载index.html?_ticker=页面......')
            url = req.headers['Location']
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
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
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            print(req.headers)
            index_system = req.headers['Location']
            print('+++已获取index.html?_system=18正式页面')
            print(7070)

            time.sleep(1)
            # print('（三）加载index.html?_ticker=页面......')
            url = index_system
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            print(req)
            print(req.headers)

            time.sleep(1)
            # print('（三）加载index.html?_ticker=页面......')
            url = req.headers['Location']
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            print(6060)
            print(req)
            print(req.headers)

            time.sleep(1)
            # print('（三）加载index.html?_ticker=页面......')
            url = req.headers['Location']
            r_header = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                'Referer': 'http://gsso.giikin.com/'}
            req = self.session.get(url=url, headers=r_header, allow_redirects=False)
            print(5050)
            print(req)
            print(req.headers)

            print('++++++已成功登录++++++')
            print('*' * 50)