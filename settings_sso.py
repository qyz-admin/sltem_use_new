import datetime
import time
import win32api,win32con
import sys
import requests

class Settings_sso():
    def __init__(self):
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.userMobile = '+86-18538110674'
        self.password = 'qyz04163510'
        self.userID = '1343'

    def sso_online_Two(self):  # 登录系统保持会话状态
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
        print('二、请求-后台登录页面......')
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
        print('******获取登录页面url成功： /oapi.dingtalk.com/connect/oauth2/sns_authorize?')

        time.sleep(1)
        print('三、dingtalk_service服务器......')
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
        print('（3.2）再次加载/dingtalk_service/getunionidbytempcode?code=页面......')
        url = index_system3
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        print('（3.3）加载//gimp.giikin.com 页面......')
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
        print('四、加载/gimp.giikin.com/portal/index/index.html 页面......')
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
        print('（4.1）加载index.html?_system=18页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('（4.2）加载/gimp.giikin.com/portal/index/index.html?_ticker=页面......')
        url = req.headers['Location']
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        # print(6060)
        # print(req)
        # print(req.headers)

        time.sleep(1)
        print('（4.3）加载/gimp.giikin.com:443/portal/index/index.html页面......')
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
