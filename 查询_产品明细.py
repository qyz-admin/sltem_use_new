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

from sqlalchemy import create_engine
from settings import Settings
from settings_sso import Settings_sso
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色


# -*- coding:utf-8 -*-
class QueryTwoT(Settings, Settings_sso):
    def __init__(self, userMobile, password, login_TmpCode, handle, proxy_handle, proxy_id):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self._online()
        # self.sso_online_Two()

        # self.sso__online_auto()
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
    #  登录后台中
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

    # 获取签收表内容
    def readFormHost(self, team):
        start = datetime.datetime.now()
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath, team)
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath, team):
        fileType = os.path.splitext(filePath)[1]
        app = xlwings.App(visible=False, add_book=False)
        app.display_alerts = False
        if 'xls' in fileType:
            wb = app.books.open(filePath, update_links=False, read_only=True)
            for sht in wb.sheets:
                try:
                    db = None
                    db = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                    print(db.columns)
                    columns_value = list(db.columns)                             # 获取数据的标题名，转为列表
                    for column_val in columns_value:
                        if '产品id' != column_val:
                            db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                    db.dropna(axis=0, how='any', inplace=True)                  # 空值（缺失值），将空值所在的行/列删除后
                except Exception as e:
                    print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                if db is not None and len(db) > 0:
                    # print(db)
                    rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
                    print('++++正在获取：' + sht.name + ' 表；共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                    productId = list(db['产品id'])
                    print(productId[0])
                    df = self.orderInfoQuery(productId[0])
                    print(df)
                    dlist = []
                    for proId in productId[1:]:
                        print(proId)
                        data = self.orderInfoQuery(proId)
                        dlist.append(data)
                    dp = df.append(dlist, ignore_index=True)
                    dp.to_excel('G:\\输出文件\\产品检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
                    print('查询已导出+++')
                else:
                    print('----------数据为空,查询失败：' + sht.name)
            wb.close()
        app.quit()

    # 查询更新（新后台的获取）
    def orderInfoQuery(self, proId):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getProductList&page=1&pageSize=10&productName=&status=&source=&isSensitive=&isGift=&isDistribution=&chooserId=&buyerId='
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsProductSearch'}
        data = {}
        data.update({'productId': proId})
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        ordersDict = []
        data = None
        if req['data']['list'] != [] and req['data']['count'] != 0:
            try:
                for result in req['data']['list']:
                    # 添加新的字典键-值对，为下面的重新赋值用
                    result['cate_id'] = 0
                    result['second_cate_id'] = 0
                    result['third_cate_id'] = 0
                    result['cate_id'] = (result['categorys']).split('>')[2]
                    result['second_cate_id'] = (result['categorys']).split('>')[1]
                    result['third_cate_id'] = (result['categorys']).split('>')[0]
                    ordersDict.append(result)
            except Exception as e:
                print('转化失败： 重新获取中', str(Exception) + str(e))
                self.orderInfoQuery(ord)
            data = pd.json_normalize(ordersDict)
            data['name'] = data['name'].str.strip()
            data['cate_id'] = data['cate_id'].str.strip()
            data['second_cate_id'] = data['second_cate_id'].str.strip()
            data['third_cate_id'] = data['third_cate_id'].str.strip()
            data = data[['id', 'name', 'cate_id', 'second_cate_id', 'third_cate_id', 'status', 'price', 'selectionName',
                         'sellerCount', 'buyerName', 'saleCount', 'logisticsCost', 'lender', 'isGift', 'createTime',
                         'categorys', 'image']]
            data.columns = ['产品id', '产品名称', '一级分类', '二级分类', '三级分类', '产品状态', '价格(￥)', '选品人',
                            '供应商数', '采购人', '商品数', 'logisticsCost', '出借人', '特殊信息', '添加时间',
                            '产品分类', '产品图片']
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return data




    # 后台  补充  产品信息
    def productInfo(self, team, ordersDict):  # 进入查询界面，
        print('正在获取需要更新的产品id信息')
        start = datetime.datetime.now()
        month_begin = (datetime.datetime.now() - relativedelta(months=3)).strftime('%Y-%m-%d')
        sql = '''SELECT DISTINCT 产品id  FROM {0} sl 
          WHERE sl.`日期`>= '{1}' 
            AND (sl.`产品名称` IS NULL or sl.`父级分类` IS NULL)
            AND (sl.`系统订单状态` NOT IN ('已删除','问题订单','支付失败','未支付'));'''.format(team, month_begin)
        # ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        if ordersDict.empty:
            print('无需要更新的产品id信息！！！')
            return
        productId = list(ordersDict['产品id'])
        print('获取耗时：', datetime.datetime.now() - start)
        max_count = len(productId)
        print(max_count)
        if max_count > 0:
            df = pd.DataFrame([['', '', '', '', '']], columns=['id', 'name', 'cate_id', 'second_cate_id', 'third_cate_id'])
            dlist = []
            for proId in productId:
                print(proId)
                data = self.productQuery(proId)
                dlist.append(data)
            dp = df.append(dlist, ignore_index=True)
            dp.to_sql('tem_product', con=self.engine1, index=False, if_exists='replace')
            print('更新中......')
            sql = '''update tem_product_cp a, tem_product b
                    set a.`产品名称`= IF(b.`name` = '', a.`产品名称`, b.`name`),
                        a.`父级分类`= IF(b.`cate_id` = '', a.`父级分类`, b.`cate_id`),
                        a.`二级分类`= IF(b.`second_cate_id` = '', a.`二级分类`, b.`second_cate_id`),
                        a.`三级分类`= IF(b.`third_cate_id` = '', a.`三级分类`, b.`third_cate_id`)
                  where a.`产品id`= b.`id`;'''.format(month_begin, team)
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)

            sql = '''update {1} a, tem_product_cp b
                    set a.`产品名称`= IF(b.`产品名称` = '', a.`产品名称`, b.`产品名称`),
                        a.`父级分类`= IF(b.`父级分类` = '', a.`父级分类`, b.`父级分类`),
                        a.`二级分类`= IF(b.`二级分类` = '', a.`二级分类`, b.`二级分类`),
                        a.`三级分类`= IF(b.`三级分类` = '', a.`三级分类`, b.`三级分类`)
                  where a.`订单编号`= b.`订单编号`;'''.format(month_begin, team)
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('共有 ' + str(len(dp)) + '条 成功更新+++++++')

    def productQuery(self, proId):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getProductList&page=1&pageSize=10&productName=&status=&source=&isSensitive=&isGift=&isDistribution=&chooserId=&buyerId='
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsProductSearch'}
        data = {}
        data.update({'productId': proId})
        proxy = '192.168.13.89:37467'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        # req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:          # 添加新的字典键-值对，为下面的重新赋值用
                result['cate_id'] = 0
                result['second_cate_id'] = 0
                result['third_cate_id'] = 0
                result['cate_id'] = (result['categorys']).split('>')[2]
                result['second_cate_id'] = (result['categorys']).split('>')[1]
                result['third_cate_id'] = (result['categorys']).split('>')[0]
                ordersDict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            self.orderInfoQuery(ord)
        data = pd.json_normalize(ordersDict)
        try:
            data['name'] = data['name'].str.strip()
            data['cate_id'] = data['cate_id'].str.strip()
            data['second_cate_id'] = data['second_cate_id'].str.strip()
            data['third_cate_id'] = data['third_cate_id'].str.strip()
            data = data[['id', 'name', 'cate_id', 'second_cate_id', 'third_cate_id']]
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        # print('++++++本次查询成功+++++++')
        print('*' * 50)
        return data


if __name__ == '__main__':
    m = QueryTwoT('+86-18538110674', 'qyz35100416')
    start: datetime = datetime.datetime.now()
    match1 = {'gat': '港台', 'gat_order_list': '港台', 'slsc': '品牌'}
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    # 1、手动导入状态
    for team in ['gat']:
        m.readFormHost(team)
    # m.productInfo('gat')

    print('查询耗时：', datetime.datetime.now() - start)