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

from sqlalchemy import create_engine
from settings import Settings
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, \
    Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色


# -*- coding:utf-8 -*-
class QueryTwo(Settings):
    def __init__(self):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self._online()
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

    def _online(self):  # 登录系统保持会话状态
        print('第一阶段获取-钉钉用户信息......')
        url = r'https://login.dingtalk.com/login/login_with_pwd'
        data = {'mobile': '+86-18538110674',
                'pwd': 'qyz04163510',
                'goto': 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                'pdmToken': '',
                'araAppkey': '1917',
                'araToken': '0#19171626965820522806235080231627186242421867G0E5F84086457BA90D5DFC95A2522D7EFFA6F79',
                'araScene': 'login',
                'captchaImgCode': '',
                'captchaSessionId': '',
                'type': 'h5'}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Origin': 'https://login.dingtalk.com',
                    'Referer': 'https://login.dingtalk.com/'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        req = req.json()
        print(req)
        req_url = req['data']
        loginTmpCode = req_url.split('loginTmpCode=')[1]        # 获取loginTmpCode值
        print(loginTmpCode)
        print('+++已获取loginTmpCode值+++')

        time.sleep(1)
        print('第二阶段请求-登录页面......')
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
        print(req.text)
        print('+++请求登录页面url成功+++')

        time.sleep(1)
        print('第三阶段请求-dingtalk服务器......')
        print('（一）加载dingtalk_service跳转页面......')
        url = req.text
        data = {'tmpCode': loginTmpCode,
                'system': 1,
                'url': '',
                'ticker': '',
                'companyId': 1}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, data=data, allow_redirects=False)
        print(req.headers)
        gimp = req.headers['Location']
        print('+++已获取跳转页面+++')
        time.sleep(1)
        print('（二）请求dingtalk_service的cookie值......')
        url = gimp
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req)
        print('+++已获取cookie值+++')

        time.sleep(2)
        print('第四阶段页面-重定向跳转中......')
        print('（一）加载chooselogin.html页面......')
        url = r'http://gsso.giikin.com/admin/login_by_dingtalk/chooselogin.html'
        data = {'user_id': 1343}
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': gimp,
                    'Origin': 'http://gsso.giikin.com'}
        req = self.session.post(url=url, headers=r_header, data=data, allow_redirects=False)
        print(req.headers)
        index = req.headers['Location']
        print('+++已获取gimp.giikin.com页面')
        time.sleep(2)
        print('（二）加载gimp.giikin.com页面......')
        url = index
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': index}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        index2 = req.headers['Location']
        print('+++已获取index.html页面')

        time.sleep(2)
        print('（三）加载index.html页面......')
        url = 'http://gimp.giikin.com/' + index2
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        index_system = req.headers['Location']
        print('+++已获取index.html?_system=18正式页面')

        time.sleep(2)
        print('第五阶段正式页面-重定向跳转中......')
        print('（一）加载index.html?_system页面......')
        url = index_system
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req.headers)
        index_system2 = req.headers['Location']
        print('+++已获取index.html?_ticker=页面......')
        time.sleep(2)
        print('（二）加载index.html?_ticker=页面......')
        url = index_system2
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'Referer': 'http://gsso.giikin.com/'}
        req = self.session.get(url=url, headers=r_header, allow_redirects=False)
        print(req)
        print(req.headers)
        print('++++++已成功登录++++++')


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
            path = r'D:\Users\Administrator\Desktop\需要用到的文件\品牌数据源'
        else:
            path = r'D:\Users\Administrator\Desktop\需要用到的文件\数据库'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath, team, query)
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
                    columns = list(db.columns)  # 获取数据的标题名，转为列表
                    columns_value = ['商品链接', '规格(中文)', '收货人', '拉黑率', '电话长度', '邮编长度', '配送地址', '地址翻译',
                                     '邮箱', '留言', '审核人', '预选物流公司(新)', '是否api下单', '特价', '异常提示', '删除原因',
                                     '备注', '删除人', 'IP', '超商店铺ID', '超商店铺名称', '超商网点地址', '超商类型',
                                     '市/区', '优化师', '问题原因', '审单类型', '是否克隆', '代下单客服',
                                     '改派的原运单号', '改派的下架时间']
                    for column_val in columns_value:
                        if column_val in columns:
                            db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
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
            				        IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '10', '1', '2')) 旬,
            			            DATE(h.下单时间) 日期,
            				        h.运营团队 团队,
            				        IF(h.`币种` = '台币', 'TW', IF(h.`币种` = '港币', 'HK', h.`币种`)) 区域,
            				        IF(h.`币种` = '台币', '台湾', IF(h.`币种` = '港币', '香港', h.`币种`)) 币种,
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
            				        IF(h.`物流状态` in ('发货中'), '在途', h.`物流状态`) 系统物流状态,
            				        h.上线时间 上线时间
                            FROM d1_host h 
                            LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                            LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                            LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`; '''.format(team)
        elif team in ('slsc'):
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
			                    IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '10', '1', '2')) 旬,
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
				                IF(h.`物流状态` in ('发货中'), '在途', h.`物流状态`) 系统物流状态,
            				    h.上线时间 上线时间
                    FROM d1_host h 
                    LEFT JOIN dim_product_slsc ON  dim_product_slsc.id = h.产品id
            --        LEFT JOIN (SELECT * FROM dim_product WHERE id IN (SELECT MAX(id) FROM dim_product GROUP BY id ) ORDER BY id) e on e.id = h.产品id
                    LEFT JOIN dim_cate ON  dim_cate.id = dim_product_slsc.third_cate_id
                    LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team in ('slrb_jl', 'slrb_js', 'slrb_hs'):
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
			                    IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '10', '1', '2')) 旬,
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
				                IF(h.`物流状态` in ('发货中'), '在途', h.`物流状态`) 系统物流状态,
            				    h.上线时间 上线时间
                    FROM d1_host h 
                    LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                    LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                    LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team == 'slrb':
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
        			                    IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '10', '1', '2')) 旬,
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
        				                IF(h.`物流状态` in ('发货中'), '在途', h.`物流状态`) 系统物流状态,
            				            h.上线时间 上线时间 
                            FROM d1_host h 
                            LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                            LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                            LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team == 'sltg':
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
                                IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '10', '1', '2')) 旬,
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
				                IF(h.`物流状态` in ('发货中'), '在途', h.`物流状态`) 系统物流状态,
            				    h.上线时间 上线时间
                    FROM d1_host h 
                    LEFT JOIN dim_product ON  dim_product.sale_id = h.商品id
                    LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                    LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = h.`物流渠道`;'''.format(team)
        elif team in ('slxmt', 'slxmt_t', 'slxmt_hfh'):
            sql = '''SELECT EXTRACT(YEAR_MONTH  FROM h.下单时间) 年月,
                            IF(DAYOFMONTH(h.`下单时间`) > '20', '3', IF(DAYOFMONTH(h.`下单时间`) < '10', '1', '2')) 旬,
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
                            IF(h.`物流状态` in ('发货中'), '在途', h.`物流状态`) 系统物流状态,
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
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=2000)
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
                                    set a.`币种`=b.`币种`,
                                        a.`数量`=b.`数量`,
            		                    a.`电话号码`=b.`电话号码` ,
            		                    a.`运单编号`= IF(b.`运单编号` = '', NULL, b.`运单编号`),
            		                    a.`系统订单状态`= IF(b.`系统订单状态` = '', NULL, b.`系统订单状态`),
            		                    a.`系统物流状态`= IF(b.`系统物流状态` = '', NULL, b.`系统物流状态`),
            		                    a.`是否改派`= b.`是否改派`,
            		                    a.`物流方式`= IF(b.`物流方式` = '', NULL, b.`物流方式`),
            		                    a.`物流名称`= b.`物流名称`,
            		                    a.`货物类型`= b.`货物类型`,
            		                    a.`审核时间`= b.`审核时间`,
            		                    a.`仓储扫描时间`= b.`仓储扫描时间`,
            		                    a.`上线时间`= b.`上线时间`,
            		                    a.`完结状态时间`= b.`完结状态时间`
            		                where a.`订单编号`= b.`订单编号`;'''.format(team)
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
            except Exception as e:
                print('更新失败：', str(Exception) + str(e))
            print('更新成功…………')




    # 更新团队订单明细（新后台的获取）
    def orderInfo(self, searchType, team, last_month):  # 进入订单检索界面，
        print('>>>>>>正式查询中<<<<<<')
        print('正在获取需要订单信息......')
        start = datetime.datetime.now()
        sql = '''SELECT id,`订单编号`  FROM {0}_order_list sl WHERE sl.`日期` = '{1}';'''.format(team, last_month)
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
            self.orderInfoQuery(ord, searchType, team)
        print('单日查询耗时：', datetime.datetime.now() - start)

    def orderInfoQuery(self, ord, searchType, team):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
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
        print('+++已成功发送请求......')
        print('正在处理json数据转化为dataframe…………')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        ordersDict = []
        for result in req['data']['list']:
            # print(result)
            try:
                # 添加新的字典键-值对，为下面的重新赋值用
                result['saleId'] = 0
                result['saleProduct'] = 0
                result['productId'] = 0
                result['spec'] = 0
                result['saleId'] = result['specs'][0]['saleId']
                result['saleProduct'] = (result['specs'][0]['saleProduct']).split('#')[2]
                result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
                result['spec'] = result['specs'][0]['spec']
            except Exception as e:
                print('转化失败：', str(Exception) + str(e) + str(result['orderNumber']))
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
            self.q.put(result)
        # print(len(req['data']['list']))
        for i in range(len(req['data']['list'])):
            ordersDict.append(self.q.get())
        data = pd.json_normalize(ordersDict)
        # print(data.columns)
        print('正在写入缓存中......')
        try:
            df = data[['orderNumber', 'currency', 'area', 'shipInfo.shipPhone', 'wayBillNumber', 'saleId', 'saleProduct', 'productId', 'spec', 'quantity',
                       'orderStatus', 'logisticsStatus', 'logisticsName', 'addTime', 'verifyTime', 'transferTime', 'onlineTime', 'deliveryTime', 'finishTime',
                       'logisticsUpdateTime', 'reassignmentTypeName', 'dpeStyle', 'amount', 'payType', 'weight']]
            print(df)
            print('正在更新临时表中......')
            df.to_sql('d1_cpy', con=self.engine1, index=False, if_exists='replace')
            sql = '''SELECT DATE(h.addTime) 日期,
            				    IF(h.`currency` = '日币', '日本', IF(h.`currency` = '泰铢', '泰国', IF(h.`currency` = '港币', '香港', IF(h.`currency` = '台币', '台湾', IF(h.`currency` = '韩元', '韩国', h.`currency`))))) 币种,
            				    h.orderNumber 订单编号,
            				    h.quantity 数量,
            				    h.`shipInfo.shipPhone` 电话号码,
            				    h.wayBillNumber 运单编号,
            				    h.orderStatus 系统订单状态,
            				    IF(h.`logisticsStatus` in ('发货中'), '在途', h.`logisticsStatus`) 系统物流状态,
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
            				    h.`weight` 包裹重量
                            FROM d1_cpy h
                                LEFT JOIN dim_product ON  dim_product.sale_id = h.saleId
                                LEFT JOIN dim_cate ON  dim_cate.id = dim_product.third_cate_id
                                LEFT JOIN dim_trans_way ON  dim_trans_way.all_name = TRIM(h.logisticsName);'''.format(team)
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            df.to_sql('d1_cpy_cp', con=self.engine1, index=False, if_exists='replace')
            print('正在更新表总表中......')
            sql = '''update {0}_order_list a, d1_cpy_cp b
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
                                a.`商品id`= IF(b.`商品id` = '', NULL, b.`商品id`),
                                a.`产品id`= IF(b.`产品id` = '', NULL, b.`产品id`),
                                a.`产品名称`= IF(b.`产品名称` = '', NULL, b.`产品名称`),
                                a.`审核时间`= IF(b.`审核时间` = '', NULL, b.`审核时间`),
                                a.`上线时间`= IF(b.`上线时间` = '', NULL, b.`上线时间`),
                                a.`仓储扫描时间`= IF(b.`仓储扫描时间` = '', NULL, b.`仓储扫描时间`),
                                a.`完结状态时间`= IF(b.`完结状态时间` = '', NULL, b.`完结状态时间`),
                                a.`包裹重量`= IF(b.`包裹重量` = '', NULL, b.`包裹重量`)
                    where a.`订单编号`=b.`订单编号`;'''.format(team)
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
        except Exception as e:
            print('更新失败：', str(Exception) + str(e))
        print('++++++本批次更新成功+++++++')

if __name__ == '__main__':
    m = QueryTwo()
    start: datetime = datetime.datetime.now()
    match1 = {'gat': '港台',
              'slsc': '品牌'}
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    # for team in ['slsc', 'gat','slgat', 'slgat_hfh', 'slgat_hs', slrb', 'slrb_jl', 'slrb_js']:
    # 1、手动导入状态
    # for team in ['gat']:
    #     query = '导入'         # 导入；，更新--->>数据更新切换
    #     m.readFormHost(team, query)
    # 2、手动更新状态
    # for team in ['gat']:
    #     query = '更新'         # 导入；，更新--->>数据更新切换
    #     m.readFormHost(team, query)




    # -----------------------------------------------系统导入状态运行（二）-----------------------------------------
    #   台湾token, 日本token, 新马token：  f5dc2a3134c17a2e970977232e1aae9b
    #   泰国token： 83583b29fc24ec0529082ff7928246a6

    # begin = datetime.date(2021, 6, 1)       # 1、手动设置时间；若无法查询，切换代理和直连的网络
    # print(begin)
    # end = datetime.date(2021, 7, 28)
    # print(end)

    yy = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y'))  # 2、自动设置时间
    mm = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%m'))
    begin = datetime.date(yy, mm, 1)
    print(begin)
    yy2 = int(datetime.datetime.now().strftime('%Y'))
    mm2 = int(datetime.datetime.now().strftime('%m'))
    dd2 = int(datetime.datetime.now().strftime('%d'))
    end = datetime.date(yy2, mm2, dd2)
    print(end)

    print(datetime.datetime.now())
    # for team in ['slsc', 'gat']:
    for team in ['gat']:
        print('++++++正在获取 ' + match1[team] + ' 信息++++++')
        for i in range((end - begin).days):  # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            yesterday = str(day) + ' 23:59:59'
            last_month = str(day)
            print('正在更新 ' + match1[team] + last_month + ' 号订单信息…………')
            searchType = '订单号'      # 运单号，订单号   查询切换
            m.orderInfo(searchType, team, last_month)
    print('更新耗时：', datetime.datetime.now() - start)
