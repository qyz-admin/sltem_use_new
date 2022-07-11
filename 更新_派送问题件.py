import pandas as pd
import os
import datetime
import time
from tqdm import tqdm

import xlwings
import xlsxwriter
import math
import requests
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
from 更新_已下架_压单_头程导入提单号 import QueryTwoLower
from 查询_订单检索 import QueryOrder

# -*- coding:utf-8 -*-
class QueryTwo(Settings, Settings_sso):
    def __init__(self, userMobile, password, login_TmpCode,handle):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self._online()
        # self.sso_online_Two()
        # self.sso__online_handle(login_TmpCode)
        # # self.sso__online_auto()
        #
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

    # 获取查询时间
    def readInfo(self, team):
        print('>>>>>>正式查询中<<<<<<')
        print('正在获取需要订单信息......')
        start = datetime.datetime.now()
        if team == '派送问题件_跟进表':
            last_time = datetime.datetime.now().strftime('%Y-%m') + '-01'
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')

        else:
            sql = '''SELECT DISTINCT 处理时间 FROM {0} d GROUP BY 处理时间 ORDER BY 处理时间 DESC'''.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['处理时间'][0])
            last_time = (rq + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            now_time = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        print('******************起止时间：' + team + last_time + ' - ' + now_time + ' ******************')
        return last_time, now_time

    def outport_getDeliveryList(self, timeStart, timeEnd):
        rq = datetime.datetime.now().strftime('%m.%d')
        # self.getOrderList(timeStart, timeEnd)
        # self.getDeliveryList(timeStart, timeEnd)

        print('正在获取excel内容…………')
        sql = '''SELECT *, IF(派送问题 LIKE "地址问题" OR 派送问题 LIKE "客户要求更改派送时间或者地址","地址问题/客户要求更改派送时间或者地址",派送问题) AS 问题件类型, 
                                            IF(备注 <> "", IF(备注 LIKE "已签收%","已签收",IF(备注 LIKE "无人接听%","无人接听",IF(备注 LIKE "拒收%","拒收",
                                            IF(备注 LIKE "%*%","未回复",IF(备注 NOT LIKE "%*%","回复",备注))))),备注) AS 回复类型
                 FROM 派送问题件_跟进表 p
                 WHERE p.创建日期 >= '{0}'  
                 ORDER BY 币种, 创建日期 , 
                 FIELD(问题件类型,'送至便利店','地址问题/客户要求更改派送时间或者地址','客户长期不在','送达客户不在','客户不接电话','拒收','合计');'''.format(timeStart)
        df1 = pd.read_sql_query(sql=sql, con=self.engine1)

        sql = '''SELECT 币种, 创建日期, 签收单量, 拒收单量, concat(ROUND(IFNULL(签收单量 / 总单量,0) * 100,2),'%') as 签收率,
                        派送问题件单量, 问题件类型,单量,短信,邮件,在线, 电话,客户回复再派量,
                        concat(ROUND(IFNULL(物流再派签收 / 物流再派,0) * 100,2),'%') as 物流再派签收率,
                        concat(ROUND(IFNULL(物流3派签收 / 物流3派,0) * 100,2),'%') as 物流3派签收率,
                        未派, 异常, 上月签收单量, 上月拒收单量, 
                        concat(ROUND(IFNULL(上月签收单量 / 上月总单量,0) * 100,2),'%') as 上月签收率, 上月派送问题件单量
                FROM (  SELECT s1.币种, s1.创建日期, s3.签收单量, s3.拒收单量, s3.总单量, 派送问题件单量, 问题件类型,
                                COUNT(订单编号) AS 单量, NULL AS 短信, NULL AS 邮件, NULL AS 在线, 
                                SUM(IF(s1.回复类型 <> "" AND 备注 <> "已签收" AND 备注 <> "无人接听",1,0)) AS 电话, 
                                SUM(IF(回复类型 = "回复",1,0)) AS 客户回复再派量, 物流再派, 物流再派签收, 物流3派, 物流3派签收, NULL AS 未派, 异常,
                                s4.签收单量 AS 上月签收单量, s4.拒收单量 AS 上月拒收单量, s4.总单量 AS 上月总单量, s5.上月派送问题件单量
                        FROM(   SELECT *, IF(派送问题 LIKE "地址问题" OR 派送问题 LIKE "客户要求更改派送时间或者地址","地址问题/客户要求更改派送时间或者地址",派送问题) AS 问题件类型,
                                        IF(备注 <> "", IF(备注 LIKE "已签收%","已签收",IF(备注 LIKE "无人接听%","无人接听",IF(备注 LIKE "拒收%","拒收",
                                        IF(备注 LIKE "%*%","未回复",IF(备注 NOT LIKE "%*%","回复",备注))))),备注) AS 回复类型
                                FROM 派送问题件_跟进表 p
                                WHERE p.创建日期 >= '{0}'  
                        ) s1
                        LEFT JOIN 
                        (   SELECT 币种, 创建日期, COUNT(订单编号) AS 派送问题件单量,
                                SUM(IF(派送次数 = 2,1,0)) AS 物流再派,
                                SUM(IF(物流状态 = "已签收" AND 派送次数 = 2,1,0)) AS 物流再派签收,
                                SUM(IF(派送次数 > 2,1,0)) AS 物流3派,
                                SUM(IF(物流状态 = "已签收" AND 派送次数 > 2,1,0)) AS 物流3派签收,
                                NULL AS 未派, 
                                SUM(IF(回复类型 = "回复" AND 物流状态 = "拒收",1,0)) AS 异常
                            FROM ( SELECT *, IF(派送问题 LIKE "地址问题" OR 派送问题 LIKE "客户要求更改派送时间或者地址","地址问题/客户要求更改派送时间或者地址",派送问题) AS 问题件类型, 
                                            IF(备注 <> "", IF(备注 LIKE "已签收%","已签收",IF(备注 LIKE "无人接听%","无人接听",IF(备注 LIKE "拒收%","拒收",
                                            IF(备注 LIKE "%*%","未回复",IF(备注 NOT LIKE "%*%","回复",备注))))),备注) AS 回复类型
                                    FROM 派送问题件_跟进表 p
                                    WHERE p.创建日期 >= '{0}'  
                            ) PP
                            GROUP BY 币种, 创建日期
                        ) s2 on s1.币种 =s2.币种 AND s1.创建日期 =s2.创建日期
                        LEFT JOIN `派送问题件_跟进表2` s3 on s1.币种 = s3.币种 AND s1.创建日期 = s3.日期
                        LEFT JOIN `派送问题件_跟进表2` s4 on s1.币种 = s4.币种 AND s1.创建日期 = DATE_SUB(s4.日期,INTERVAL -1 MONTH)
                        LEFT JOIN (SELECT 币种, 创建日期, COUNT(订单编号) AS 上月派送问题件单量
                                    FROM 派送问题件_跟进表 p
                                    WHERE p.创建日期 >= DATE_SUB('{0}',INTERVAL 1 MONTH)  AND p.创建日期 < '{0}'  
                                    GROUP BY 币种, 创建日期
                        ) s5 on s1.币种 = s5.币种 AND s1.创建日期 = DATE_SUB(s5.创建日期,INTERVAL -1 MONTH)
                        GROUP BY s1.币种, s1.创建日期, s1.问题件类型
                ) s
                ORDER BY s.币种, s.创建日期 , 
                FIELD(s.问题件类型,'送至便利店','地址问题/客户要求更改派送时间或者地址','客户长期不在','送达客户不在','客户不接电话','拒收','合计');'''.format(timeStart)
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        db2 = df[(df['币种'].str.contains('台币'))]
        db3 = df[(df['币种'].str.contains('港币'))]
        print(df)
        print(db2)
        print(db3)
        print('正在写入excel…………')
        file_pathT = 'F:\\神龙签收率\\A订单改派跟进\\{0} 派送问题件跟进情况.xlsx'.format(rq)

        df0 = pd.DataFrame([])
        df0.to_excel(file_pathT, index=False)
        writer = pd.ExcelWriter(file_pathT, engine='openpyxl')  # 初始化写入对象
        book = load_workbook(file_pathT)
        writer.book = book
        db2.drop(['币种'], axis=1).to_excel(excel_writer=writer, sheet_name='台湾', index=False)
        db3.drop(['币种'], axis=1).to_excel(excel_writer=writer, sheet_name='香港', index=False)
        df1.to_excel(excel_writer=writer, sheet_name='明细', index=False)
        if 'Sheet1' in book.sheetnames:  # 删除新建文档时的第一个工作表
            del book['Sheet1']
        writer.save()
        writer.close()
        try:
            print('正在运行 派送问题件表 宏…………')
            app = xlwings.App(visible=False, add_book=False)  # 运行宏调整
            app.display_alerts = False
            wbsht = app.books.open('D:/Users/Administrator/Desktop/slgat_签收计算(ver5.24).xlsm')
            wbsht1 = app.books.open(file_pathT)
            wbsht.macro('派送问题件_修饰')()
            wbsht1.save()
            wbsht1.close()
            wbsht.close()
            app.quit()
            app.quit()
        except Exception as e:
            print('运行失败：', str(Exception) + str(e))
        print('----已写入excel')

    # 查询更新（新后台的获取-派送问题件）
    def getDeliveryList(self, timeStart, timeEnd):  # 进入订单检索界面
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deliveryProblemPackage'}
        data = {'order_number': None, 'waybill_number': None, 'question_level': None, 'question_type': None, 'order_trace_id': None, 'ship_phone': None, 'page': 1, 'pageSize': 90,
                'addtime': None, 'question_time': None, 'trace_time': None, 'create_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'finishtime': None,
                'sale_id': None, 'product_id': None, 'logistics_id': None, 'area_id': None, 'currency_id': None, 'order_status': None, 'logistics_status': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        print('+++已成功发送请求......')
        req = json.loads(req.text)          # json类型数据转换为dict字典
        max_count = req['data']['count']    # 获取 请求订单量

        ordersDict = []
        if max_count != 0 and max_count != []:
            try:
                for result in req['data']['list']:                  # 添加新的字典键-值对，为下面的重新赋值
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
                    data = self._getDeliveryList(timeStart, timeEnd, n)
                    dlist.append(data)
                dp = df.append(dlist, ignore_index=True)
            else:
                dp = df
            dp = dp[['order_number',  'currency', 'addtime', 'create_time', 'finishtime', 'lastQuestionName', 'orderStatus', 'logisticsStatus',
                     'reassignmentTypeName', 'logisticsName',  'questionAddtime', 'userName', 'traceName', 'traceTime', 'content','failNum']]
            dp.columns = ['订单编号', '币种', '下单时间', '创建时间', '完成时间', '派送问题', '订单状态', '物流状态',
                          '订单类型', '物流渠道',  '派送问题首次时间', '处理人', '处理记录', '处理时间', '备注', '派送次数']
            print('正在写入......')
            dp.to_sql('customer_up', con=self.engine1, index=False, if_exists='replace')
            # dp.to_excel('G:\\输出文件\\派送问题件-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
            sql = '''REPLACE INTO 派送问题件_跟进表(订单编号,币种, 下单时间,完成时间,订单状态,物流状态,订单类型,物流渠道, 创建日期, 创建时间, 派送问题, 派送问题首次时间, 派送次数,处理人, 处理记录, 处理时间,备注, 记录时间) 
                    SELECT 订单编号,币种, 下单时间,完成时间,订单状态,物流状态,订单类型,物流渠道, DATE_FORMAT(创建时间,'%Y-%m-%d') 创建日期, 创建时间, 派送问题, 派送问题首次时间, 派送次数, 处理人, 处理记录, IF(处理时间 = '',NULL,处理时间) 处理时间,备注,NOW() 记录时间 
                    FROM customer_up;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            print('写入成功......')
        else:
            print('没有需要获取的信息！！！')
            return
        print('*' * 50)
    def _getDeliveryList(self, timeStart, timeEnd, n):  # 进入派送问题件界面
        print('+++正在查询第 ' + str(n) + ' 页信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.deliveryQuestion&action=getDeliveryList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deliveryProblemPackage'}
        data = {'order_number': None, 'waybill_number': None, 'question_level': None, 'question_type': None, 'order_trace_id': None, 'ship_phone': None, 'page': n, 'pageSize': 90,
                'addtime': None, 'question_time': None, 'trace_time': None, 'create_time': timeStart + ' 00:00:00,' + timeEnd + ' 23:59:59', 'finishtime': None,
                'sale_id': None, 'product_id': None, 'logistics_id': None, 'area_id': None, 'currency_id': None, 'order_status': None, 'logistics_status': None}
        proxy = '47.75.114.218:10020'  # 使用代理服务器
        # proxies = {'http': 'socks5://' + proxy, 'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersDict = []
        try:
            for result in req['data']['list']:  # 添加新的字典键-值对，为下面的重新赋值
                ordersDict.append(result.copy())
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersDict)
        print('++++++第 ' + str(n) + ' 批次查询成功+++++++')
        print('*' * 50)
        return data
    # 查询更新（新后台的获取-订单完成）
    def getOrderList(self, timeStart, timeEnd):  # 进入订单检索界面
        begin = datetime.datetime.strptime(timeStart, '%Y-%m-%d')
        begin = begin.date()
        end = datetime.datetime.strptime(timeEnd, '%Y-%m-%d')
        end = end.date()
        print('正在查询日期---起止时间：' + timeStart + ' - ' + timeEnd)
        currencyId = [13, 6]            # 6 是港币；13 是台币
        logisticsStatus = [2, 3]
        match = {6: '港币', 13: '台币'}
        match2 = {2: '已签收', 3: '拒收'}
        dlist = []
        df =pd.DataFrame([])
        for i in range((end - begin).days):  # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            day_time = str(day)
            for id in currencyId:
                print('+++正在查询： ' + day_time + match[id] + '完成 信息')
                dict = []
                res = {}
                count = self._getOrderList(id, None, day_time, day_time)
                res['币种'] = match[id]
                res['日期'] = day_time
                res['总单量'] = count
                res['签收单量'] = ''
                res['拒收单量'] = ''
                dict.append(res)
                for log in logisticsStatus:
                        print('+++正在查询： ' + match[id] + match2[log] + ' 信息')
                        count2 = self._getOrderList(id, log,  day_time, day_time)
                        if log == 2:
                            res['签收单量'] = count2
                        elif log == 3:
                            res['拒收单量'] = count2
                        # dict.append(res)
                data = pd.json_normalize(dict)
                print(data)
                dlist.append(data)
        dp = df.append(dlist, ignore_index=True)
        dp.to_sql('cache_cp', con=self.engine1, index=False, if_exists='replace')
        sql = '''REPLACE INTO 派送问题件_跟进表2(币种,日期,总单量,签收单量, 拒收单量) 
                SELECT 币种,日期,总单量,签收单量, 拒收单量
                FROM cache_cp;'''
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
    def _getOrderList(self, id, log, timeStart, timeEnd):  # 进入订单检索界面
        print('+++正在查询信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500, 'order_number': None, 'shippingNumber': None, 'orderNumberFuzzy': None, 'shipUsername': None,
                'phone': None, 'email': None, 'ip': None, 'productIds': None, 'saleIds': None, 'payType': None, 'logisticsId': None,
                'logisticsStyle': None, 'logisticsMode': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': id, 'emailStatus': None,
                'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': '', 'warehouse': None, 'isEmptyWayBillNumber': None,
                'logisticsStatus': log, 'orderStatus': None, 'tuan': None, 'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None,
                'volumeEnd': None, 'volumeStart': None, 'chooser_id': None, 'service_id': None, 'autoVerifyStatus': None, 'shipZip': None,
                'remark': None, 'shipState': None, 'weightStart': None, 'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None,
                'order': None, 'sortField': None, 'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None,
                'timeStart': None, 'timeEnd': None, 'finishTimeStart': timeStart + '00:00:00', 'finishTimeEnd': timeEnd + '23:59:59'}
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
        max_count = req['data']['count']  # 获取 请求订单量
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return max_count




if __name__ == '__main__':
    start: datetime = datetime.datetime.now()
    '''
    # -----------------------------------------------自动获取 问题件 状态运行（一）-----------------------------------------
    # 1、 物流问题件；2、物流客诉件；3、物流问题件；4、全部；--->>数据更新切换
    '''

    select = 99
    if int(select) == 99:
        handle = '手0动'
        login_TmpCode = '78b998328b4834f59bc8d4f734cd78f0'
        m = QueryTwo('+86-18538110674', 'qyz35100416', login_TmpCode,handle)
        start: datetime = datetime.datetime.now()

        if int(select) == 1:
            timeStart, timeEnd = m.readInfo('物流问题件')

        elif int(select) == 99:         # 查询更新-派送问题件
            timeStart, timeEnd = m.readInfo('派送问题件_跟进表')
            # m.getOrderList('2022-07-07', '2022-07-12')
            # m.getOrderList(timeStart, timeEnd)                        # 订单完成单量 更新

            # m.getDeliveryList('2022-06-12', '2022-06-30')
            # m.getDeliveryList('2022-07-07', '2022-07-12')
            # m.getDeliveryList(timeStart, timeEnd)                     # 派送问题件 更新



            m.outport_getDeliveryList('2022-07-01', '2022-07-10')
            # m.outport_getDeliveryList(timeStart, timeEnd)



    print('查询耗时：', datetime.datetime.now() - start)