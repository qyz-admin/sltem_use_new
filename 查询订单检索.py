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
import win32api,win32con
import math
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
        Settings_sso.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
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
    def readFormHost(self, team, searchType):
        start = datetime.datetime.now()
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath, team, searchType)
                # self.cs_wbsheetHost(filePath, team, searchType)
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath, team, searchType):
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
                    if '订单号' in columns_value:
                        db.rename(columns={'订单号': '订单编号'}, inplace=True)
                    for column_val in columns_value:
                        if '订单编号' != column_val:
                            db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                    db.dropna(axis=0, how='any', inplace=True)                  # 空值（缺失值），将空值所在的行/列删除后
                except Exception as e:
                    print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                if db is not None and len(db) > 0:
                    # print(db)
                    rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
                    print('++++正在获取：' + sht.name + ' 表；共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                    orderId = list(db['订单编号'])
                    max_count = len(orderId)                                    # 使用len()获取列表的长度，上节学的
                    if max_count > 500:
                        ord = ', '.join(orderId[0:500])
                        df = self.orderInfoQuery(ord, searchType)
                        dlist = []
                        n = 0
                        while n < max_count-500:                                # 这里用到了一个while循环，穿越过来的
                            n = n + 500
                            ord = ','.join(orderId[n:n + 500])
                            data = self.orderInfoQuery(ord, searchType)
                            dlist.append(data)
                        print('正在写入......')
                        dp = df.append(dlist, ignore_index=True)
                    else:
                        ord = ','.join(orderId[0:max_count])
                        dp = self.orderInfoQuery(ord, searchType)
                    dp.columns = ['订单编号', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                                  '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                                  '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                                  '完成时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                                  '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
                    dp.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
                    print('查询已导出+++')
                else:
                    print('----------数据为空,查询失败：' + sht.name)
            wb.close()
        app.quit()

    # 测试 self.InfoQuery(db, searchType)
    def cs_wbsheetHost(self, filePath, team, searchType):
        match2 = {'gat': '港台'}
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
                    print(db.columns)
                    columns_value = list(db.columns)  # 获取数据的标题名，转为列表
                    if '订单号' in columns_value:
                        db.rename(columns={'订单号': '订单编号'}, inplace=True)
                    for column_val in columns_value:
                        if '订单编号' != column_val:
                            db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                except Exception as e:
                    print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                if db is not None and len(db) > 0:
                    print('++++正在导入更新：' + sht.name + ' 共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                    self.cs_InfoQuery(db, searchType)
                else:
                    print('----------数据为空导入失败：' + sht.name)
            wb.close()
        app.quit()
    def cs_InfoQuery(self, db, searchType):  # 调用多线程
        orderId = list(db['订单编号'])
        max_count = len(orderId)  # 使用len()获取列表的长度，上节学的
        n = 0
        print('主线程开始执行……………………')
        threads = []  # 多线程用线程池--
        while n < max_count:  # 这里用到了一个while循环，穿越过来的
            ord = ', '.join(orderId[n:n + 500])
            n = n + 500
            threads.append(Thread(target=self.orderInfoQuery, args=(ord, searchType)))  # -----也即是子线程
            print('子线程分配完成++++++')
        if threads:  # 当所有的线程都分配完成之后，通过调用每个线程的start()方法再让他们开始。
            print(len(threads))
            for th in threads:
                th.start()  # print ("开启子线程…………")
            for th in threads:
                th.join()  # print ("退出子线程")
            print('主线程执行结束---------')
            dlist = []
            print(self.q.qsize())
            for i in range(len(threads)):  # print(i)
                if not self.q.empty():
                    print(self.q.get())
                    dlist.append(self.q.get(block=False))
                    print(11)
                else:
                    print('取出失败---：', str(Exception))
            print('-----执行结束---------')
            df1 = dlist[0]
            df2 = dlist[1:]
            dp = df1.append(df2, ignore_index=True)
            print(dp)
        else:
            print("没有需要运行子线程！！！")

        # print('主线程执行结束---------')
        # dlist = []
        # print(9)
        # for i in range(len(threads)):  # print(i)
        #     try:
        #         print(self.q.get())
        #         dlist.append(self.q.get())
        #     except Exception as e:
        #         print('取出失败---：', str(Exception) + str(e))
        # print('-----执行结束---------')
        # print('查询耗时：', datetime.datetime.now() - start)
        # pf = pd.DataFrame(list(dlist))  # 将字典列表转换为DataFrame
        # print(pf)
        # print(1)
        # rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # dp = dlist[0].append(dlist[1:], ignore_index=True)
        # dp.to_excel('H:\\桌面\\test\\订单查询{}.xlsx'.format(rq), sheet_name='查询', index=False)
    def cs_orderInfoQuery(self, ord, searchType):  # 进入订单检索界面
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
        try:
            for result in req['data']['list']:
                # 添加新的字典键-值对，为下面的重新赋值用
                result['saleId'] = 0
                result['saleProduct'] = 0
                result['productId'] = 0
                result['spec'] = 0
                result['saleId'] = result['specs'][0]['saleId']
                result['saleProduct'] = (result['specs'][0]['saleProduct']).split('#')[2]
                result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
                result['spec'] = result['specs'][0]['spec']
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
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
            self.orderInfoQuery(ord, searchType)
        data = pd.json_normalize(ordersDict)
        try:
            self.q.put(data)
        except Exception as e:
            print('放入失败---：', str(Exception) + str(e))
        # return data


    # 一、订单——查询更新（新后台的获取）
    def orderInfoQuery(self, ord, searchType):  # 进入订单检索界面
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
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersdict = []
        # print('正在处理json数据转化为dataframe…………')
        try:
            for result in req['data']['list']:
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
            # time.sleep(10)
            # print(team)
            # print(searchType)
            # self.readFormHost(team, searchType)
            # self.orderInfoQuery(ord, searchType)
        #     self.q.put(result)
        # for i in range(len(req['data']['list'])):
        #     ordersdict.append(self.q.get())
        data = pd.json_normalize(ordersdict)
        df = None
        try:
            df = data[['orderNumber', 'currency', 'area', 'productId', 'saleProduct', 'saleName', 'spec',
                    'shipInfo.shipName', 'shipInfo.shipPhone', 'percent', 'phoneLength', 'shipInfo.shipAddress',
                    'amount', 'quantity', 'orderStatus', 'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime',
                    'logisticsName', 'dpeStyle', 'hasLowPrice', 'collId', 'saleId', 'reassignmentTypeName',
                    'logisticsStatus', 'weight', 'delReason', 'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime',
                    'finishTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer',
                    'autoVerify', 'cloneUser', 'isClone', 'warehouse', 'smsStatus', 'logisticsControl',
                    'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName', 'update_time']]
        except Exception as e:
            print('------查询为空')
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return df


    # 二、时间-查询更新（新后台的获取 ）
    def order_TimeQuery(self, timeStart, timeEnd):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500, 'order_number': None, 'shippingNumber': None,
                'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None,
                'type': None, 'collId': None, 'isClone': None,
                'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': '',
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
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersdict = []
        try:
            for result in req['data']['list']:
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
            df = data[['orderNumber', 'currency', 'area', 'productId', 'saleProduct', 'saleName', 'spec',
                    'shipInfo.shipName', 'shipInfo.shipPhone', 'percent', 'phoneLength', 'shipInfo.shipAddress',
                    'amount', 'quantity', 'orderStatus', 'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime',
                    'logisticsName', 'dpeStyle', 'hasLowPrice', 'collId', 'saleId', 'reassignmentTypeName',
                    'logisticsStatus', 'weight', 'delReason', 'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime',
                    'finishTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer',
                    'autoVerify', 'cloneUser', 'isClone', 'warehouse', 'smsStatus', 'logisticsControl',
                    'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName', 'update_time']]
            df.columns = ['订单编号', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                          '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                          '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                          '完成时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                          '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
        except Exception as e:
            print('------查询为空')
        print('******首次查询成功******')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        max_count = req['data']['count']
        print(max_count)
        if max_count > 500:
            in_count = math.ceil(max_count/500)
            print()
            dlist = [in_count]
            n = 1
            while n < in_count:  # 这里用到了一个while循环，穿越过来的
                n = n + 1
                data = self._timeQuery(timeStart, timeEnd, n)
                dlist.append(data)
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
            dp.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
            print('查询已导出+++')
        else:
            df.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
            print('查询已导出+++')

    def _timeQuery(self, timeStart, timeEnd, n):  # 进入订单检索界面
        print('......正在查询信息中......')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': n, 'pageSize': 500, 'order_number': None, 'shippingNumber': None,
                'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None,
                'type': None, 'collId': None, 'isClone': None,
                'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': '',
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
        print('......已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        ordersdict = []
        try:
            for result in req['data']['list']:
                # print(result)
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
            df = data[['orderNumber', 'currency', 'area', 'productId', 'saleProduct', 'saleName', 'spec',
                    'shipInfo.shipName', 'shipInfo.shipPhone', 'percent', 'phoneLength', 'shipInfo.shipAddress',
                    'amount', 'quantity', 'orderStatus', 'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime',
                    'logisticsName', 'dpeStyle', 'hasLowPrice', 'collId', 'saleId', 'reassignmentTypeName',
                    'logisticsStatus', 'weight', 'delReason', 'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime',
                    'finishTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer',
                    'autoVerify', 'cloneUser', 'isClone', 'warehouse', 'smsStatus', 'logisticsControl',
                    'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName', 'update_time']]
            df.columns = ['订单编号', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                          '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                          '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                          '完成时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                          '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
        except Exception as e:
            print('------查询为空')
        print('******本批次查询成功')
        return df

if __name__ == '__main__':
    # select = input("请输入需要查询的选项：1=> 按订单查询； 2=> 按时间查询；\n")
    m = QueryTwo('+86-18538110674', 'qyz04163510')
    start: datetime = datetime.datetime.now()
    match1 = {'gat': '港台', 'gat_order_list': '港台', 'slsc': '品牌'}
    # -----------------------------------------------查询状态运行（一）-----------------------------------------
    # 1、按订单查询状态
    # team = 'gat'
    # searchType = '订单号'
    # m.readFormHost(team, searchType)        # 导入；，更新--->>数据更新切换
    # 2、按时间查询状态
    # m.order_TimeQuery('2021-11-01', '2021-11-09')

    select = 1                                  # 1、 正在按订单查询；2、正在按时间查询；--->>数据更新切换
    if int(select) == 1:
        print("1-->>> 正在按订单查询+++")
        team = 'gat'
        searchType = '订单号'
        m.readFormHost(team, searchType)        # 导入；，更新--->>数据更新切换
    elif int(select) == 2:
        print("2-->>> 正在按时间查询+++")
        timeStart = '2021-09-01'
        timeEnd = '2021-10-31'
        m.order_TimeQuery(timeStart, timeEnd)

    print('查询耗时：', datetime.datetime.now() - start)