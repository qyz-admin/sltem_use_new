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
class QueryOrder(Settings, Settings_sso):
    def __init__(self, userMobile, password, login_TmpCode):
        Settings.__init__(self)
        Settings_sso.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue(maxsize=10)  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self.sso_online_Two()
        # self._online_Two()

        self.sso__online_auto()

        # self.sso__online_handle(login_TmpCode)
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
    def readFormHost(self, team, searchType,pople_Query):
        start = datetime.datetime.now()
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                if pople_Query == '客服查询':
                    self.wbsheetHost_pople(filePath, team, searchType)
                else:
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
                    # db = db[['订单编号']]
                    columns_value = list(db.columns)                             # 获取数据的标题名，转为列表
                    if '订单号' in columns_value:
                        db.rename(columns={'订单号': '订单编号'}, inplace=True)
                        # db = db[['订单号']]
                    # if '订单编号' in column_val:
                    #     db = db[['订单编号']]
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
                        # print(df)
                        dlist = []
                        n = 0
                        while n < max_count-500:                                # 这里用到了一个while循环，穿越过来的
                            n = n + 500
                            ord = ','.join(orderId[n:n + 500])
                            data = self.orderInfoQuery(ord, searchType)
                            dlist.append(data)
                        print('正在写入......')
                        # print(dlist)
                        dp = df.append(dlist, ignore_index=True)
                    else:
                        ord = ','.join(orderId[0:max_count])
                        dp = self.orderInfoQuery(ord, searchType)
                    dp.columns = ['订单编号', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                                  '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                                  '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                                  '完成时间', '销售退货时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                                  '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间', '异常提示', '异常拉黑率',
                                  '拉黑率总量','拉黑率签收','拉黑率拒收','留言']
                    dp.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')   # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
                    print('查询已导出+++')
                else:
                    print('----------数据为空,查询失败：' + sht.name)
            wb.close()
        app.quit()
    def wbsheetHost_pople(self, filePath, team, searchType):
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
                    # db = db[['订单编号']]
                    columns_value = list(db.columns)                             # 获取数据的标题名，转为列表
                    if '订单号' in columns_value:
                        db.rename(columns={'订单号': '订单编号'}, inplace=True)
                        # db = db[['订单号']]
                    # if '订单编号' in column_val:
                    #     db = db[['订单编号']]
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
                    if max_count > 10:
                        ord = ', '.join(orderId[0:10])
                        df = self.orderInfo_pople(ord, searchType)
                        # print(df)
                        dlist = []
                        n = 0
                        while n < max_count-10:                                # 这里用到了一个while循环，穿越过来的
                            n = n + 10
                            ord = ','.join(orderId[n:n + 10])
                            data = self.orderInfo_pople(ord, searchType)
                            dlist.append(data)
                        print('正在写入......')
                        # print(dlist)
                        dp = df.append(dlist, ignore_index=True)
                    else:
                        ord = ','.join(orderId[0:max_count])
                        dp = self.orderInfo_pople(ord, searchType)
                    # dp.columns = ['订单编号', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                    #               '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                    #               '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                    #               '完成时间', '销售退货时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                    #               '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间', '异常提示', '异常拉黑率',
                    #               '拉黑率总量','拉黑率签收','拉黑率拒收','留言']
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
                    columns_value = list(db.columns)  # 获取数据的标题名，转为列表
                    if '订单号' in columns_value:
                        db.rename(columns={'订单号': '订单编号'}, inplace=True)
                    for column_val in columns_value:
                        if '订单编号' != column_val:
                            db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                except Exception as e:
                    print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                if db is not None and len(db) > 0:
                    print('++++正在查询：' + sht.name + ' 共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                    self.cs_InfoQuery(db, searchType)
                else:
                    print('----------数据为空导入失败：' + sht.name)
            wb.close()
        app.quit()
    def cs_InfoQuery(self, db, searchType):  # 调用多线程
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        orderId = list(db['订单编号'])
        max_count = len(orderId)  # 使用len()获取列表的长度，上节学的
        if max_count > 500:
            ord = ', '.join(orderId[0:500])
            # df = self.cs_orderInfoQuery(ord, searchType)
            self.cs_orderInfoQuery(ord, searchType)
            df = self.q.get()
            print(df)
            print(99)
            # self.q.join()
            print('主线程开始执行……………………')
            threads = []  # 多线程用线程池--
            n = 0
            count = 1
            while n < max_count - 500:  # 这里用到了一个while循环，穿越过来的
                n = n + 500
                ord = ','.join(orderId[n:n + 500])
                print(str(count) + '次')
                # threads.append(Thread(target=self.cs_orderInfoQuery, args=(ord, searchType)))  # -----也即是子线程
                t = Thread(target=self.cs_orderInfoQuery, args=(ord, searchType))  # -----也即是子线程
                threads.append(t)
                t.start()
                count = count + 1
            for th in threads:
                th.join()
            print('主线程执行结束---------')
            dlist = []
            for j in range(self.q.qsize()):
                dlist.append(self.q.get())
            print(dlist)
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
            print(dp)
        else:
            ord = ','.join(orderId[0:max_count])
            df = self.cs_orderInfoQuery(ord, searchType)
        dp.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')
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
        # print('+++已成功发送请求......')
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
            # print(df)
            self.q.put(df)
        except Exception as e:
            print('------查询为空')
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        # return df

    # 一、订单——查询更新（新后台的获取）
    def orderInfoQuery(self, ord, searchType):  # 进入订单检索界面
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
        print('+++已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        # print(req)
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

                result['auto_VerifyTip'] = ''
                result['auto_VerifyTip_zl'] = ''
                result['auto_VerifyTip_qs'] = ''
                result['auto_VerifyTip_js'] = ''
                if result['autoVerifyTip'] == "":
                    result['auto_VerifyTip'] = '0.00%'
                else:
                    if '未读到拉黑表记录' in result['autoVerifyTip']:
                        result['auto_VerifyTip'] = '0.00%'
                    else:
                        t3 = result['autoVerifyTip']
                        result['auto_VerifyTip_zl'] = (t3.split('订单配送总量：')[1]).split(',')[0]
                        result['auto_VerifyTip_qs'] = (t3.split('送达订单量：')[1]).split(',')[0]
                        result['auto_VerifyTip_js'] = (t3.split('拒收订单量：')[1]).split(',')[0]
                        if '拉黑率问题' not in result['autoVerifyTip']:
                            t2 = result['autoVerifyTip'].split(',拉黑率')[1]
                            result['auto_VerifyTip'] = t2.split('%;')[0] + '%'
                        else:
                            t2 = result['autoVerifyTip'].split('拒收订单量：')[1]
                            t2 = t2.split('%;')[0]
                            result['auto_VerifyTip'] = t2.split('拉黑率')[1] + '%'
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
                    'finishTime', 'refundTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer',
                    'autoVerify', 'cloneUser', 'isClone', 'warehouse', 'smsStatus', 'logisticsControl','logisticsRefuse', 'logisticsUpdateTime',
                    'stateTime', 'collDomain', 'typeName', 'update_time', 'autoVerifyTip','auto_VerifyTip','auto_VerifyTip_zl','auto_VerifyTip_qs','auto_VerifyTip_js','notes']]
        except Exception as e:
            print('------查询为空')
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return df

    # 一、订单——客服查询（新后台的获取）
    def orderInfo_pople(self, ord, searchType):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getQueryOrder'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 10,
                'orderNumberFuzzy': None,
                'shipUsername': None,
                'phone': None,
                'email': None,
                'ip': None
                }
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
        # print(req)
        ordersdict = []
        # print('正在处理json数据转化为dataframe…………')
        try:
            for result in req['data']['list']:
                result['saleId'] = 0        # 添加新的字典键-值对，为下面的重新赋值用
                result['saleName'] = 0
                result['productId'] = 0
                result['saleProduct'] = 0
                result['spec'] = 0
                result['saleId'] = result['specs'][0]['saleId']
                result['saleName'] = result['specs'][0]['saleName']
                result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
                result['saleProduct'] = (result['specs'][0]['saleProduct']).split('#')[2]
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
        # df = None
        # print(df)
        # try:
        #     df = data[['orderNumber', 'currency', 'area', 'productId', 'saleProduct', 'saleName', 'spec',
        #             'shipInfo.shipName', 'shipInfo.shipPhone', 'percent', 'phoneLength', 'shipInfo.shipAddress',
        #             'amount', 'quantity', 'orderStatus', 'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime',
        #             'logisticsName', 'dpeStyle', 'hasLowPrice', 'collId', 'saleId', 'reassignmentTypeName',
        #             'logisticsStatus', 'weight', 'delReason', 'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime',
        #             'finishTime', 'refundTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'optimizer',
        #             'autoVerify', 'cloneUser', 'isClone', 'warehouse', 'smsStatus', 'logisticsControl','logisticsRefuse', 'logisticsUpdateTime',
        #             'stateTime', 'collDomain', 'typeName', 'update_time','notes']]
        # except Exception as e:
        #     print('------查询为空')
        print('++++++本批次查询成功+++++++')
        print('*' * 50)
        return data


    # 二、时间-查询更新（新后台的获取 line运营）
    def order_TimeQuery(self, timeStart, timeEnd, areaId):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        if areaId == 179:
            data = {'page': 1, 'pageSize': 500, 'order_number': None, 'shippingNumber': None,
                    'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None,
                    'productIds': None,
                    'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None,
                    'logisticsMode': None,
                    'type': None, 'collId': None, 'isClone': None,
                    'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': areaId, 'reassignmentType': None,
                    'lowerstatus': '',
                    'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None,
                    'tuan': None,
                    'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None,
                    'volumeStart': None, 'chooser_id': None,
                    'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None,
                    'weightStart': None,
                    'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None,
                    'sortField': None,
                    'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None,
                    'timeStart': timeStart + '00:00:00', 'timeEnd': timeEnd + '23:59:59'}
        else:
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
                # print(result)
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
                    result['chooser'] = ''
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
            if areaId == 179:
                df = data[['orderNumber', 'befrom', 'currency', 'area', 'productId', 'saleProduct', 'shipInfo.shipPhone', 'percent', 'shipZip', 'amount', 'quantity', 'orderStatus',
                           'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime', 'logisticsName', 'dpeStyle', 'reassignmentTypeName', 'logisticsStatus', 'weight', 'delReason',
                           'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime', 'finishTime', 'hasLowPrice', 'collId', 'saleId', 'refundTime',
                           'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer','autoVerify', 'cloneUser', 'isClone', 'warehouse',
                           'logisticsControl','logisticsRefuse', 'logisticsUpdateTime', 'stateTime','collDomain', 'typeName', 'update_time']]
                df.columns = ['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '联系电话', '拉黑率', '邮编', '应付金额', '数量', '订单状态',
                              '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型', '订单类型', '物流状态', '重量', '删除原因',
                              '问题原因', '下单人', '转采购时间', '发货时间', '上线时间', '完成时间', '是否低价', '站点ID', '商品ID', '销售退货时间',
                              '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型','克隆人',  '克隆ID', '发货仓库',
                              '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
            else:
                df = data[['orderNumber', 'befrom', 'currency', 'area', 'productId', 'saleProduct', 'saleName', 'spec',
                        'shipInfo.shipName', 'shipInfo.shipPhone', 'percent', 'phoneLength', 'shipInfo.shipAddress',
                        'amount', 'quantity', 'orderStatus', 'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime',
                        'logisticsName', 'dpeStyle', 'hasLowPrice', 'collId', 'saleId', 'reassignmentTypeName',
                        'logisticsStatus', 'weight', 'delReason', 'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime',
                        'finishTime', 'refundTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer',
                        'autoVerify', 'autoVerifyTip', 'cloneUser', 'isClone', 'warehouse', 'smsStatus', 'logisticsControl',
                        'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName', 'update_time']]
                df.columns = ['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                              '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                              '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                              '完成时间', '销售退货时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '异常提示', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                              '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
        except Exception as e:
            print('------查询为空')
        print('******首次查询成功******')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        max_count = req['data']['count']
        print(max_count)
        if max_count > 500:
            in_count = math.ceil(max_count/500)
            dlist = []
            n = 1
            while n < in_count:  # 这里用到了一个while循环，穿越过来的
                print('剩余查询次数' + str(in_count - n))
                n = n + 1
                data = self._timeQuery(timeStart, timeEnd, n, areaId)
                dlist.append(data)
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
            dp.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
            print('查询已导出+++')
        else:
            df.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
            print('查询已导出+++')

    # 二、时间-查询更新（新后台的获取 全部）
    def order_TimeQueryT(self, timeStart, timeEnd, areaId, select):  # 进入订单检索界面
        print('+++正在查询 ' + timeStart + ' 到 ' + timeEnd + ' 号订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        if areaId == 179:
            data = {'page': 1, 'pageSize': 500, 'order_number': None, 'shippingNumber': None,
                    'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None,
                    'productIds': None,
                    'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None,
                    'logisticsMode': None,
                    'type': None, 'collId': None, 'isClone': None,
                    'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': areaId, 'reassignmentType': None,
                    'lowerstatus': '',
                    'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None,
                    'tuan': None,
                    'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None,
                    'volumeStart': None, 'chooser_id': None,
                    'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None,
                    'weightStart': None,
                    'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None,
                    'sortField': None,
                    'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None,
                    'timeStart': timeStart + '00:00:00', 'timeEnd': timeEnd + '23:59:59'}
        else:
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
        df = None
        print('******首次查询成功******')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        max_count = req['data']['count']
        print(max_count)
        if max_count > 0:
            in_count = math.ceil(max_count/500)
            df = pd.DataFrame([])
            dlist = []
            n = 1
            while n <= in_count:  # 这里用到了一个while循环，穿越过来的
                print('查询第 ' + str(n) + ' 页中，剩余查询次数' + str(in_count - n))
                data = self._timeQuery(timeStart, timeEnd, n, areaId)
                dlist.append(data)
                n = n + 1
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
            dp.to_excel('G:\\输出文件\\订单检索-{0}{1}.xlsx'.format(99, rq), sheet_name='查询', index=False, engine='xlsxwriter')
            if select != '':
                if select.split('|')[0] == '检查头程直发渠道':
                    db1 = dp[(dp['币种'].str.contains('台币'))]
                    db2 = db1[(db1['订单类型'].str.contains('直发下架|未下架未改派'))]
                    db3 = db2[(db2['订单状态'].str.contains('已转采购|截单中(面单已打印,等待仓库审核)|待审核|待发货|问题订单|问题订单审核|截单'))]
                    db4 = db3[(db3['物流渠道'].str.contains('台湾-铱熙无敌-新竹特货|台湾-铱熙无敌-新竹普货|台湾-立邦普货头程-易速配尾程| '))]
                    # print(db)
                    wb_name = '检查头程直发渠道'
                    db4.to_excel('G:\\输出文件\\订单检索-{0}{1}.xlsx'.format(wb_name, rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
                if select.split('|')[1] == '删单原因':
                    # db1 = dp[(dp['订单状态'].str.contains('已删除'))]
                    db0 = dp[(dp['运营团队'].str.contains('神龙家族-港澳台|火凤凰-港澳台|火凤凰-港台(繁体)'))]
                    wb_name = '删单原因'

                    print('正在导入临时表中......')
                    db0 = db0[['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '收货人', '联系电话', '拉黑率','配送地址', '应付金额', '数量',
                               '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型', '订单类型', '物流状态',  '重量',
                               '删除原因', '问题原因', '下单人', '备注', 'IP', '体积', '审单类型', '异常提示', '克隆人', '克隆ID', '发货仓库',
                               '拒收原因', '物流更新时间', '状态时间', '更新时间']]
                    db0.insert(0, '操作人', '')
                    db0.to_sql('cache', con=self.engine1, index=False, if_exists='replace')
                    # 更新删除订单的原因
                    self.del_reson()
                    sql = '''SELECT * FROM {0};'''.format('cache')
                    db00 = pd.read_sql_query(sql=sql, con=self.engine1)
                    db00.to_excel('G:\\输出文件\\神龙-火凤凰 昨日删单明细{1}.xlsx'.format(wb_name, rq), sheet_name='查询', index=False, engine='xlsxwriter')
            else:
                wb_name ='时间查询'
                dp.to_excel('G:\\输出文件\\订单检索-{0}{1}.xlsx'.format(wb_name, rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
            print('查询已导出+++')
        else:
            print('无信息+++')
    # 订单检索 时间查询的公用函数
    def _timeQuery(self, timeStart, timeEnd, n, areaId):  # 进入订单检索界面
        # print('......正在查询信息中......')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        if areaId == 179:
            data = {'page': n, 'pageSize': 500, 'order_number': None, 'shippingNumber': None,
                    'orderNumberFuzzy': None, 'shipUsername': None, 'phone': None, 'email': None, 'ip': None,
                    'productIds': None,
                    'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None,
                    'logisticsMode': None,
                    'type': None, 'collId': None, 'isClone': None,
                    'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': areaId, 'reassignmentType': None,
                    'lowerstatus': '',
                    'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None,
                    'tuan': None,
                    'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None,
                    'volumeStart': None, 'chooser_id': None,
                    'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None,
                    'weightStart': None,
                    'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None,
                    'sortField': None,
                    'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None,
                    'timeStart': timeStart + '00:00:00', 'timeEnd': timeEnd + '23:59:59'}
        else:
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
        # print('......已成功发送请求......')
        req = json.loads(req.text)  # json类型数据转换为dict字典
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
            if areaId == 179:
                df = data[['orderNumber', 'befrom', 'currency', 'area', 'productId', 'saleProduct', 'shipInfo.shipPhone', 'percent', 'shipZip', 'amount', 'quantity', 'orderStatus',
                           'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime', 'logisticsName', 'dpeStyle', 'reassignmentTypeName', 'logisticsStatus', 'weight', 'delReason',
                           'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime', 'finishTime', 'hasLowPrice', 'collId', 'saleId', 'refundTime',
                           'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer','autoVerify', 'cloneUser', 'isClone', 'warehouse',
                           'logisticsControl','logisticsRefuse', 'logisticsUpdateTime', 'stateTime','collDomain', 'typeName', 'update_time']]
                df.columns = ['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '联系电话', '拉黑率', '邮编', '应付金额', '数量', '订单状态',
                              '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型', '订单类型', '物流状态', '重量', '删除原因',
                              '问题原因', '下单人', '转采购时间', '发货时间', '上线时间', '完成时间', '是否低价', '站点ID', '商品ID', '销售退货时间',
                              '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型','克隆人',  '克隆ID', '发货仓库',
                              '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
            else:
                df = data[['orderNumber', 'befrom', 'currency', 'area', 'productId', 'saleProduct', 'saleName', 'spec', 'shipInfo.shipName', 'shipInfo.shipPhone', 'percent', 'phoneLength',
                           'shipInfo.shipAddress', 'amount', 'quantity', 'orderStatus', 'wayBillNumber', 'payType', 'addTime', 'username', 'verifyTime','logisticsName', 'dpeStyle',
                           'hasLowPrice', 'collId', 'saleId', 'reassignmentTypeName', 'logisticsStatus', 'weight', 'delReason', 'questionReason', 'service', 'transferTime', 'deliveryTime', 'onlineTime',
                           'finishTime', 'refundTime', 'remark', 'ip', 'volume', 'shipInfo.shipState', 'shipInfo.shipCity', 'chooser', 'optimizer', 'autoVerify', 'autoVerifyTip', 'cloneUser', 'isClone', 'warehouse', 'smsStatus',
                           'logisticsControl', 'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName', 'update_time']]
                df.columns = ['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                              '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                              '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                              '完成时间', '销售退货时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '异常提示', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                              '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型', '更新时间']
        except Exception as e:
            print('------查询为空')
        print('******本批次查询成功')
        return df

    # 更新删除订单的原因
    def del_reson(self):
        print('正在更新 订单删除 信息……………………………………………………………………………………………………………………………………………………………………………………')
        start = datetime.datetime.now()
        sql = '''SELECT 订单编号 FROM {0} s WHERE s.`订单状态` = '已删除';'''.format('cache')
        db = pd.read_sql_query(sql=sql, con=self.engine1)
        if db.empty:
            print('无需要更新订单信息！！！')
            return
        print(db['订单编号'][0])
        orderId = list(db['订单编号'])
        max_count = len(orderId)  # 使用len()获取列表的长度，上节学的
        if max_count > 500:
            ord = ', '.join(orderId[0:500])
            df = self._del_reson(ord, '')
            dlist = []
            n = 0
            while n < max_count - 500:  # 这里用到了一个while循环，穿越过来的
                n = n + 500
                ord = ','.join(orderId[n:n + 500])
                data = self._del_reson(ord, '')
                dlist.append(data)
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
        else:
            ord = ','.join(orderId[0:max_count])
            dp = self._del_reson(ord, '')
        if dp is None or len(dp) == 0:
            print('查询为空，不需更新+++')
        else:
            # print(dp)
            dp.to_sql('cache_cp', con=self.engine1, index=False, if_exists='replace')
            print('正在更新删除原因中......')
            sql = '''update cache a, cache_cp b set a.`操作人`= IF(b.`操作人` = '' or a.`操作人` = '', NULL, b.`操作人`) where a.`订单编号`=b.`订单编号`;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('查询耗时：', datetime.datetime.now() - start)
    # 更新删除订单的原因 -函数调用
    def _del_reson(self, ord, areaId):  # 进入订单检索界面
        # print('......正在查询信息中......')
        url = r'https://gimp.giikin.com/service?service=gorder.order&action=getRemoveOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:89.0) Gecko/20100101 Firefox/89.0',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/deletedOrder'}
        if areaId == 179:
            data = {'page': 1, 'pageSize': 500, 'orderPrefix': ord, 'shipUsername': None, 'shippingNumber': None, 'email': None, 'saleIds': None, 'ip': None,
                    'productIds': None, 'phone': None, 'optimizer': None, 'payment': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                    'emailStatus': None, 'befrom': None, 'areaId': areaId, 'orderStatus': None, 'timeStart': None, 'timeEnd': None, 'payType': None,
                    'questionId': None, 'reassignmentType': None, 'delUserId': None, 'delReasonIds': None,'delTimeStart':None, 'delTimeEnd': None}
        else:
            data = {'page': 1, 'pageSize': 500, 'orderPrefix': ord, 'shipUsername': None, 'shippingNumber': None, 'email': None, 'saleIds': None, 'ip': None,
                    'productIds': None, 'phone': None, 'optimizer': None, 'payment': None, 'type': None, 'collId': None, 'isClone': None, 'currencyId': None,
                    'emailStatus': None, 'befrom': None, 'areaId': None, 'orderStatus': None, 'timeStart': None, 'timeEnd': None, 'payType': None,
                    'questionId': None, 'reassignmentType': None, 'delUserId': None, 'delReasonIds': None,'delTimeStart':None, 'delTimeEnd': None}
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
                ordersdict.append(result)
        except Exception as e:
            print('转化失败： 重新获取中', str(Exception) + str(e))
        data = pd.json_normalize(ordersdict)
        df = None
        try:
            # df = data[['orderNumber', 'currency', 'area', 'orderStatus', 'addTime', 'username', 'verifyTime',
            #            'dpeStyle', 'reassignmentTypeName', 'logisticsStatus', 'delReason', 'questionReason',
            #            'transferTime', 'deliveryTime', 'hasLowPrice', 'remark', 'ip', 'autoVerify', 'warehouse']]
            # df.columns = ['订单编号', '币种', '运营团队', '订单状态', '下单时间', '操作人', '审核时间',
            #               '货物类型', '订单类型', '物流状态', '删除原因', '问题原因',
            #               '转采购时间', '发货时间', '是否低价', '备注', 'IP', '审单类型', '发货仓库']
            df = data[['orderNumber', 'username']]
            df.columns = ['订单编号', '操作人']
        except Exception as e:
            print('------查询为空')
        print('******本批次查询成功')
        return df


    # 删除订单的  分析导出
    def del_order(self):
        print('+++正在分析 昨日 删单原因中')
        listT = []  # 查询sql的结果 存放池
        print('正在获取 删单明细…………')
        # sql ='''SELECT * FROM `cache` c  WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台');'''
        # df0 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df0)

        print('正在获取 删单明细 信息…………')
        sql ='''SELECT *,concat(ROUND(SUM(IF(删单原因 IS NULL OR 删单原因 = '',总订单量-订单量,订单量)) / SUM(总订单量) * 100,2),'%') as '删单率'
                FROM (
                      SELECT s1.*,总订单量,总删单量
                      FROM (
                            SELECT 币种,运营团队,删单原因,COUNT(订单编号) AS 订单量
                            FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
                                  FROM `worksheet` c
                            ) w
                            GROUP BY 币种,运营团队,删单原因
                      ) s1
                      LEFT JOIN
                      (
                            SELECT 币种,运营团队,COUNT(订单编号) AS 总订单量,
                                  SUM(IF(订单状态 = '已删除',1,0)) AS 总删单量
                            FROM `worksheet` w
                            GROUP BY 币种,运营团队
                      ) s2 ON s1.`币种`=s2.`币种` AND s1.`运营团队`=s2.`运营团队`
                ) s
                WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台')
                GROUP BY 币种,运营团队,删单原因
                ORDER BY FIELD(币种,'台币','港币','合计'),
                         FIELD(运营团队,'神龙家族-港澳台','火凤凰-港澳台','神龙-运营1组','Line运营','金鹏家族-小虎队','合计'),
                         订单量 DESC;'''
#         sql ='''SELECT 币种,运营团队,删单原因2 AS '删单原因(单)',订单量2 AS '删单量(单)',删单率
# FROM (
# SELECT 币种,运营团队,
#       删单原因,
# --       IF(删单原因 IS NULL ,CONCAT('总订单量：',总订单量,'单; 总删单量：',总删单量,'单;'),删单原因) AS 删单原因2,
#       IF(删单原因 IS NULL ,CONCAT('总订单量：',总订单量),删单原因) AS 删单原因2,
#       IF(删单原因 IS NULL ,CONCAT('总删单量：',总删单量),订单量) AS 订单量2,
#       订单量,总订单量,总删单量,
#       concat(ROUND(SUM(IF(删单原因 IS NULL OR 删单原因 = '',总订单量-订单量,订单量)) / SUM(总订单量) * 100,2),'%') as '删单率'
#                 FROM (
#                       SELECT s1.*,总订单量,总删单量
#                       FROM (
#                             SELECT 币种,运营团队,删单原因,COUNT(订单编号) AS 订单量
#                             FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
#                                   FROM `worksheet` c
#                             ) w
#                             GROUP BY 币种,运营团队,删单原因
#                       ) s1
#                       LEFT JOIN
#                       (
#                             SELECT 币种,运营团队,COUNT(订单编号) AS 总订单量,
#                                   SUM(IF(订单状态 = '已删除',1,0)) AS 总删单量
#                             FROM `worksheet` w
#                             GROUP BY 币种,运营团队
#                       ) s2 ON s1.`币种`=s2.`币种` AND s1.`运营团队`=s2.`运营团队`
#                 ) s
#                 WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台')
#                 GROUP BY 币种,运营团队,删单原因
# --                 WITH rollup
#                 ORDER BY FIELD(币种,'台币','港币','合计'),
#                          FIELD(运营团队,'神龙家族-港澳台','火凤凰-港澳台','神龙-运营1组','Line运营','金鹏家族-小虎队','合计'),
#                          订单量 DESC
# ) ss;'''
        df1 = pd.read_sql_query(sql=sql, con=self.engine1)
        print(df1)
        # da = df1['总订单量'].values[0]
        da = df1.总订单量.values[0]
        print(da)

        for row in df1.itertuples():
            tem = getattr(row, '运营团队')
            delreson = getattr(row, '删单原因')
            if tem == '神龙家族-港澳台' and delreson == None:
                tem_count = getattr(row, '总订单量')
                tem_count2 = getattr(row, '总删单量')
                tem_count3 = getattr(row, '删单率')
                print(tem_count)
                print(tem_count2)
                print(tem_count3)


        # ax = df1.plot()
        # fig = ax.get_figure()
        # fig.savefig(r"H:\桌面\需要用到的文件\文件夹\out2.jpeg")

        # plt.rcParams['font.sans-serif'] = ['Arial Unicode MS']  # 显示中文字体
        # fig = plt.figure(figsize=(3, 4), dpi=1400)  # dpi表示清晰度
        # ax = fig.add_subplot(111, frame_on=False)
        # ax.xaxis.set_visible(False)  # hide the x axis
        # ax.yaxis.set_visible(False)  # hide the y axis
        # table(ax, df1, loc='center')  # 将df换成需要保存的dataframe即可
        # plt.savefig(r"H:\桌面\需要用到的文件\文件夹\out.jpeg")

        # im = Image.fromarray(df1)
        # im.save(r"H:\桌面\需要用到的文件\文件夹\out.jpeg")

        # listT.append(df1)
        #
        # print('正在获取 删单原因汇总 信息…………')
        # sql ='''SELECT s1.*
        #       FROM (
        #             SELECT 币种,删除原因,COUNT(订单编号) AS 订单量,
        #             		SUM(IF(拉黑率 > 80 ,1,0)) AS 拉黑率80以上,
		# 					SUM(IF(拉黑率 < 80 ,1,0)) AS 拉黑率80以下
        #             FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #                                     FROM `cache` c
        #                                     WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台')
        #                         ) w
        #             GROUP BY 币种,删单原因
        #       )  s1
        #             WHERE 删除原因 IS NOT NULL AND 删除原因 <> ""
        #             GROUP BY 币种,删除原因
        #       ORDER BY 订单量 desc
        #       LIMIT 5;'''
        # df2 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df2)
        #
        # print('正在获取 删单原因明细（恶意订单-电话） 信息…………')
        # sql ='''SELECT 币种,删单原因,联系电话,COUNT(订单编号) AS 订单量
        #         FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #               FROM `cache` c
        #               WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '%恶意%'
        #             ) w
        #         GROUP BY 币种,`联系电话`
		# 		ORDER BY 订单量 desc;'''
        # df30 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df30)
        # print('正在获取 删单原因明细（恶意订单-ip） 信息…………')
        # sql = '''SELECT 币种,删单原因,IP,COUNT(订单编号) AS 订单量
        #                 FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #                       FROM `cache` c
        #                       WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '%恶意%'
        #                     ) w
        #                 GROUP BY 币种,`IP`
        # 				ORDER BY 订单量 desc;'''
        # df31 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df31)
        #
        # print('正在获取 删单原因明细（拉黑率-电话） 信息…………')
        # sql = '''SELECT 币种,删单原因,联系电话,COUNT(订单编号) AS 订单量
        #                 FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #                       FROM `cache` c
        #                       WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '%拉黑率%'
        #                     ) w
        #                 GROUP BY 币种,`联系电话`
        # 				ORDER BY 订单量 desc;'''
        # df40 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df40)
        # print('正在获取 删单原因明细（拉黑率-ip） 信息…………')
        # sql = '''SELECT 币种,删单原因,IP,COUNT(订单编号) AS 订单量
        #                         FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #                               FROM `cache` c
        #                               WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '%拉黑率%'
        #                             ) w
        #                         GROUP BY 币种,`IP`
        #         				ORDER BY 订单量 desc;'''
        # df41 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df41)
        #
        # print('正在获取 删单原因明细（系统删除-删单原因） 信息…………')
        # sql = '''SELECT 币种,删单原因,COUNT(订单编号) AS 订单量
        #         FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #              FROM `cache` c
        #              WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND `删除人` IS NULL
        #         ) w
        #         GROUP BY 币种,`删单原因`
		# 		ORDER BY 订单量 desc;'''
        # df50 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df50)
        # print('正在获取 删单原因明细（系统删除-ip） 信息…………')
        # sql = '''SELECT 币种,删单原因,IP,COUNT(订单编号) AS 订单量
        #         FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #              FROM `cache` c
        #              WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND `删除人` IS NULL
        #         ) w
        #         GROUP BY 币种,`IP`
        #         ORDER BY 订单量 desc;'''
        # df51 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df51)
        # print('正在获取 删单原因明细（系统删除-电话） 信息…………')
        # sql = '''SELECT 币种,删单原因,联系电话,COUNT(订单编号) AS 订单量
        #         FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #              FROM `cache` c
        #              WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND `删除人` IS NULL
        #         ) w
        #         GROUP BY 币种,`联系电话`
        #         ORDER BY 订单量 desc;'''
        # df52 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df52)
        #
        # print('正在获取 删单原因明细（重复订单-电话） 信息…………')
        # sql = '''SELECT 币种,删单原因,联系电话,COUNT(订单编号) AS 订单量
        #                 FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #                       FROM `cache` c
        #                       WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '%恶意%'
        #                     ) w
        #                 GROUP BY 币种,`联系电话`
        # 				ORDER BY 订单量 desc;'''
        # df60 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df60)
        # print('正在获取 删单原因明细（重复订单-ip） 信息…………')
        # sql = '''SELECT 币种,删单原因,IP,COUNT(订单编号) AS 订单量
        #                         FROM (SELECT *,IF(删除原因 LIKE '%恶意%',';恶意订单',IF(删除原因 LIKE '%拉黑率%',';拉黑率订单',删除原因)) 删单原因
        #                               FROM `cache` c
        #                               WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '%恶意%'
        #                             ) w
        #                         GROUP BY 币种,`IP`
        #         				ORDER BY 订单量 desc;'''
        # df61 = pd.read_sql_query(sql=sql, con=self.engine1)
        # listT.append(df61)
        #
        # print('正在写入excel…………')
        # today = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        # file_path = 'H:\\桌面\\需要用到的文件\\输出文件\\工作数量 {}.xlsx'.format(today)
        #
        # writer2 = pd.ExcelWriter(file_path, engine='openpyxl')
        # df1.to_excel(writer2, index=False)                  # 删单
        # df2.to_excel(writer2, index=False, startrow=20)     # 删单原因汇总
        #
        # df30.to_excel(writer2, index=False, startcol=9)     # 恶意订单-电话
        # df31.to_excel(writer2, index=False, startcol=14)    # 恶意订单-ip
        #
        # df40.to_excel(writer2, index=False, startcol=19)     # 拉黑率-电话
        # df41.to_excel(writer2, index=False, startcol=24)     # 拉黑率-ip
        #
        # df50.to_excel(writer2, index=False, startcol=29)     # 系统删除-删单原因
        # df51.to_excel(writer2, index=False, startcol=34)     # 系统删除-ip
        # df52.to_excel(writer2, index=False, startcol=39)     # 系统删除-电话
        #
        # df60.to_excel(writer2, index=False, startcol=44)     # 重复订单-电话）
        # df61.to_excel(writer2, index=False, startcol=49)     # 重复订单-ip
        #
        # writer2.save()
        # writer2.close()
        # print()


if __name__ == '__main__':
    # select = input("请输入需要查询的选项：1=> 按订单查询； 2=> 按时间查询；\n")
    m = QueryOrder('+86-18538110674', 'qyz35100416','5e35cd9579fe31a89eac01de6eacceec')
    # m = QueryOrder('+86-15565053520', 'sunan1022wang.@&')
    start: datetime = datetime.datetime.now()
    match1 = {'gat': '港台', 'gat_order_list': '港台', 'slsc': '品牌'}
    # -----------------------------------------------查询状态运行（一）-----------------------------------------
    # 1、按订单查询状态
    # team = 'gat'
    # searchType = '订单号'运单号
    # m.readFormHost(team, searchType)        # 导入；，更新--->>数据更新切换
    # 2、按时间查询状态
    # m.order_TimeQuery('2021-11-01', '2021-11-09')auto_VerifyTip
    # m.del_reson()

    select = 3                                 # 1、 正在按订单查询；2、正在按时间查询；--->>数据更新切换
    if int(select) == 1:
        print("1-->>> 正在按订单查询+++")
        team = 'gat'
        searchType = '订单号'
        pople_Query = '订单检索'                # 客服查询；订单检索
        m.readFormHost(team, searchType,pople_Query)        # 导入；，更新--->>数据更新切换
    elif int(select) == 2:
        print("2-->>> 正在按时间查询+++")
        timeStart = '2022-03-01'
        timeEnd = '2022-03-01'
        areaId = ''
        m.order_TimeQuery(timeStart, timeEnd, areaId)

    elif int(select) == 3:
        m.del_order()

    print('查询耗时：', datetime.datetime.now() - start)