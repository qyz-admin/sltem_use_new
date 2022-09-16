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
    def __init__(self, userMobile, password, login_TmpCode, handle):
        Settings.__init__(self)
        Settings_sso.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue(maxsize=10)  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self.sso_online_Two()
        # self._online_Two()

        # self.sso__online_auto()
        if handle == '手动':
            self.sso__online_handle(login_TmpCode)
            print(11)
        else:
            # print(11)
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
    def readFormHost(self, team, searchType,pople_Query, timeStart, timeEnd):
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
                elif pople_Query == '电话检索':
                    self.wbsheetHost_iphone(filePath, team, searchType, timeStart, timeEnd)
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
                if sht.api.Visible == -1:
                    try:
                        tem = None
                        db = None
                        db = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                        print(db.columns)
                        # db = db[['订单编号']]
                        columns_value = list(db.columns)                             # 获取数据的标题名，转为列表
                        if searchType == '订单号':
                            tem = '订单编号'
                            if '订单号' in columns_value:
                                db.rename(columns={'订单号': '订单编号'}, inplace=True)
                            for column_val in columns_value:
                                if '订单编号' != column_val:
                                    db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                            db.dropna(axis=0, how='any', inplace=True)                  # 空值（缺失值），将空值所在的行/列删除后
                        elif searchType == '运单号':
                            tem = '运单编号'
                            if '运单号' in columns_value:
                                db.rename(columns={'运单号': '运单编号'}, inplace=True)
                            elif '查件单号' in columns_value:
                                db.rename(columns={'查件单号': '运单编号'}, inplace=True)
                    except Exception as e:
                        print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                    if db is not None and len(db) > 0:
                        # print(db)
                        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
                        print('++++正在获取：' + sht.name + ' 表；共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                        orderId = list(db[tem])
                        max_count = len(orderId)                                    # 使用len()获取列表的长度，上节学的
                        # print(orderId)
                        # print(max_count)
                        if max_count > 500:
                            ord = ','.join(orderId[0:500])
                            # print(ord)
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
                else:
                    print('----不需查询：' + sht.name)
            wb.close()
        app.quit()

    def wbsheetHost_pople(self, filePath, team, searchType):
        fileType = os.path.splitext(filePath)[1]
        app = xlwings.App(visible=False, add_book=False)
        app.display_alerts = False
        if 'xls' in fileType:
            wb = app.books.open(filePath, update_links=False, read_only=True)
            for sht in wb.sheets:
                if sht.api.Visible == -1:
                    db = None
                    tem = None
                    try:
                        db = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                        print(db.columns)
                        # db = db[['订单编号']]
                        if searchType == '订单号':
                            tem = '订单编号'
                            columns_value = list(db.columns)                             # 获取数据的标题名，转为列表
                            if '订单号' in columns_value:
                                db.rename(columns={'订单号': '订单编号'}, inplace=True)
                            for column_val in columns_value:
                                if '订单编号' != column_val:
                                    db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                        elif searchType == '运单号':
                            tem = '运单编号'
                            columns_value = list(db.columns)                            # 获取数据的标题名，转为列表
                            if '运单号' in columns_value:
                                db.rename(columns={'运单号': '运单编号'}, inplace=True)
                            for column_val in columns_value:
                                if '运单编号' != column_val:
                                    db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表

                        db.dropna(axis=0, how='any', inplace=True)                      # 空值（缺失值），将空值所在的行/列删除后
                    except Exception as e:
                        print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                    if db is not None and len(db) > 0:
                        # print(db)
                        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
                        print('++++正在获取：' + sht.name + ' 表；共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                        orderId = list(db[tem])
                        max_count = len(orderId)                                    # 使用len()获取列表的长度，上节学的
                        if max_count > 90:
                            ord = ','.join(orderId[0:90])
                            df = self.orderInfo_pople(ord, searchType)
                            # print(df)
                            dlist = []
                            n = 0
                            while n < max_count-90:                                # 这里用到了一个while循环，穿越过来的
                                n = n + 90
                                ord = ','.join(orderId[n:n + 90])
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
                else:
                    print('----不需查询：' + sht.name)
            wb.close()
        app.quit()

    def wbsheetHost_iphone(self, filePath, team, searchType, timeStart, timeEnd):
        fileType = os.path.splitext(filePath)[1]
        app = xlwings.App(visible=False, add_book=False)
        app.display_alerts = False
        if 'xls' in fileType:
            wb = app.books.open(filePath, update_links=False, read_only=True)
            for sht in wb.sheets:
                if sht.api.Visible == -1:
                    try:
                        db = None
                        db = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                        print(db.columns)
                    except Exception as e:
                        print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                    if db is not None and len(db) > 0:
                        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
                        print('++++正在获取：' + sht.name + ' 表；共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                        orderId = list(db['电话'])
                        # print(orderId)
                        max_count = len(orderId)                                    # 使用len()获取列表的长度，上节学的
                        if max_count > 10:
                            # ord = ','.join(orderId[0:100])
                            ord = ','.join('%s' % d for d in orderId[0:10])
                            # print(ord)
                            df = self.orderInfo_iphone(ord, searchType, timeStart, timeEnd)
                            dlist = []
                            n = 0
                            while n < max_count-10:                                # 这里用到了一个while循环，穿越过来的
                                n = n + 10
                                # ord = ','.join(orderId[n:n + 100])
                                ord = ','.join('%s' % d for d in orderId[n:n + 10])
                                # print(ord)
                                data = self.orderInfo_iphone(ord, searchType, timeStart, timeEnd)
                                dlist.append(data)
                            print('正在写入......')
                            # print(dlist)
                            dp = df.append(dlist, ignore_index=True)
                        else:
                            # print(orderId[0:max_count])
                            # ord = ','.join(orderId[0:max_count])
                            ord = ','.join('%s' %d for d in orderId[0:max_count])
                            # print(ord)
                            dp = self.orderInfo_iphone(ord, searchType, timeStart, timeEnd)
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
                else:
                    print('----不需查询：' + sht.name)
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
                'orderPrefix': None,
                'shippingNumber': None,
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
        print(data)
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
    # 一、订单——电话查询（新后台的获取）
    def orderInfo_iphone(self, ord, searchType, timeStart, timeEnd):  # 进入订单检索界面
        print('+++正在查询订单信息中')
        url = r'https://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/orderToolsOrderSearch'}
        data = {'page': 1, 'pageSize': 500, 'orderPrefix': None,'shippingNumber': None,
                'orderNumberFuzzy': None, 'shipUsername': None, 'phone': ord, 'email': None, 'ip': None, 'productIds': None,
                'saleIds': None, 'payType': None, 'logisticsId': None, 'logisticsStyle': None, 'logisticsMode': None,
                'type': None, 'collId': None, 'isClone': None,
                'currencyId': None, 'emailStatus': None, 'befrom': None, 'areaId': None, 'reassignmentType': None, 'lowerstatus': '',
                'warehouse': None, 'isEmptyWayBillNumber': None, 'logisticsStatus': None, 'orderStatus': None, 'tuan': None,
                'tuanStatus': None, 'hasChangeSale': None, 'optimizer': None, 'volumeEnd': None, 'volumeStart': None, 'chooser_id': None,
                'service_id': None, 'autoVerifyStatus': None, 'shipZip': None, 'remark': None, 'shipState': None, 'weightStart': None,
                'weightEnd': None, 'estimateWeightStart': None, 'estimateWeightEnd': None, 'order': None, 'sortField': None,
                'orderMark': None, 'remarkCheck': None, 'preSecondWaybill': None, 'whid': None,
                'isChangeMark': None, 'percentStart': None, 'percentEnd': None, 'userid': None, 'questionId': None, 'delUserId': None, 'transferNumber': None,
                'smsStatus': None, 'designer_id': None, 'logistics_remarks': None, 'clone_type': None, 'categoryId': None, 'addressType': None,
                'timeStart': timeStart,  'timeEnd': timeEnd
        }
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


    # 二、时间-查询更新（新后台的获取 line运营）
    def order_TimeQuery(self, timeStart, timeEnd, areaId, query):  # 进入订单检索界面
        print('+++正在查询订单信息中起止： ' + timeStart + ':' + timeEnd)
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
        if query == '下单时间':
            data.update({'timeStart': timeStart + '00:00:00',
                         'timeEnd': timeEnd + '23:59:59',
                         'finishTimeStart': None,
                         'finishTimeEnd': None})
        elif query == '完成时间':
            data.update({'timeStart': None,
                         'timeEnd': None,
                         'finishTimeStart': timeStart + '00:00:00',
                         'finishTimeEnd': timeEnd + '23:59:59'})

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
                        'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName',  'update_time']]
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
                data = self._timeQuery(timeStart, timeEnd, n, areaId, query)
                dlist.append(data)
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
            dp.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
            print('查询已导出+++')
        else:
            df.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')  # Xlsx是python用来构造xlsx文件的模块，可以向excel2007+中写text，numbers，formulas 公式以及hyperlinks超链接。
            print('查询已导出+++')

    # 二、时间-查询更新（新后台的获取 全部 & 昨日订单删除分析）
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
                print('查询第 ' + str(n) + ' 页中，剩余次数' + str(in_count - n))
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
                    db0.insert(0, '删除人', '')
                    db0.to_sql('cache', con=self.engine1, index=False, if_exists='replace')
                    # 更新删除订单的原因
                    self.del_people()
                    self.del_order_day()
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
    def _timeQuery(self, timeStart, timeEnd, n, areaId, query):  # 进入订单检索界面
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
        if query == '下单时间':
            data.update({'timeStart': timeStart + '00:00:00',
                         'timeEnd': timeEnd + '23:59:59',
                         'finishTimeStart': None,
                         'finishTimeEnd': None})
        elif query == '完成时间':
            data.update({'timeStart': None,
                         'timeEnd': None,
                         'finishTimeStart': timeStart + '00:00:00',
                         'finishTimeEnd': timeEnd + '23:59:59'})
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
                           'logisticsControl', 'logisticsRefuse', 'logisticsUpdateTime', 'stateTime', 'collDomain', 'typeName','update_time']]
                df.columns = ['订单编号', '平台', '币种', '运营团队', '产品id', '产品名称', '出货单名称', '规格(中文)', '收货人', '联系电话', '拉黑率', '电话长度',
                              '配送地址', '应付金额', '数量', '订单状态', '运单号', '支付方式', '下单时间', '审核人', '审核时间', '物流渠道', '货物类型',
                              '是否低价', '站点ID', '商品ID', '订单类型', '物流状态', '重量', '删除原因', '问题原因', '下单人', '转采购时间', '发货时间', '上线时间',
                              '完成时间', '销售退货时间', '备注', 'IP', '体积', '省洲', '市/区', '选品人', '优化师', '审单类型', '异常提示', '克隆人', '克隆ID', '发货仓库', '是否发送短信',
                              '物流渠道预设方式', '拒收原因', '物流更新时间', '状态时间', '来源域名', '订单来源类型',  '更新时间']
        except Exception as e:
            print('------查询为空')
        print('******本批次查询成功')
        return df

    # 更新删除订单的原因
    def del_people(self):
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
            df = self._del_people(ord, '')
            dlist = []
            n = 0
            while n < max_count - 500:  # 这里用到了一个while循环，穿越过来的
                n = n + 500
                ord = ','.join(orderId[n:n + 500])
                data = self._del_people(ord, '')
                dlist.append(data)
            print('正在写入......')
            dp = df.append(dlist, ignore_index=True)
        else:
            ord = ','.join(orderId[0:max_count])
            dp = self._del_people(ord, '')
        if dp is None or len(dp) == 0:
            print('查询为空，不需更新+++')
        else:
            # print(dp)
            dp.to_sql('cache_cp', con=self.engine1, index=False, if_exists='replace')
            print('正在更新删除原因中......')
            sql = '''update cache a, cache_cp b set a.`删除人`= IF(b.`删除人` = '', NULL, b.`删除人`) where a.`订单编号`=b.`订单编号`;'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            sql = '''update `cache` a
                    SET a.`删除人` = IF( a.`删除人` = '', NULL, a.`删除人` ),
                        a.`订单状态` = IF( a.`订单状态` = '', NULL, a.`订单状态` ),
                        a.`审核人` = IF( a.`审核人` = '', NULL, a.`审核人` ),
                        a.`删除原因` = IF( a.`删除原因` = '', NULL, a.`删除原因` ),
                        a.`问题原因` = IF( a.`问题原因` = '', NULL, a.`问题原因` ),
                        a.`下单人` = IF( a.`下单人` = '', NULL, a.`下单人` ),
                        a.`IP` = IF( a.`IP` = '', NULL, a.`IP` ),
                        a.`删除原因` = IF(a.`删除原因` LIKE ';%', RIGHT(a.`删除原因`,CHAR_LENGTH(a.`删除原因`)-1), a.`删除原因`);'''
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        print('查询耗时：', datetime.datetime.now() - start)
    def _del_people(self, ord, areaId):  # 进入订单检索界面
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
            df.columns = ['订单编号', '删除人']
        except Exception as e:
            print('------查询为空')
        print('******本批次查询成功')
        return df # 更新删除订单的原因 -函数调用

    # 删除订单的  分析导出
    def del_order_day(self):
        print('正在分析 昨日 删单原因中')
        sql ='''SELECT *,concat(ROUND(SUM(IF(删单原因 IS NULL OR 删单原因 = '',总订单量-订单量,订单量)) / SUM(总订单量) * 100,2),'%') as '删单率'
                FROM (SELECT s1.*,总订单量,总删单量, 系统删单量
                      FROM (SELECT 币种,运营团队,删单原因,COUNT(订单编号) AS 订单量
                            FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                  FROM `cache` c
                            ) w
                            GROUP BY 币种,运营团队,删单原因
                      ) s1
                      LEFT JOIN
                      ( SELECT 币种,运营团队,COUNT(订单编号) AS 总订单量, SUM(IF(订单状态 = '已删除',1,0)) AS 总删单量, SUM(IF(订单状态 = '已删除' AND 删除人 IS NULL,1,0)) AS 系统删单量
                        FROM `cache` w
                        GROUP BY 币种,运营团队
                      ) s2 
                      ON s1.`币种`=s2.`币种` AND s1.`运营团队`=s2.`运营团队`
                ) s
                WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台')
                GROUP BY 币种,运营团队,删单原因
                ORDER BY FIELD(币种,'台币','港币','合计'),
                         FIELD(运营团队,'神龙家族-港澳台','火凤凰-港澳台','神龙-运营1组','Line运营','金鹏家族-小虎队','合计'),
                         订单量 DESC;'''
        df1 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df1)
        # 初始化设置
        sl_tem, sl_tem_lh, sl_tem_ey, sl_tem_cf = '', '', '', ''
        hfh_tem, hfh_tem_lh, hfh_tem_ey, hfh_tem_cf = '', '', '', ''
        for row in df1.itertuples():
            tem = getattr(row, '运营团队')
            delreson = getattr(row, '删单原因')
            count = getattr(row, '订单量')
            if tem == '神龙家族-港澳台' and delreson == None:
                sl_tem = '*神  龙:   昨日单量：' + str(int(getattr(row, '总订单量'))) + '；删单量：' + str(int(getattr(row, '总删单量'))) + '；删单率：' + str(getattr(row, '删单率')) + '；系统删单量：' + str(int(getattr(row, '系统删单量'))) + '单;'
                # print(sl_tem)
            elif tem == '神龙家族-港澳台' and '拉黑率订单' == delreson:
                sl_tem_lh = '，\n            其中占比较多的是：拉黑率订单：' + str(int(count)) + '单, '
                # print(sl_tem_lh)
            elif tem == '神龙家族-港澳台' and '恶意订单' == delreson:
                sl_tem_ey = '恶意订单：' + str(int(count)) + '单, '
                # print(sl_tem_ey)
            elif tem == '神龙家族-港澳台' and '重复订单' == delreson:
                sl_tem_cf = '重复订单：' + str(int(count)) + '单;'
                # print(sl_tem_cf)

            elif tem == '火凤凰-港澳台' and delreson == None:
                hfh_tem = '*火凤凰:  昨日单量：' + str(int(getattr(row, '总订单量'))) + '；删单量：' + str(int(getattr(row, '总删单量'))) + '；删单率：' + str(getattr(row, '删单率')) + '；系统删单量：' + str(int(getattr(row, '系统删单量'))) + '单;'
                # print(hfh_tem)
            elif tem == '火凤凰-港澳台' and '拉黑率订单' == delreson:
                hfh_tem_lh = '，\n            其中占比较多的是：拉黑率订单：' + str(int(count)) + '单, '
                # print(hfh_tem_lh)
            elif tem == '火凤凰-港澳台' and '恶意订单' == delreson:
                hfh_tem_ey = '恶意订单：' + str(int(count)) + '单, '
                # print(hfh_tem_ey)
            elif tem == '火凤凰-港澳台' and '重复订单' == delreson:
                hfh_tem_cf = '重复订单：' + str(int(count)) + '单;'
                # print(hfh_tem_cf)
        print('*' * 50)
        print(sl_tem + sl_tem_lh + sl_tem_ey + sl_tem_cf)
        print(hfh_tem + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf)

        print('正在获取 删单明细 拉黑率信息 二…………')
        sql ='''SELECT 币种, 删单原因, IF(联系电话 = '总计',NULL,联系电话) AS 拉黑率订单, 订单量, 拉黑率70以上, 拉黑率70以下
                FROM(
                    (SELECT s1.*
					    FROM (SELECT IFNULL(币种,'总计') 币种,IFNULL(删单原因,'总计') 删单原因,IFNULL(联系电话,'总计') 联系电话, COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                        FROM `cache` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '拉黑率%'
                                ) w
                                GROUP BY 币种,删单原因, 联系电话
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 <> "总计" AND 删单原因 <> "总计"
                        GROUP BY 币种,删单原因, 联系电话
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                    UNION ALL
                    (SELECT s1.*
                        FROM (  SELECT IFNULL(币种,'总计') 币种,IFNULL(删单原因,'总计') 删单原因,IFNULL(ip,'总计') ip, COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                        FROM `cache` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '拉黑率%'
                                ) w
                                GROUP BY 币种,删单原因, ip
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 <> "总计" AND 删单原因 <> "总计"
                        GROUP BY 币种,删单原因, ip
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                ) s;'''
        df2 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df2)
        df2.to_sql('cache_cp', con=self.engine1, index=False, if_exists='replace')
        print('正在记录 拉黑率信息......')
        sql = '''REPLACE INTO {0}(币种, 删单类型, 删单明细, 单量, 记录日期, 更新时间) SELECT 币种, 删单原因, 拉黑率订单, 订单量, CURDATE() 记录日期, NOW() 更新时间 FROM cache_cp s WHERE s.拉黑率订单 IS NOT NULL;'''.format('day_delete_cache')
        pd.read_sql_query(sql=sql, con=self.engine1,chunksize=10000)

        sl_Black ,sl_Black_iphone , sl_Black_ip = '','',''
        k = 0
        k2 = 0
        for row in df2.itertuples():
            tem_Black = getattr(row, '拉黑率订单')
            count = getattr(row, '订单量')
            if tem_Black == None:
                sl_Black = '*拉黑率删除:  ' + str(int(getattr(row, '订单量'))) + '单；拉黑率70以上的：' + str(int(getattr(row, '拉黑率70以上'))) +'单；'
                # print(sl_Black)
            elif tem_Black != None and '.' not in tem_Black:
                if count >= 10:
                    if k == 0:
                        sl_Black_iphone = '\n           同一电话有：(0' + str(int(tem_Black)) + ':' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        sl_Black_iphone =sl_Black_iphone + '(0' + str(int(tem_Black)) + ':' + str(int(count)) + '单);'
                        k = k + 1
                else:
                    if k == 0:
                        sl_Black_iphone = '\n           同一电话有：(0' + str(int(tem_Black)) + ':' + str(int(count)) +'单),'
                        k = k + 1
                    elif k > 0:
                        sl_Black_iphone =sl_Black_iphone + '(0' + str(int(tem_Black)) + ':' + str(int(count)) + '单);'
                        k = k + 1
                # print(sl_Black_iphone)
            elif tem_Black != None and '.' in tem_Black:
                if count >= 10:
                    if k2 == 0:
                        sl_Black_ip = '\n           同一ip有：   (' + str(getattr(row, '拉黑率订单')) + ':' + str(int(getattr(row, '订单量'))) + '单),'
                        k2 = k2 + 1
                    elif k > 0:
                        sl_Black_ip = sl_Black_ip + '(' + str(tem_Black) + ':' + str(int(count)) + '单);'
                        k2 = k2 + 1
                else:
                    if k2 == 0:
                        sl_Black_ip = '\n           同一ip有：   (' + str(tem_Black) + ':' + str(int(count)) +'单),'
                        k2 = k2 + 1
                    elif k > 0:
                        sl_Black_ip = sl_Black_ip + '（' + str(tem_Black) + ':' + str(int(count)) + '单);'
                        k2 = k2 + 1
                # print(sl_Black_ip)
        print('*' * 50)
        print(sl_Black + sl_Black_iphone + sl_Black_ip)

        print('正在获取 删单明细 恶意删除信息 三…………')
        sql ='''SELECT 币种, 删单原因, IF(联系电话 = '总计',NULL,联系电话) AS 恶意删除, 订单量, 拉黑率70以上, 拉黑率70以下
                FROM(
                    (SELECT s1.*
                        FROM (  SELECT IFNULL(币种,'总计') 币种,IFNULL(删单原因,'总计') 删单原因,IFNULL(联系电话,'总计') 联系电话,COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                        FROM `cache` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '恶意%'
                                ) w
                                GROUP BY 币种,删单原因, 联系电话
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 <> "总计" AND 删单原因 <> "总计"
                        GROUP BY 币种,删单原因, 联系电话
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                    UNION ALL
                    (SELECT s1.*
                        FROM (  SELECT IFNULL(币种,'总计') 币种,IFNULL(删单原因,'总计') 删单原因,IFNULL(ip,'总计') ip,COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                        FROM `cache` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '恶意%'
                                ) w
                                GROUP BY 币种,删单原因, ip
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 <> "总计" AND 删单原因 <> "总计"
                        GROUP BY 币种,删单原因, ip
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                ) s;'''
        df3 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df3)
        df3.to_sql('cache_cp', con=self.engine1, index=False, if_exists='replace')
        print('正在记录 恶意删除信息......')
        sql = '''REPLACE INTO {0}(币种, 删单类型, 删单明细, 单量, 记录日期, 更新时间) SELECT 币种, 删单原因, 恶意删除, 订单量, CURDATE() 记录日期, NOW() 更新时间 FROM cache_cp s WHERE s.恶意删除 IS NOT NULL;'''.format('day_delete_cache')
        pd.read_sql_query(sql=sql, con=self.engine1,chunksize=10000)

        st_ey, st_ey_iphone, st_ey_ip = '','',''
        k = 0
        k2 = 0
        for row in df3.itertuples():
            tem_Black = getattr(row, '恶意删除')
            count = getattr(row, '订单量')
            if tem_Black == None:
                st_ey = '*恶意删除： ' + str(int(count)) + '单；拉黑率70以上的：' + str(int(getattr(row, '拉黑率70以上'))) + '单；低于70的：' + str(int(getattr(row, '拉黑率70以下'))) + '单；'
            elif tem_Black != None and '.' not in tem_Black:
                if count >= 10:
                    if k == 0:
                        st_ey_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        st_ey_iphone = st_ey_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1
                else:
                    if k == 0:
                        st_ey_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) +'单),'
                        k = k + 1
                    elif k > 0:
                        st_ey_iphone = st_ey_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1

            elif tem_Black != None and '.' in tem_Black:
                if count >= 10:
                    if k2 == 0:
                        st_ey_ip = ';\n           同一ip有：    (' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_ey_ip = st_ey_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
                else:
                    if k2 == 0:
                        st_ey_ip = ';\n           同一ip有：    (' + str(tem_Black) + ': ' + str(int(count)) +'单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_ey_ip = st_ey_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
        print('*' * 50)
        print(st_ey + st_ey_iphone + st_ey_ip)

        print('正在获取 删单明细 重复删除信息 四…………')
        sql = '''SELECT IFNULL(币种,'币种') 币种,IFNULL(删单原因,'总计') 重复删除,COUNT(订单编号) AS 订单量
                    FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                            FROM `cache` c
                            WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '重复订单%'
                    ) w
                GROUP BY 币种,删单原因;'''
        df4 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df4)
        cf_del = ''
        for row in df4.itertuples():
            tem_Black = getattr(row, '重复删除')
            count = getattr(row, '订单量')
            cf_del = '*重复删除：' + str(count) + '单, 查询后都是客户上笔未收到或者是连续订多笔订单重复删除；'
        print('*' * 50)
        print(cf_del)

        print('正在获取 删单明细 系统删除信息 五…………')
        sql = '''SELECT s1.*
                        FROM ( 
        					    (SELECT *
                                    FROM (SELECT IFNULL(币种,'币种') 币种,IFNULL(删单原因,'总计') 系统删除,COUNT(订单编号) AS 订单量
                                            FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                                    FROM `cache` c
                                                    WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除人 IS NULL
                                            ) w
                                            GROUP BY 币种,删单原因
                                            WITH ROLLUP
                                    ) w1
                                    WHERE 币种 <> '币种'
                                    ORDER BY 订单量 DESC
                                    LIMIT 4
                                )
                                UNION ALL
                                 ( SELECT 币种,联系电话 AS 系统删除,COUNT(订单编号) AS 订单量
                                    FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                            FROM `cache` c
                                            WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除人 IS NULL
                                    ) w
                                    GROUP BY 币种,联系电话
                                    ORDER BY 订单量 DESC
                                    LIMIT 4
                                )
                                UNION ALL
                                (SELECT 币种,ip AS 系统删除,COUNT(订单编号) AS 订单量
                                    FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '重复订单%','重复订单',删除原因))) 删单原因
                                            FROM `cache` c
                                            WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除人 IS NULL
                                    ) w
                                    GROUP BY 币种,ip
                                    ORDER BY 订单量 DESC
                                    LIMIT 4
                                )
                        ) s1;'''
        df5 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df5)
        df5.to_sql('cache_cp', con=self.engine1, index=False, if_exists='replace')
        print('正在记录 系统删除信息......')
        sql = '''REPLACE INTO {0}(币种, 删单类型, 删单明细, 单量, 记录日期, 更新时间) 
                SELECT 币种,'系统删除'  删单类型, 系统删除, 订单量, CURDATE() 记录日期, NOW() 更新时间
				FROM cache_cp s 
				WHERE s.系统删除 LIKE '%.%' OR s.系统删除 LIKE '%9%' OR s.系统删除 LIKE '%8%' OR s.系统删除 LIKE '%7%' OR s.系统删除 LIKE '%6%' OR s.系统删除 LIKE '%5%' 
				   OR s.系统删除 LIKE '%4%' OR s.系统删除 LIKE '%3%' OR s.系统删除 LIKE '%2%' OR s.系统删除 LIKE '%1%' OR s.系统删除 LIKE '%0%';'''.format('day_delete_cache')
        pd.read_sql_query(sql=sql, con=self.engine1,chunksize=10000)

        st_del, st_del_iphone, st_del_ip = '', '', ''
        k = 0
        k2 = 0
        for row in df5.itertuples():
            tem_Black = getattr(row, '系统删除')
            count = getattr(row, '订单量')
            if '总计' in tem_Black:
                st_del = '*系统删除： ' + str(int(count)) + '单；其中比较多的是：'
            elif '订单' in tem_Black:
                st_del = st_del + str(tem_Black) + ':' + str(int(count)) + '单,'

            elif tem_Black != None and '.' not in tem_Black:
                if count >= 10:
                    if k == 0:
                        st_del_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        st_del_iphone = st_del_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1
                else:
                    if k == 0:
                        st_del_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        st_del_iphone = st_del_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1

            elif tem_Black != None and '.' in tem_Black:
                if count >= 10:
                    if k2 == 0:
                        st_del_ip = ';\n           同一ip有：   (' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_del_ip = st_del_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
                else:
                    if k2 == 0:
                        st_del_ip = ';\n           同一ip有：   (' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_del_ip = st_del_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
        print('*' * 50)
        print(st_del + st_del_iphone + st_del_ip)

        print('正在获取 连续3天以上的黑名单 电话、IP信息 六…………')
        sql ='''UPDATE day_delete_cache d 
                    SET d.`删单明细` = IF(d.`删单明细` LIKE '0%', RIGHT(d.`删单明细`,LENGTH(d.`删单明细`)-1),d.`删单明细`);'''
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
        sql = '''SELECT *
                FROM (
                        SELECT 币种,删单类型,删单明细, COUNT(记录日期) AS 次数
                        FROM day_delete_cache d
                        WHERE d.`记录日期` >= IF(DATE_FORMAT('2022-06-29','%w') = 2 ,DATE_SUB(CURDATE(), INTERVAL 5 DAY),IF(DATE_FORMAT('2022-06-29','%w') = 1 ,DATE_SUB(CURDATE(), INTERVAL 6 DAY),DATE_SUB(CURDATE(), INTERVAL 2 DAY)))
                        GROUP BY 币种,删单类型,删单明细
                ) s
                WHERE s.次数 >= 3
                ORDER BY 币种,删单类型,删单明细, 次数 DESC;'''
        df6 = pd.read_sql_query(sql=sql, con=self.engine1)
        db61 = df6[~(df6['删单类型'].str.contains('系统删除'))]
        # day_del = db61.to_markdown()
        day_del = '''注意：连续3天同电话\IP的信息>>>'''
        day_del2 = '恶意订单:'
        day_del3 = '拉黑率订单:'
        for row in db61.itertuples():
            tem = getattr(row, '删单类型')
            info = getattr(row, '删单明细')
            count = getattr(row, '次数')
            # day_del = day_del + '\n' + tem + ':' + info + ':  ' + str(int(count)) + '单,'
            if '恶意订单' in tem:
                day_del2 = day_del2 + '\n' + info + ':  ' + str(int(count)) + '单,'
            elif '拉黑率订单' in tem:
                day_del3 = day_del3 + '\n' + info + ':  ' + str(int(count)) + '单,'

        print('*' * 50)
        print(day_del + '\n' + day_del2 + '\n' + day_del3)


        # url = "https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e"  # url为机器人的webhook
        url = "https://oapi.dingtalk.com/robot/send?access_token=9a92f00296846dcd3ec8b52d7bacce114a9e34cb2d5dbfad9ce3371ab8d037f9"  # url为机器人的webhook  港台客服
        content = r'r"H:\桌面\需要用到的文件\文件夹\out2.jpeg"'  # 钉钉消息内容，注意test是自定义的关键字，需要在钉钉机器人设置中添加，这样才能接收到消息
        mobile_list = ['18538110674']  # 要@的人的手机号，可以是多个，注意：钉钉机器人设置中需要添加这些人，否则不会接收到消息
        isAtAll = '是'  # 是否@所有人
        headers = {'Content-Type': 'application/json', "Charset": "UTF-8"}
        data = {"msgtype": "text",
                # "markdown": {# 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                #         "title": 'TEST',
                #         "text": "#### 昨日删单率分析" + "\n" +
                #         "* " + sl_tem +
                #         "   + " + sl_tem_lh + sl_tem_ey + sl_tem_cf + "\n" +
                #         "* " + hfh_tem +
                #         "   + " + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf
                # },
                # "markdown": {  # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                #     "title": 'TEST',
                #     "text": "#### 昨日删单率分析" + "\n" +
                #             "* " + sl_tem +
                #             "   + " + sl_tem_lh + sl_tem_ey + sl_tem_cf + "\n" +
                #             "* " + hfh_tem +
                #             "   + " + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf
                # },
                "text": {
                    # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                    "content": 'TEST神龙 - 火凤凰 昨日订单删除分析' + '\n' +
                               sl_tem + sl_tem_lh + sl_tem_ey + sl_tem_cf + '\n' +
                               hfh_tem + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf + '\n' +
                               sl_Black + sl_Black_iphone + sl_Black_ip + '\n' +
                               st_ey + st_ey_iphone + st_ey_ip + '\n' +
                               cf_del + '\n' +
                               st_del + st_del_iphone + st_del_ip + '\n' + '\n' +
                               day_del + '\n' + day_del2 + '\n' + day_del3
                    # "content": 'TEST'
                },
                "at": {# 要@的人
                        # "atMobiles": mobile_list,
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


    # 删除订单的  分析导出 测试
    def del_order_day_two(self):
        print('正在分析 昨日 删单原因中')
        sql ='''SELECT *,concat(ROUND(SUM(IF(删单原因 IS NULL OR 删单原因 = '',总订单量-订单量,订单量)) / SUM(总订单量) * 100,2),'%') as '删单率'
                FROM (SELECT s1.*,总订单量,总删单量, 系统删单量
                      FROM (SELECT 币种,运营团队,删单原因,COUNT(订单编号) AS 订单量
                            FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',IF(删除原因 LIKE '拉黑率%','拉黑率订单',删除原因)) 删单原因
                                  FROM `worksheet` c
                            ) w
                            GROUP BY 币种,运营团队,删单原因
                      ) s1
                      LEFT JOIN
                      ( SELECT 币种,运营团队,COUNT(订单编号) AS 总订单量, SUM(IF(订单状态 = '已删除',1,0)) AS 总删单量, SUM(IF(订单状态 = '已删除' AND 删除人 IS NULL,1,0)) AS 系统删单量
                        FROM `worksheet` w
                        GROUP BY 币种,运营团队
                      ) s2 
                      ON s1.`币种`=s2.`币种` AND s1.`运营团队`=s2.`运营团队`
                ) s
                WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台')
                GROUP BY 币种,运营团队,删单原因
                ORDER BY FIELD(币种,'台币','港币','合计'),
                         FIELD(运营团队,'神龙家族-港澳台','火凤凰-港澳台','神龙-运营1组','Line运营','金鹏家族-小虎队','合计'),
                         订单量 DESC;'''
        df1 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df1)
        # 初始化设置
        sl_tem, sl_tem_lh, sl_tem_ey, sl_tem_cf = '', '', '', ''
        hfh_tem, hfh_tem_lh, hfh_tem_ey, hfh_tem_cf = '', '', '', ''
        for row in df1.itertuples():
            tem = getattr(row, '运营团队')
            delreson = getattr(row, '删单原因')
            count = getattr(row, '订单量')
            if tem == '神龙家族-港澳台' and delreson == None:
                sl_tem = '*神  龙:   昨日单量：' + str(int(getattr(row, '总订单量'))) + '；删单量：' + str(int(getattr(row, '总删单量'))) + '；删单率：' + str(getattr(row, '删单率')) + '；系统删单量：' + str(int(getattr(row, '系统删单量'))) + '单;'
                # print(sl_tem)
            elif tem == '神龙家族-港澳台' and '拉黑率订单' in delreson:
                sl_tem_lh = '，\n            其中占比较多的是：拉黑率订单：' + str(int(count)) + '单, '
                # print(sl_tem_lh)
            elif tem == '神龙家族-港澳台' and '恶意订单' in delreson:
                sl_tem_ey = '恶意订单：' + str(int(count)) + '单, '
                # print(sl_tem_ey)
            elif tem == '神龙家族-港澳台' and '重复订单' in delreson:
                sl_tem_cf = '重复订单：' + str(int(count)) + '单;'
                # print(sl_tem_cf)

            elif tem == '火凤凰-港澳台' and delreson == None:
                hfh_tem = '*火凤凰:  昨日单量：' + str(int(getattr(row, '总订单量'))) + '；删单量：' + str(int(getattr(row, '总删单量'))) + '；删单率：' + str(getattr(row, '删单率')) + '；系统删单量：' + str(int(getattr(row, '系统删单量'))) + '单;'
                # print(hfh_tem)
            elif tem == '火凤凰-港澳台' and '拉黑率订单' in delreson:
                hfh_tem_lh = '，\n            其中占比较多的是：拉黑率订单：' + str(int(count)) + '单, '
                # print(hfh_tem_lh)
            elif tem == '火凤凰-港澳台' and '恶意订单' in delreson:
                hfh_tem_ey = '恶意订单：' + str(int(count)) + '单, '
                # print(hfh_tem_ey)
            elif tem == '火凤凰-港澳台' and '重复订单' in delreson:
                hfh_tem_cf = '重复订单：' + str(int(count)) + '单;'
                # print(hfh_tem_cf)
        print('*' * 50)
        print(sl_tem + sl_tem_lh + sl_tem_ey + sl_tem_cf)
        print(hfh_tem + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf)

        print('正在获取 删单明细 拉黑率信息 二…………')
        sql ='''SELECT 币种, 删单原因, 联系电话 AS 拉黑率订单, 订单量, 拉黑率70以上, 拉黑率70以下
                FROM(
                    (SELECT s1.*
                        FROM (  SELECT 币种,删单原因, 联系电话,COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '拉黑率%','拉黑率订单',删除原因) 删单原因
                                        FROM `worksheet` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '拉黑率%'
                                ) w
                                GROUP BY 币种,删单原因, 联系电话
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 IS NOT NULL AND 删单原因 IS NOT NULL AND 删单原因 <> ""
                        GROUP BY 币种,删单原因, 联系电话
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                    UNION ALL
                    (SELECT s1.*
                        FROM (  SELECT 币种,删单原因, ip,COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '拉黑率%','拉黑率订单',删除原因) 删单原因
                                        FROM `worksheet` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '拉黑率%'
                                ) w
                                GROUP BY 币种,删单原因, ip
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 IS NOT NULL AND 删单原因 IS NOT NULL AND 删单原因 <> ""
                        GROUP BY 币种,删单原因, ip
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                ) s;'''
        df2 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df2)
        sl_Black ,sl_Black_iphone , sl_Black_ip = '','',''
        k = 0
        k2 = 0
        for row in df2.itertuples():
            tem_Black = getattr(row, '拉黑率订单')
            count = getattr(row, '订单量')
            if tem_Black == None:
                sl_Black = '*拉黑率删除:  ' + str(int(getattr(row, '订单量'))) + '单；拉黑率70以上的：' + str(int(getattr(row, '拉黑率70以上'))) +'单；'
                # print(sl_Black)
            elif tem_Black != None and '.' not in tem_Black:
                if count >= 10:
                    if k == 0:
                        sl_Black_iphone = '\n           同一电话有：(0' + str(int(tem_Black)) + ':' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        sl_Black_iphone =sl_Black_iphone + '(0' + str(int(tem_Black)) + ':' + str(int(count)) + '单);'
                        k = k + 1
                else:
                    if k == 0:
                        sl_Black_iphone = '\n           同一电话有：(0' + str(int(tem_Black)) + ':' + str(int(count)) +'单),'
                        k = k + 1
                    elif k > 0:
                        sl_Black_iphone =sl_Black_iphone + '(0' + str(int(tem_Black)) + ':' + str(int(count)) + '单);'
                        k = k + 1
                    # print(sl_Black_iphone)

            elif tem_Black != None and '.' in tem_Black:
                if count >= 10:
                    if k2 == 0:
                        sl_Black_ip = '\n           同一ip有：   (' + str(getattr(row, '拉黑率订单')) + ':' + str(int(getattr(row, '订单量'))) + '单),'
                        k2 = k2 + 1
                    elif k > 0:
                        sl_Black_ip = sl_Black_ip + '(' + str(tem_Black) + ':' + str(int(count)) + '单);'
                        k2 = k2 + 1
                else:
                    if k2 == 0:
                        sl_Black_ip = '\n           同一ip有：   (' + str(tem_Black) + ':' + str(int(count)) +'单),'
                        k2 = k2 + 1
                    elif k > 0:
                        sl_Black_ip = sl_Black_ip + '（' + str(tem_Black) + ':' + str(int(count)) + '单);'
                        k2 = k2 + 1
                    # print(sl_Black_ip)
        print('*' * 50)
        print(sl_Black + sl_Black_iphone + sl_Black_ip)


        print('正在获取 删单明细 恶意删除信息 三…………')
        sql ='''SELECT 币种, 删单原因, 联系电话 AS 恶意删除, 订单量, 拉黑率70以上, 拉黑率70以下
                FROM(
                    (SELECT s1.*
                        FROM (  SELECT 币种,删单原因, 联系电话,COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',删除原因) 删单原因
                                        FROM `worksheet` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '恶意%'
                                ) w
                                GROUP BY 币种,删单原因, 联系电话
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 IS NOT NULL AND 删单原因 IS NOT NULL AND 删单原因 <> ""
                        GROUP BY 币种,删单原因, 联系电话
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                    UNION ALL
                    (SELECT s1.*
                        FROM (  SELECT 币种,删单原因, ip,COUNT(订单编号) AS 订单量, SUM(IF(拉黑率 > 70 ,1,0)) AS 拉黑率70以上,SUM(IF(拉黑率 < 70 ,1,0)) AS 拉黑率70以下
                                FROM (SELECT *,IF(删除原因 LIKE '恶意%','恶意订单',删除原因) 删单原因
                                        FROM `worksheet` c
                                        WHERE 币种 = '台币' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '恶意%'
                                ) w
                                GROUP BY 币种,删单原因, ip
                                WITH ROLLUP
                        )  s1
                        WHERE 币种 IS NOT NULL AND 删单原因 IS NOT NULL AND 删单原因 <> ""
                        GROUP BY 币种,删单原因, ip
                        ORDER BY 订单量 desc
                        LIMIT 5
                    ) 
                ) s;'''
        df3 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df3)
        st_ey, st_ey_iphone, st_ey_ip = '','',''
        k = 0
        k2 = 0
        for row in df3.itertuples():
            tem_Black = getattr(row, '恶意删除')
            count = getattr(row, '订单量')
            if tem_Black == None:
                st_ey = '*恶意删除： ' + str(int(count)) + '单；拉黑率70以上的：' + str(int(getattr(row, '拉黑率70以上'))) + '单；低于70的：' + str(int(getattr(row, '拉黑率70以下'))) + '单；'
            elif tem_Black != None and '.' not in tem_Black:
                if count >= 10:
                    if k == 0:
                        st_ey_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        st_ey_iphone = st_ey_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1
                else:
                    if k == 0:
                        st_ey_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) +'单),'
                        k = k + 1
                    elif k > 0:
                        st_ey_iphone = st_ey_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1

            elif tem_Black != None and '.' in tem_Black:
                if count >= 10:
                    if k2 == 0:
                        st_ey_ip = ';\n           同一ip有：    (' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_ey_ip = st_ey_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
                else:
                    if k2 == 0:
                        st_ey_ip = ';\n           同一ip有：    (' + str(tem_Black) + ': ' + str(int(count)) +'单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_ey_ip = st_ey_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
        print('*' * 50)
        print(st_ey + st_ey_iphone + st_ey_ip)

        print('正在获取 删单明细 重复删除信息 四…………')
        sql = '''SELECT IFNULL(币种,'币种') 币种,IFNULL(删单原因,'总计') 重复删除,COUNT(订单编号) AS 订单量
                    FROM (SELECT *,IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '恶意%','恶意订单',删除原因)) 删单原因
                            FROM `worksheet` c
                            WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除原因 LIKE '重复订单%'
                    ) w
                GROUP BY 币种,删单原因;'''
        df4 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df4)
        cf_del = ''
        for row in df4.itertuples():
            tem_Black = getattr(row, '重复删除')
            count = getattr(row, '订单量')
            cf_del = '*重复删除：' + str(count) + '单, 查询后都是客户上笔未收到或者是连续订多笔订单重复删除；'
        print('*' * 50)
        print(cf_del)


        print('正在获取 删单明细 系统删除信息 物五…………')
        sql = '''SELECT s1.*
                        FROM ( 
        					    (SELECT *
                                    FROM (SELECT IFNULL(币种,'币种') 币种,IFNULL(删单原因,'总计') 系统删除,COUNT(订单编号) AS 订单量
                                            FROM (SELECT *,IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '恶意%','恶意订单',删除原因)) 删单原因
                                                    FROM `worksheet` c
                                                    WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除人 IS NULL
                                            ) w
                                            GROUP BY 币种,删单原因
                                            WITH ROLLUP
                                    ) w1
                                    WHERE 币种 <> '币种'
                                    ORDER BY 订单量 DESC
                                    LIMIT 4
                                )
                                UNION ALL
                                 ( SELECT 币种,联系电话 AS 系统删除,COUNT(订单编号) AS 订单量
                                    FROM (SELECT *,IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '恶意%','恶意订单',删除原因)) 删单原因
                                            FROM `worksheet` c
                                            WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除人 IS NULL
                                    ) w
                                    GROUP BY 币种,联系电话
                                    ORDER BY 订单量 DESC
                                    LIMIT 4
                                )
                                UNION ALL
                                (SELECT 币种,ip AS 系统删除,COUNT(订单编号) AS 订单量
                                    FROM (SELECT *,IF(删除原因 LIKE '拉黑率%','拉黑率订单',IF(删除原因 LIKE '恶意%','恶意订单',删除原因)) 删单原因
                                            FROM `worksheet` c
                                            WHERE 币种 = '台币' AND 订单状态 = '已删除' AND 运营团队 IN ('神龙家族-港澳台','火凤凰-港澳台') AND 删除人 IS NULL
                                    ) w
                                    GROUP BY 币种,ip
                                    ORDER BY 订单量 DESC
                                    LIMIT 4
                                )
                        ) s1;'''
        df5 = pd.read_sql_query(sql=sql, con=self.engine1)
        # print(df5)
        st_del, st_del_iphone, st_del_ip = '', '', ''
        k = 0
        k2 = 0
        for row in df5.itertuples():
            tem_Black = getattr(row, '系统删除')
            count = getattr(row, '订单量')
            if '总计' in tem_Black:
                st_del = '*系统删除： ' + str(int(count)) + '单；其中比较多的是：'
            elif '订单' in tem_Black:
                st_del = st_del + str(tem_Black) + ':' + str(int(count)) + '单,'

            elif tem_Black != None and '.' not in tem_Black:
                if count >= 10:
                    if k == 0:
                        st_del_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        st_del_iphone = st_del_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1
                else:
                    if k == 0:
                        st_del_iphone = ';\n           同一电话有：(' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k = k + 1
                    elif k > 0:
                        st_del_iphone = st_del_iphone + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k = k + 1

            elif tem_Black != None and '.' in tem_Black:
                if count >= 10:
                    if k2 == 0:
                        st_del_ip = ';\n           同一ip有：   (' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_del_ip = st_del_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
                else:
                    if k2 == 0:
                        st_del_ip = ';\n           同一ip有：   (' + str(tem_Black) + ': ' + str(int(count)) + '单),'
                        k2 = k2 + 1
                    elif k2 > 0:
                        st_del_ip = st_del_ip + '(0' + str(tem_Black) + ': ' + str(int(count)) + '单);'
                        k2 = k2 + 1
        print('*' * 50)
        print(st_del + st_del_iphone + st_del_ip)


        url = "https://oapi.dingtalk.com/robot/send?access_token=68eeb5baf4625d0748b15431800b185fec8056a3dbac2755457f3905b0c8ea1e"  # url为机器人的webhook
        content = r'r"H:\桌面\需要用到的文件\文件夹\out2.jpeg"'  # 钉钉消息内容，注意test是自定义的关键字，需要在钉钉机器人设置中添加，这样才能接收到消息
        mobile_list = ['18538110674']  # 要@的人的手机号，可以是多个，注意：钉钉机器人设置中需要添加这些人，否则不会接收到消息
        isAtAll = '是'  # 是否@所有人
        headers = {'Content-Type': 'application/json', "Charset": "UTF-8"}
        data = {"msgtype": "markdown",
                # "markdown": {# 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                #         "title": 'TEST',
                #         "text": "#### 昨日删单率分析" + "\n" +
                #         "* " + sl_tem +
                #         "   + " + sl_tem_lh + sl_tem_ey + sl_tem_cf + "\n" +
                #         "* " + hfh_tem +
                #         "   + " + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf
                # },
                "markdown": {  # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                    "title": 'TEST',
                    "text": "#### 昨日删单率分析" + "\n" +
                            "| 一个普通标题 | 一个普通标题 | 一个普通标题 |一个普通标题 | 一个普通标题 | 一个普通标题 |一个普通标题 | 一个普通标题 |"
                            # "* " + sl_tem + sl_tem_lh + sl_tem_ey + sl_tem_cf + '\n' +
                            # "   + " + hfh_tem + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf + '\n' +
                            # "* " + sl_Black + sl_Black_iphone + sl_Black_ip + '\n' +
                            # "   + " + st_ey + st_ey_iphone + st_ey_ip + '\n' +
                            # "* " + cf_del + '\n' +
                            # "   + " + st_del + st_del_iphone + st_del_ip
                },
                # "text": {
                #     # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                #     "content": 'TEST' + '\n' +
                #                sl_tem + sl_tem_lh + sl_tem_ey + sl_tem_cf + '\n' +
                #                hfh_tem + hfh_tem_lh + hfh_tem_ey + hfh_tem_cf + '\n' +
                #                sl_Black + sl_Black_iphone + sl_Black_ip + '\n' +
                #                st_ey + st_ey_iphone + st_ey_ip + '\n' +
                #                cf_del + '\n' +
                #                st_del + st_del_iphone + st_del_ip
                # },
                "at": {# 要@的人
                        # "atMobiles": mobile_list,
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

        # print('正在获取 删单原因汇总 信息…………')
        # sql ='''SELECT s1.*
        #       FROM (
        #             SELECT 币种,删除原因,COUNT(订单编号) AS 订单量,
        #             		SUM(IF(拉黑率 > 80 ,1,0)) AS 拉黑率80以上,
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
    m = QueryOrder('+86-18538110674', 'qyz04163510.', '202900d829aa38ccbab4fcd676fe1cae', '手0动')
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

    select = 33                                 # 1、 正在按订单查询；2、正在按时间查询；--->>数据更新切换
    if int(select) == 1:
        print("1-->>> 正在按订单查询+++")
        team = 'gat'
        searchType = '订单号'
        pople_Query = '订单检索'                # 客服查询；订单检索
        m.readFormHost(team, searchType, pople_Query, 'timeStart', 'timeEnd')        # 导入；，更新--->>数据更新切换

    elif int(select) == 2:
        print("1-->>> 正在按运单号查询+++")
        team = 'gat'
        searchType = '运单号'
        pople_Query = '客服查询'  # 客服查询；订单检索
        m.readFormHost(team, searchType, pople_Query, 'timeStart', 'timeEnd')  # 导入；，更新--->>数据更新切换

    elif int(select) == 3:
        print("2-->>> 正在按下单时间查询+++")
        timeStart = '2022-03-01'
        timeEnd = '2022-03-01'
        areaId = ''
        query = '下单时间'
        m.order_TimeQuery(timeStart, timeEnd, areaId, query)

    elif int(select) == 33:
        print("2-->>> 正在按完成时间查询+++")
        timeStart = '2022-09-15'
        timeEnd = '2022-09-15'
        areaId = ''
        query = '完成时间'
        m.order_TimeQuery(timeStart, timeEnd, areaId, query)

    if int(select) == 4:
        print("1-->>> 正在按电话查询+++")
        team = 'gat'
        searchType = '电话'
        pople_Query = '电话检索'                # 电话查询；订单检索
        timeStart = '2022-08-01 00:00:00'
        timeEnd = '2022-08-31 23:59:59'
        m.readFormHost(team, searchType, pople_Query, timeStart, timeEnd)

    elif int(select) == 9:
        m.del_order_day()


    print('查询耗时：', datetime.datetime.now() - start)