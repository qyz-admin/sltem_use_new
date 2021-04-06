import pandas as pd
import os
import datetime
import xlwings as xl

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
class QueryControl(Settings):
    def __init__(self):
        Settings.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
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

    def writeSqlReplace(self, dataFrame):
        dataFrame.to_sql('tem', con=self.engine1, index=False, if_exists='replace')

    def replaceInto(self, team, dfColumns):
        columns = list(dfColumns)
        columns = ', '.join(columns)
        if team == 'slrb':
            print(team + '---9')
            sql = 'REPLACE INTO {}({}, 添加时间) SELECT *, NOW() 添加时间 FROM tem; '.format(team, columns)
        else:
            print(team)
            sql = 'INSERT IGNORE INTO {}({}, 添加时间) SELECT *, NOW() 添加时间 FROM tem; '.format(team, columns)
        try:
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=100)
        except Exception as e:
            print('插入失败：', str(Exception) + str(e))

    def readSql(self, sql):
        db = pd.read_sql_query(sql=sql, con=self.engine1)
        # db = pd.read_sql(sql=sql, con=self.engine1) or team == 'slgat'
        return db

    # 更新团队产品明细（新后台的第一部分）
    def productIdInfo(self, tokenid, searchType, team):  # 进入订单检索界面，
        print('正在获取需要更新的产品id信息')
        start = datetime.datetime.now()
        month_begin = (datetime.datetime.now() - relativedelta(months=4)).strftime('%Y-%m-%d')
        sql = '''SELECT id,`订单编号`  FROM {0}_order_list sl 
    			WHERE sl.`日期`> '{1}' 
    				AND (sl.`产品名称` IS NULL or sl.`父级分类` IS NULL or  sl.`物流方式` IS NULL)
    				AND ( NOT sl.`系统订单状态` IN ('已删除','问题订单','支付失败','未支付'));'''.format(team, month_begin)
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        if ordersDict.empty:
            print('无需要更新的产品id信息！！！')
            # sys.exit()
            return
        orderId = list(ordersDict['订单编号'])
        print('获取耗时：', datetime.datetime.now() - start)
        max_count = len(orderId)    # 使用len()获取列表的长度，上节学的
        n = 0
        while n < max_count:        # 这里用到了一个while循环，穿越过来的
            pro = ', '.join(orderId[n:n + 10])
            print(pro)
            n = n + 10
            self.productIdquery(tokenid, pro, team)

    def productIdquery(self, tokenid, productid, team):  # 进入订单检索界面，
        start = datetime.datetime.now()
        # productid = '508746'
        # token = '7dd7c0085722cf49493c5ab2ecbc6234'
        url = r'http://gimp.giikin.com/service?service=gorder.customer&action=getProductList&page=1&pageSize=10'\
              r'&productName=&status=&source=&isSensitive=&isGift=&isDistribution=&chooserId=&buyerId='\
              r'&productId=' + str(productid) + '&_token=' + str(tokenid)
        r_header = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36',
            'Referer': 'http://gimp.giikin.com/front/orderToolsServiceQuery'}
        rq = requests.get(url=url, headers=r_header)
        print('已成功发送请求++++++')
        req = rq.json()  # json类型数据
        print('正在转化数据为dataframe…………')
        print(req)
        ordersDict = []
        for result in req['data']['list']:
            print(result)
            # 添加新的字典键-值对，为下面的重新赋值用
        print('正在写入缓存中......')


if __name__ == '__main__':
    m = QueryControl()
    match1 = {'slgat': '港台',
              'sltg': '泰国',
              'slxmt': '新马',
              'slzb': '直播团队',
              'slyn': '越南',
              'slrb': '日本'}
    # messagebox.showinfo("提示！！！", "当前查询已完成--->>> 请前往（ 输出文件 ）查看")
    #  各团队全部订单表-函数
    # m.tgOrderQuan('sltg')

    # team = 'slgat'
    # for tem in ['台湾', '香港']:
    #     m.OrderQuan(team, tem)

    #  订单花费明细查询
    # match9 = {'slgat_zqsb': '港台',
    #           'sltg_zqsb': '泰国',
    #           'slxmt_zqsb': '新马',
    #           'slrb_zqsb_rb': '日本'}
    # team = 'sltg_zqsb'
    # m.sl_tem_cost(team, match9[team])

    team = 'slgat_hfh'  # 第一部分查询
    token = '7dd7c0085722cf49493c5ab2ecbc6234'
    pro = '508746'
    m.productIdquery(token, pro, team)
    # m.productIdInfo(token, '订单号', team)