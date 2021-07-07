import pandas as pd
import os
import datetime
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

from mysqlControl import MysqlControl
# -*- coding:utf-8 -*-
class QueryUpdate(Settings):
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
        self.m = MysqlControl()
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

    # 获取签收表内容---港澳台更新签收总表
    def readFormHost(self):
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath)
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath):
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
                    for column_val in columns_value:
                        if '订单编号' != column_val:
                            db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                except Exception as e:
                    print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                if db is not None and len(db) > 0:
                    print('++++正在导入查询：' + sht.name + ' 共：' + str(len(db)) + '行',
                          'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                    # 将返回的dateFrame导入数据库的临时表
                    self.writeCacheHost(db)
                    print('++++正在获取：' + sht.name + '--->>>到查询缓存表')
                else:
                    print('----------数据为空导入失败：' + sht.name)
            wb.close()
        app.quit()

    # 写入更新缓存表
    def writeCacheHost(self, dataFrame):
        dataFrame.to_sql('sheet1_iphone', con=self.engine1, index=False, if_exists='replace')
        print('正在获取查询数据内容…………')
        sql = '''SELECT si.`订单编号`,a.系统订单状态, a.系统物流状态,IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态,
			            IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态
                FROM sheet1_iphone si
                LEFT JOIN gat_order_list a 	ON si.`订单编号` = a.`订单编号`
                LEFT JOIN ( SELECT * 
						    FROM gat 
					 	    WHERE id IN (SELECT MAX(id) FROM gat WHERE gat.添加时间 > '2021-02-01 00:00:00' GROUP BY 运单编号) 
						    ORDER BY id
					    ) b ON si.`订单编号` = b.`订单编号`
                LEFT JOIN gat_logisitis_match c ON b.物流状态 = c.签收表物流状态
                LEFT JOIN gat_return d ON si.订单编号 = d.订单编号;'''
        sql = '''SELECT gat_zqsb.订单编号,gat_zqsb.系统订单状态,gat_zqsb.`系统物流状态`,gat_zqsb.`物流状态`,gat_zqsb.`最终状态`
                FROM sheet1_iphone
	            LEFT JOIN gat_zqsb ON sheet1_iphone.`订单编号` = gat_zqsb.`订单编号`;'''
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        print('正在写入excel…………')
        df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\订单检索-查询{}.xlsx'.format('1'),
                    sheet_name='查询', index=False)
        print('----已写入excel')

if __name__ == '__main__':
    m = QueryUpdate()
    start: datetime = datetime.datetime.now()
    match1 = {'slgat': '神龙-港台',
              'slgat_hfh': '火凤凰-港台',
              'slgat_hs': '红杉-港台',
              'slgat_js': '金狮-港台',
              'gat': '港台'}
    # -----------------------------------------------手动查询状态运行（一）-----------------------------------------

    m.readFormHost()
    print('输出耗时：', datetime.datetime.now() - start)