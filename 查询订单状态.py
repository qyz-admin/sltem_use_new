import pandas as pd
import os
import datetime
import xlwings

import requests
import json
import sys
from sso_updata import QueryTwo
from queue import Queue
from dateutil.relativedelta import relativedelta
from threading import Thread #  使用 threading 模块创建线程
import pandas.io.formats.excel

from sqlalchemy import create_engine
from settings import Settings
from emailControl import EmailControl
from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from openpyxl.styles import Font, Border, Side, PatternFill, colors, Alignment  # 设置字体风格为Times New Roman，大小为16，粗体、斜体，颜色蓝色

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
        # self.sso = QueryTwo()
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
    def readFormHost(self, upload):
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            if dir[:2] != '~$':
                print(filePath)
                self.wbsheetHost(filePath, upload)
        print('处理耗时：', datetime.datetime.now() - start)
    # 工作表的订单信息
    def wbsheetHost(self, filePath, upload):
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
                    if upload == '查询-订单号':
                        columns_value = list(db.columns)  # 获取数据的标题名，转为列表
                        for column_val in columns_value:
                            if '订单编号' != column_val:
                                db.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                    elif upload == '查询-运单号':
                        columns_value = list(db.columns)
                        for column_val in columns_value:
                            if '运单编号' != column_val:
                                db.drop(labels=[column_val], axis=1, inplace=True)
                except Exception as e:
                    print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                if db is not None and len(db) > 0:
                    print('++++正在导入查询：' + sht.name + '表； 共：' + str(len(db)) + '行', 'sheet共：' + str(sht.used_range.last_cell.row) + '行')
                    self.writeCacheHost(db, upload)
                    print('++++正在获取：' + sht.name + '--->>>到查询缓存表')
                else:
                    print('----------数据为空导入失败：' + sht.name)
            wb.close()
        app.quit()

    # 写入更新缓存表
    def writeCacheHost(self, dataFrame, upload):
        month_begin = (datetime.datetime.now() - relativedelta(months=3)).strftime('%Y-%m-%d')
        df = None
        if upload == '查询-订单号':
            dataFrame.to_sql('sheet1_iphone', con=self.engine1, index=False, if_exists='replace')
            sql = '''SELECT gat_zqsb.订单编号,gat_zqsb.系统订单状态,gat_zqsb.`系统物流状态`,gat_zqsb.`物流状态`,gat_zqsb.`最终状态`
                            FROM sheet1_iphone
            	            LEFT JOIN gat_zqsb ON sheet1_iphone.`订单编号` = gat_zqsb.`订单编号`;'''
            df = pd.read_sql_query(sql=sql, con=self.engine1)
        elif upload == '查询-运单号':
            dataFrame.to_sql('sheet1_iphone_cy', con=self.engine1, index=False, if_exists='replace')
            print('正在获取查询数据内容…………')
            sql = '''SELECT b.订单编号, b.运单编号, IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态
                    FROM sheet1_iphone_cy a
                    LEFT JOIN (SELECT * FROM gat WHERE id IN (SELECT MAX(id) FROM gat WHERE gat.添加时间 > '{0} 00:00:00' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                    LEFT JOIN gat_logisitis_match c ON b.物流状态 = c.签收表物流状态;'''.format(month_begin)
            df = pd.read_sql_query(sql=sql, con=self.engine1)
        print('正在写入excel…………')
        rq = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
        df.to_excel('G:\\输出文件\\订单检索-查询{}.xlsx'.format(rq),sheet_name='查询', index=False)
        print('----已写入excel')


    def trans_way_cost(self, team):
        match = {'gat': '港台', 'slsc': '品牌'}
        month_yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        month_yesterday = '2021-08-12'
        print(month_yesterday)
        month_now = (datetime.datetime.now()).strftime('%Y%m')
        month_now = '202108'
        print(month_now)

        sql = '''SELECT 年月,日期,团队,币种,订单编号,数量,电话号码,运单编号,是否改派,物流方式,商品id,ds.产品id,产品名称,价格,下单时间,审核时间,仓储扫描时间,完结状态,完结状态时间,物流花费,包裹重量,包裹体积,规格中文,产品量
                        FROM gat_order_list ds
                        LEFT JOIN (SELECT 产品id, COUNT(订单编号) 产品量
                                    FROM gat_order_list ds
        					        WHERE ds.`日期` = '{0}' AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND ds.是否改派 = '直发'
        					        GROUP BY ds.`产品id`
        				) dds on ds.`产品id` = dds.`产品id`
                        WHERE ds.`年月` = '{1}' AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND ds.`包裹重量` <> 0 AND ds.`是否改派` = '直发'
        			        AND ds.`规格中文` IN (SELECT `规格中文`
                                                 FROM (SELECT 产品id,`规格中文`,COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
                                                      FROM gat_order_list d 
        											  WHERE d.`年月` = '{1}' and d.`是否改派` = '直发' AND d.`产品id` <> 0 AND d.`包裹重量` <> 0
        											    AND d.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND d.是否改派 = '直发'
        											  GROUP BY d.`产品id`,d.`规格中文`
        											  ORDER BY d. 产品id
        										    ) s1
        										WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
        									   )	
        			        AND ds.`产品id` IN (SELECT s.`产品id`
        										FROM (SELECT 年月,日期,团队,币种,订单编号,数量,电话号码,运单编号,是否改派,物流方式,商品id,产品id,产品名称,价格,下单时间,审核时间,仓储扫描时间,完结状态,完结状态时间,物流花费,包裹重量,包裹体积,规格中文
        											 FROM gat_order_list ds
        											 WHERE ds.`年月` = '{1}' AND ds.`产品id` <> 0 AND ds.`包裹重量` <> 0 and ds.`是否改派` = '直发'
        											   AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND ds.是否改派 = '直发'
        											   AND ds.`规格中文` IN (SELECT `规格中文`
        																	 FROM (SELECT 产品id,`规格中文`,COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
        																		  FROM gat_order_list d 
        																		  WHERE d.`年月` = '{1}' and d.`是否改派` = '直发' AND d.`产品id` <> 0 AND d.`包裹重量` <> 0
        																			AND d.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND d.是否改派 = '直发'
        																		  GROUP BY d.`产品id`,d.`规格中文`
        																		  ORDER BY d. 产品id
        																		 ) s1
        																	WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
        																   ) ORDER BY 日期
        											 ) s
        											 GROUP BY  s.`产品id`
        											 HAVING count(s.`产品id`) >1
        										)
                        ORDER BY 日期;;'''.format(month_yesterday, month_now)
        print('正在获取 ' + match[team] + ' 运费总直发情况…………')
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df.to_sql('gat_trans_way', con=self.engine1, index=False, if_exists='replace')
        print('正在获取运费总直发内容…………')
        sql = '''SELECT * 	
                        FROM( SELECT * 
        			            FROM gat_trans_way ds
        			            LEFT JOIN (SELECT *
        								    FROM (SELECT 产品id '产品id2',`规格中文` '规格中文2',包裹重量 '包裹重量2', COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
        											FROM gat_trans_way d 
        											GROUP BY d.`产品id`,d.`规格中文`
        											ORDER BY d. 产品id
        											) s1
        								    WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
        								    ) dss ON ds.`产品ID` = dss.`产品id2` and ds.`规格中文`= dss.`规格中文2`
                        ) s WHERE s.重量差 IS not null;'''
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df = df[['年月', '日期', '币种', '订单编号', '数量', '电话号码', '运单编号', '是否改派', '物流方式', '商品id', '产品id', '产品名称',
                 '价格', '仓储扫描时间', '完结状态', '物流花费', '包裹重量', '包裹体积', '规格中文',
                 '产品量', 'MIN(包裹重量)', 'MAX(包裹重量)', '重量差']]
        print('正在写入excel…………')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        df.to_excel('G:\\输出文件\\运费总直发-查询{}.xlsx'.format(rq),
                    sheet_name='查询', index=False)
        print('----已写入excel')

        sql = '''SELECT 年月,日期,团队,币种,订单编号,数量,电话号码,运单编号,是否改派,物流方式,商品id,ds.产品id,产品名称,价格,下单时间,审核时间,仓储扫描时间,完结状态,完结状态时间,物流花费,包裹重量,包裹体积,规格中文,产品量
                FROM gat_order_list ds
                LEFT JOIN (SELECT 产品id, COUNT(订单编号) 产品量
                            FROM gat_order_list ds
					        WHERE ds.`日期` = '{0}' AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND ds.`币种` = '台湾' AND ds.是否改派 = '直发'
					        GROUP BY ds.`产品id`
				) dds on ds.`产品id` = dds.`产品id`
                WHERE ds.`年月` = '{1}' AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND ds.`包裹重量` <> 0 AND ds.`是否改派` = '直发'
			        AND ds.`规格中文` IN (SELECT `规格中文`
                                         FROM (SELECT 产品id,`规格中文`,COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
                                              FROM gat_order_list d 
											  WHERE d.`年月` = '{1}' and d.`是否改派` = '直发' AND d.`产品id` <> 0 AND d.`包裹重量` <> 0
											    AND d.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND d.`币种` = '台湾' AND d.是否改派 = '直发'
											  GROUP BY d.`产品id`,d.`规格中文`
											  ORDER BY d. 产品id
										    ) s1
										WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
									   )	
			        AND ds.`产品id` IN (SELECT s.`产品id`
										FROM (SELECT 年月,日期,团队,币种,订单编号,数量,电话号码,运单编号,是否改派,物流方式,商品id,产品id,产品名称,价格,下单时间,审核时间,仓储扫描时间,完结状态,完结状态时间,物流花费,包裹重量,包裹体积,规格中文
											 FROM gat_order_list ds
											 WHERE ds.`年月` = '{1}' AND ds.`产品id` <> 0 AND ds.`包裹重量` <> 0 and ds.`是否改派` = '直发'
											   AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)')  AND ds.`币种` = '台湾' AND ds.是否改派 = '直发'
											   AND ds.`规格中文` IN (SELECT `规格中文`
																	 FROM (SELECT 产品id,`规格中文`,COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
																		  FROM gat_order_list d 
																		  WHERE d.`年月` = '{1}' and d.`是否改派` = '直发' AND d.`产品id` <> 0 AND d.`包裹重量` <> 0
																			AND d.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND d.`币种` = '台湾' AND d.是否改派 = '直发'
																		  GROUP BY d.`产品id`,d.`规格中文`
																		  ORDER BY d. 产品id
																		 ) s1
																	WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
																   ) ORDER BY 日期
											 ) s
											 GROUP BY  s.`产品id`
											 HAVING count(s.`产品id`) >1
										)
                ORDER BY 日期;;'''.format(month_yesterday, month_now)
        print('正在获取 ' + match[team] + ' 运费台湾直发-情况…………')
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df.to_sql('gat_trans_way', con=self.engine1, index=False, if_exists='replace')
        print('正在获取-运费台湾直发-内容…………')
        sql = '''SELECT * 	
                FROM( SELECT * 
			            FROM gat_trans_way ds
			            LEFT JOIN (SELECT *
								    FROM (SELECT 产品id '产品id2',`规格中文` '规格中文2',包裹重量 '包裹重量2', COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
											FROM gat_trans_way d 
											GROUP BY d.`产品id`,d.`规格中文`
											ORDER BY d. 产品id
											) s1
								    WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
								    ) dss ON ds.`产品ID` = dss.`产品id2` and ds.`规格中文`= dss.`规格中文2`
                ) s WHERE s.重量差 IS not null;'''
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df = df[['年月', '日期', '币种', '订单编号', '数量', '电话号码', '运单编号', '是否改派', '物流方式', '商品id', '产品id', '产品名称',
                 '价格', '仓储扫描时间', '完结状态', '物流花费', '包裹重量', '包裹体积', '规格中文',
                 '产品量', 'MIN(包裹重量)', 'MAX(包裹重量)', '重量差']]
        print('正在写入excel…………')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        df.to_excel('G:\\输出文件\\运费台湾直发-查询{}.xlsx'.format(rq),
                    sheet_name='查询', index=False)
        print('----已写入excel')

        sql = '''SELECT 年月,日期,团队,币种,订单编号,数量,电话号码,运单编号,是否改派,物流方式,商品id,ds.产品id,产品名称,价格,下单时间,审核时间,仓储扫描时间,完结状态,完结状态时间,物流花费,包裹重量,包裹体积,规格中文,产品量
                FROM gat_order_list ds
                LEFT JOIN (SELECT 产品id, COUNT(订单编号) 产品量
                            FROM gat_order_list ds
					        WHERE ds.`日期` = '{0}' AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') 
					        GROUP BY ds.`产品id`
				) dds on ds.`产品id` = dds.`产品id`
                WHERE ds.`年月` = '{1}' AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') AND ds.`包裹重量` <> 0 AND ds.`是否改派` = '直发'
			        AND ds.`规格中文` IN (SELECT `规格中文`
                                         FROM (SELECT 产品id,`规格中文`,COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
                                              FROM gat_order_list d 
											  WHERE d.`年月` = '{1}' and d.`是否改派` = '直发' AND d.`产品id` <> 0 AND d.`包裹重量` <> 0
											    AND d.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') 
											  GROUP BY d.`产品id`,d.`规格中文`
											  ORDER BY d. 产品id
										    ) s1
										WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
									   )	
			        AND ds.`产品id` IN (SELECT s.`产品id`
										FROM (SELECT 年月,日期,团队,币种,订单编号,数量,电话号码,运单编号,是否改派,物流方式,商品id,产品id,产品名称,价格,下单时间,审核时间,仓储扫描时间,完结状态,完结状态时间,物流花费,包裹重量,包裹体积,规格中文
											 FROM gat_order_list ds
											 WHERE ds.`年月` = '{1}' AND ds.`产品id` <> 0 AND ds.`包裹重量` <> 0 and ds.`是否改派` = '直发'
											   AND ds.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)')  AND ds.`币种` = '台湾' AND ds.是否改派 = '直发'
											   AND ds.`规格中文` IN (SELECT `规格中文`
																	 FROM (SELECT 产品id,`规格中文`,COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
																		  FROM gat_order_list d 
																		  WHERE d.`年月` = '{1}' and d.`是否改派` = '直发' AND d.`产品id` <> 0 AND d.`包裹重量` <> 0
																			AND d.系统订单状态 IN ('已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)') 
																		  GROUP BY d.`产品id`,d.`规格中文`
																		  ORDER BY d. 产品id
																		 ) s1
																	WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
																   ) ORDER BY 日期
											 ) s
											 GROUP BY  s.`产品id`
											 HAVING count(s.`产品id`) >1
										)
                ORDER BY 日期;;'''.format(month_yesterday, month_now)
        print('正在获取 ' + match[team] + ' 运费情况…………')
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df.to_sql('gat_trans_way', con=self.engine1, index=False, if_exists='replace')
        print('正在获取运费内容…………')
        sql = '''SELECT * 	
                FROM( SELECT * 
			            FROM gat_trans_way ds
			            LEFT JOIN (SELECT *
								    FROM (SELECT 产品id '产品id2',`规格中文` '规格中文2',包裹重量 '包裹重量2', COUNT(订单编号) 单量, MIN(包裹重量), MAX(包裹重量),  MAX(包裹重量)-MIN(包裹重量) as 重量差
											FROM gat_trans_way d 
											GROUP BY d.`产品id`,d.`规格中文`
											ORDER BY d. 产品id
											) s1
								    WHERE s1.`重量差` > 100 AND s1.`单量` >= 2
								    ) dss ON ds.`产品ID` = dss.`产品id2` and ds.`规格中文`= dss.`规格中文2`
                ) s WHERE s.重量差 IS not null;'''
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df = df[['年月', '日期', '币种', '订单编号', '数量', '电话号码', '运单编号', '是否改派', '物流方式', '商品id', '产品id', '产品名称',
                 '价格', '仓储扫描时间', '完结状态', '物流花费', '包裹重量', '包裹体积', '规格中文',
                 '产品量', 'MIN(包裹重量)', 'MAX(包裹重量)', '重量差']]
        print('正在写入excel…………')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        df.to_excel('G:\\输出文件\\运费-查询{}.xlsx'.format(rq),
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
    team = 'gat'
    # -----------------------------------------------手动查询状态运行（一）-----------------------------------------
    # upload = '查询-订单号'
    upload = '查询-运单号'
    m.readFormHost(upload)


    # m.trans_way_cost(team)  # 同产品下的规格运费查询
    print('输出耗时：', datetime.datetime.now() - start)

