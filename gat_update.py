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
    def readFormHost(self, team):
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
        path = r'D:\Users\Administrator\Desktop\需要用到的文件\数据库'
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
                    self.replaceSqlHost(team)
                    print('++++----->>>' + sht.name + '：订单更新完成++++')
                else:
                    print('----------数据为空导入失败：' + sht.name)
            wb.close()
        app.quit()

    # 写入更新缓存表
    def writeCacheHost(self, dataFrame):
        dataFrame.to_sql('gat_update', con=self.engine1, index=False, if_exists='replace')
    # 更新总表
    def replaceSqlHost(self, team):
        try:
            print('正在更新单表中......')
            sql = '''update {0}_order_list a, gat_update b
                                set a.`运单编号`= IF(b.`运单编号` = '', NULL, b.`运单编号`),
        		                    a.`系统订单状态`= IF(b.`系统订单状态` = '', NULL, b.`系统订单状态`),
        		                    a.`系统物流状态`= IF(b.`系统物流状态` = '', NULL, b.`系统物流状态`),
        		                    a.`是否改派`= IF(b.`是否改派` = '', NULL, b.`是否改派`),
        		                    a.`物流方式`= IF(b.`物流方式` = '', NULL, b.`物流方式`),
        		                    a.`物流名称`= IF(b.`物流名称` = '', NULL, b.`物流名称`),
        		                    a.`付款方式`= IF(b.`付款方式` = '', NULL, b.`付款方式`),
        		                    a.`产品id`= IF(b.`产品id` = '', NULL, b.`产品id`),
        		                    a.`产品名称`= IF(b.`产品名称` = '', NULL, b.`产品名称`),
        		                    a.`父级分类`= IF(b.`父级分类` = '', NULL, b.`父级分类`),
        		                    a.`二级分类`= IF(b.`二级分类` = '', NULL, b.`二级分类`),
        		                    a.`三级分类`= IF(b.`三级分类` = '', NULL, b.`三级分类`)
        		                where a.`订单编号`= b.`订单编号`;'''.format(team)
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
            print('正在更新总表中......')
            sql = '''update {0}_zqsb a, gat_update b
                                            set a.`运单编号`= IF(b.`运单编号` = '', NULL, b.`运单编号`),
                    		                    a.`系统订单状态`= IF(b.`系统订单状态` = '', NULL, b.`系统订单状态`),
                    		                    a.`系统物流状态`= IF(b.`系统物流状态` = '', NULL, b.`系统物流状态`),
                    		                    a.`是否改派`= IF(b.`是否改派` = '', NULL, b.`是否改派`),
                    		                    a.`物流方式`= IF(b.`物流方式` = '', NULL, b.`物流方式`),
                    		                    a.`物流名称`= IF(b.`物流名称` = '', NULL, b.`物流名称`),
                    		                    a.`付款方式`= IF(b.`付款方式` = '', NULL, b.`付款方式`),
                    		                    a.`产品id`= IF(b.`产品id` = '', NULL, b.`产品id`),
                    		                    a.`产品名称`= IF(b.`产品名称` = '', NULL, b.`产品名称`),
                    		                    a.`父级分类`= IF(b.`父级分类` = '', NULL, b.`父级分类`),
                    		                    a.`二级分类`= IF(b.`二级分类` = '', NULL, b.`二级分类`),
                    		                    a.`三级分类`= IF(b.`三级分类` = '', NULL, b.`三级分类`)
                    		                where a.`订单编号`= b.`订单编号`;'''.format(team)
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
        except Exception as e:
            print('更新失败：', str(Exception) + str(e))
        print('更新成功…………')

    # 导出需要更新的签收表---港澳台
    def EportOrder(self, team):
        today = datetime.date.today().strftime('%Y.%m.%d')
        match = {'slgat': '神龙-港台',
                 'slgat_hfh': '火凤凰-港台',
                 'slgat_hs': '红杉-港台',
                 'slgat_js': '金狮-港台',
                 'gat': '港台',
                 'slsc': '品牌'}
        if team in ('gat'):
            month_last = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            month_yesterday = datetime.datetime.now().strftime('%Y-%m-%d')
            month_begin = (datetime.datetime.now() - relativedelta(months=3)).strftime('%Y-%m-%d')
            print(month_begin)
        else:
            month_last = '2021-05-01'
            month_yesterday = '2021-06-010'
            month_begin = '2021-02-01'
        print('正在检查产品id、父级分类为空的信息---')
        sql = '''SELECT id,日期,`订单编号`,`商品id` FROM {0}_order_list sl
            			WHERE sl.`日期`> '{1}' AND sl.`父级分类` IS NULL
            				AND ( NOT sl.`系统订单状态` IN ('已删除','问题订单','支付失败','未支付'));'''.format(team, month_begin)
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        print(df)
        df.to_sql('d1_cp', con=self.engine1, index=False, if_exists='replace')
        sql = '''SELECT 订单编号,
            			商品id,
            			dp.`id` productId,
            			dp.`name` productName,
            			dc.ppname cate,
            			dc.pname second_cate,
            			dc.`name` third_cate
            	FROM d1_cp
            	LEFT JOIN dim_product dp ON  dp.sale_id = d1_cp.商品id
            	LEFT JOIN dim_cate dc ON  dc.id = dp.third_cate_id;'''
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        df.to_sql('tem_product_id', con=self.engine1, index=False, if_exists='replace')
        print('正在更新产品详情…………')
        sql = '''update {0}_order_list a, tem_product_id b
            		    set a.`产品id`= b.`productId`,
            		        a.`产品名称`= IF(b.`productName` = '',NULL, b.`productName`),
            				a.`父级分类`= IF(b.`cate` = '',NULL, b.`cate`),
            				a.`二级分类`= IF(b.`second_cate` = '',NULL, b.`second_cate`),
            				a.`三级分类`= IF(b.`third_cate` = '',NULL, b.`third_cate`)
            			where a.`订单编号`= b.`订单编号`;'''.format(team)
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
        print('更新完成+++')

        print('正在获取---' + match[team] + ' ---更新数据内容…………')
        sql = '''SELECT 日期, a.订单编号 订单编号, a.运单编号 运单编号,IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, 
                        c.`物流状态代码` 物流状态代码,系统订单状态, 
                        IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        IF(是否改派='二次改派', '改派', 是否改派) 是否改派,
                        物流方式,物流名称,付款方式,产品id,产品名称,父级分类,二级分类,三级分类
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {0} WHERE id IN (SELECT MAX(id) FROM {0} WHERE {0}.添加时间 > '{1}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {0}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {0}_return d ON a.订单编号 = d.订单编号
                    WHERE a.日期 >= '{2}' AND a.日期 <= '{3}'
                    AND a.系统订单状态 IN ('已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)')
                    ORDER BY a.`下单时间`;'''.format(team, month_begin, month_last, month_yesterday) # 港台查询函数导出
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        print('正在写入excel…………')
        df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {} 更新-签收表.xlsx'.format(today, match[team]),
                    sheet_name=match[team], index=False)
        print('----已写入excel')

    # 导出总的签收表---港澳台
    def EportOrderBook(self, team):
        today = datetime.date.today().strftime('%Y.%m.%d')
        match = {'slgat': '神龙-港台',
                 'slgat_hfh': '火凤凰-港台',
                 'slgat_hs': '红杉-港台',
                 'slgat_js': '金狮-港台',
                 'gat': '港台',
                 'slsc': '品牌'}
        if team in ('gat'):
            month_last = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            month_yesterday = datetime.datetime.now().strftime('%Y-%m-%d')
        else:
            month_last = '2021-05-01'
            month_yesterday = '2021-06-010'
        print('正在获取---' + match[team] + ' ---全部数据内容…………')
        sql = '''SELECT * FROM {0}_zqsb a WHERE a.日期 >= '{1}' AND a.日期 <= '{2}' ORDER BY a.`下单时间`;'''.format(team, month_last, month_yesterday)     # 港台查询函数导出
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        print('正在写入---' + match[team] + ' ---临时缓存…………')             # 备用临时缓存表
        df.to_sql('d1_{0}'.format(team), con=self.engine1, index=False, if_exists='replace')

        for tem in ('"神龙家族-港澳台"|slgat', '"红杉家族-港澳台", "红杉家族-港澳台2"|slgat_hs', '"火凤凰-港澳台"|slgat_hfh', '"金狮-港澳台"|slgat_js'):
            tem1 = tem.split('|')[0]
            tem2 = tem.split('|')[1]
            sql = '''SELECT * FROM d1_{0} sl WHERE sl.`团队`in ({1});'''.format(team, tem1)
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            df.to_sql('d1_{0}'.format(tem2), con=self.engine1, index=False, if_exists='replace')
            df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}签收表.xlsx'.format(today, match[tem2]),
                        sheet_name=match[tem2], index=False)
            print(tem2 + '----已写入excel')
            print('正在打印' + match[tem2] + ' 物流时效…………')
            self.m.data_wl(tem2)

if __name__ == '__main__':
    m = QueryUpdate()
    start: datetime = datetime.datetime.now()
    match1 = {'slgat': '神龙-港台',
              'slgat_hfh': '火凤凰-港台',
              'slgat_hs': '红杉-港台',
              'slgat_js': '金狮-港台',
              'gat': '港台'}
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    team = 'gat'
    # 获取签收表---港澳台

    # m.EportOrder(team)  # 最近两个月的订单信息导出
    # print('获取耗时：', datetime.datetime.now() - start)



    # 更新签收表---港澳台

    # m.readFormHost(team)
    # print('更新耗时：', datetime.datetime.now() - start)
    m.EportOrderBook(team)  # 最近两个月的订单信息导出
    print('输出耗时：', datetime.datetime.now() - start)