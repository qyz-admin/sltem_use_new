import pandas as pd
from sqlalchemy import create_engine
from settings import Settings
from queryControl import QueryControl
from emailControl import EmailControl
from bpsControl import BpsControl
from sltemMonitoring import SltemMonitoring

from openpyxl import load_workbook  # 可以向不同的sheet写入数据
from tkinter import messagebox
import os
import zipfile
import xlwings as xl
import datetime
from dateutil.relativedelta import relativedelta
import time


class MysqlControl(Settings):
    def __init__(self):
        Settings.__init__(self)
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
        self.d = QueryControl()

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
        elif team == 'slgat':
            print(team + '---909')  # 当天和前天的添加时间比较，判断是否一样数据
            sql = 'INSERT IGNORE INTO {}({}, 添加时间, 更新时间) SELECT *, CURDATE() 添加时间, NOW() 更新时间 FROM tem; '.format(team,
                                                                                                                 columns)
        else:
            print(team)
            sql = 'INSERT IGNORE INTO {}({}, 添加时间) SELECT *, NOW() 添加时间 FROM tem; '.format(team, columns)
            # sql = 'INSERT IGNORE INTO {}_copy({}, 添加时间) SELECT *, NOW() 添加时间 FROM tem; '.format(team, columns)
        try:
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=100)
        except Exception as e:
            print('插入失败：', str(Exception) + str(e))

    def readSql(self, sql):
        db = pd.read_sql_query(sql=sql, con=self.engine1)
        # db = pd.read_sql(sql=sql, con=self.engine1) or team == 'slgat'
        return db

    def update_gk_product(self):        # 更新产品id的列表
        (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1))
        yy = int((datetime.datetime.now() - datetime.timedelta(days=15)).strftime('%Y'))
        mm = int((datetime.datetime.now() - datetime.timedelta(days=15)).strftime('%m'))
        dd = int((datetime.datetime.now() - datetime.timedelta(days=15)).strftime('%d'))
        begin = datetime.date(yy, mm, dd)
        # begin = datetime.date(2021, 3, 10)
        print(begin)
        yy2 = int(datetime.datetime.now().strftime('%Y'))
        mm2 = int(datetime.datetime.now().strftime('%m'))
        dd2 = int(datetime.datetime.now().strftime('%d'))
        end = datetime.date(yy2, mm2, dd2)
        # end = datetime.date(2021, 4, 19)
        print(end)
        for i in range((end - begin).days):  # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            month_last = str(day)
            # sql = '''SELECT * FROM  gk_product WHERE gk_product.rq >= '{0}';'''.format(month_last)
            sql = '''SELECT product_id,
				            rq,
				            product_name,
				            cate_id,
				            second_cate_id,
				            third_cate_id,
				            null seller_id,
				            null selector,
				            null buyer_id,
				            price,
				            gs.`status`
		            FROM gk_sale gs
		            WHERE gs.rq >= '{0}';'''.format(month_last)
            print('正在获取 ' + month_last + ' 号以后的产品详情…………')
            df = pd.read_sql_query(sql=sql, con=self.engine2)
            print('正在写入产品缓存中…………')
            df.to_sql('tem_product', con=self.engine1, index=False, if_exists='replace')
            try:
                print('正在更新中…………')
                sql = 'REPLACE INTO dim_product SELECT *, NOW() 更新时间 FROM tem_product; '
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
            except Exception as e:
                print('插入失败：', str(Exception) + str(e))
            print('更新完成…………')

    def creatMyOrderSl(self, team):  # 最近五天的全部订单信息
        match = {'slgat': '"神龙家族-港澳台"',
                 'slgat_hfh': '"火凤凰-港澳台"',
                 'slgat_hs': '"红杉家族-港澳台", "红杉家族-港澳台2"',
                 'sltg': '"神龙家族-泰国"',
                 'slxmt': '"神龙家族-新加坡", "神龙家族-马来西亚", "神龙家族-菲律宾"',
                 'slxmt_t': '"神龙-T新马菲"',
                 'slxmt_hfh': '"火凤凰-新加坡", "火凤凰-马来西亚", "火凤凰-菲律宾"',
                 'slrb': '"神龙家族-日本团队"',
                 'slrb_js': '"金狮-日本"',
                 'slrb_hs': '"红杉家族-日本", "红杉家族-日本666"',
                 'slrb_jl': '"精灵家族-日本", "精灵家族-韩国", "精灵家族-品牌"'}
        # 12-1月的
        if team in ('sltg', 'slrb', 'slrb_jl', 'slrb_js', 'slrb_hs', 'slgat', 'slgat_hfh', 'slgat_hs', 'slxmt', 'slxmt_t', 'slxmt_hfh'):
            # 获取日期时间
            sql = 'SELECT 日期 FROM {0}_order_list WHERE id = (SELECT MAX(id) FROM {0}_order_list);'.format(team)
            rq = pd.read_sql_query(sql=sql, con=self.engine1)
            rq = pd.to_datetime(rq['日期'][0])
            yy = int((rq - datetime.timedelta(days=3)).strftime('%Y'))
            mm = int((rq - datetime.timedelta(days=3)).strftime('%m'))
            dd = int((rq - datetime.timedelta(days=3)).strftime('%d'))
            print(dd)
            begin = datetime.date(yy, mm, dd)
            print(begin)
            yy2 = int(datetime.datetime.now().strftime('%Y'))
            mm2 = int(datetime.datetime.now().strftime('%m'))
            dd2 = int(datetime.datetime.now().strftime('%d'))
            end = datetime.date(yy2, mm2, dd2)
            print(end)
        else:
            # 11-12月的
            begin = datetime.date(2021, 4, 1)
            print(begin)
            end = datetime.date(2021, 5, 6)
            print(end)
        for i in range((end - begin).days):  # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            # print(str(day))
            yesterday = str(day) + ' 23:59:59'
            last_month = str(day)
            if team in ('slxmt', 'slxmt_t', 'slxmt_hfh'):
                sql = '''SELECT a.id,
                            a.month 年月,
                            a.month_mid 旬,
                            a.rq 日期,
                            dim_area.name 团队,
                            a.region_code 区域,
                            dim_currency_lang.pname 币种,
                            a.beform 订单来源,
                            a.order_number 订单编号,
                            a.qty 数量,
                            a.ship_phone 电话号码,
                            UPPER(a.waybill_number) 运单编号,
                            a.order_status 系统订单状态id,
                            a.logistics_status 系统物流状态id,
                            IF(a.second=0,'直发','改派') 是否改派,
                            dim_trans_way.all_name 物流方式,
                            dim_trans_way.simple_name 物流名称,
                            dim_trans_way.remark 运输方式,
                            a.logistics_type 货物类型,
                            IF(a.low_price=0,'否','是') 是否低价,
                            a.product_id 产品id,
             		        gk_sale.product_name 产品名称,
                            dim_cate.ppname 父级分类,
                            dim_cate.pname 二级分类,
                            dim_cate.name 三级分类,
                            dim_payment.pay_name 付款方式,
                            a.amount 价格,
                            a.addtime 下单时间,
                            a.verity_time 审核时间,
                            a.delivery_time 仓储扫描时间,
                            a.finish_status 完结状态,
                            a.endtime 完结状态时间,
                            a.salesRMB 价格RMB,
                            intervals.intervals 价格区间,
                            null 成本价,
                            a.logistics_cost 物流花费,
                            null 打包花费,
                            a.other_fee 其它花费,
                            a.weight 包裹重量,
                            a.volume 包裹体积,
                            a.ship_zip 邮编,
                            a.turn_purchase_time 添加物流单号时间,
                            a.del_reason 订单删除原因,
                            a.ship_state 省洲
                    FROM gk_order a
                            left join dim_area ON dim_area.id = a.area_id
                            left join dim_payment ON dim_payment.id = a.payment_id
	                        LEFT JOIN gk_sale ON gk_sale.id = a.sale_id
             		--		left join (SELECT * FROM gk_sale WHERE id IN (SELECT MAX(id) FROM gk_sale GROUP BY product_id ) ORDER BY id) gs ON gs.product_id = a.product_id
                            left join dim_trans_way ON dim_trans_way.id = a.logistics_id
                            left join dim_cate ON dim_cate.id = a.third_cate_id
                            left join intervals ON intervals.id = a.intervals
                            left join dim_currency_lang ON dim_currency_lang.id = a.currency_lang_id
                    WHERE  a.rq = '{0}' AND a.rq <= '{1}'
                        AND dim_area.name IN ({2});'''.format(last_month, yesterday, match[team])
            elif team in ('slrb_jl', 'slrb_js', 'slrb_hs'):
                sql = '''SELECT a.id,
                            a.month 年月,
                            a.month_mid 旬,
                            a.rq 日期,
                            dim_area.name 团队,
                            a.region_code 区域,
                            dim_currency_lang.pname 币种,
                            a.beform 订单来源,
                            a.order_number 订单编号,
                            a.qty 数量,
                            a.ship_phone 电话号码,
                            UPPER(a.waybill_number) 运单编号,
                            a.order_status 系统订单状态id,
                            a.logistics_status 系统物流状态id,
                            IF(a.second=0,'直发','改派') 是否改派,
                            dim_trans_way.all_name 物流方式,
                            dim_trans_way.simple_name 物流名称,
                            dim_trans_way.remark 运输方式,
                            a.logistics_type 货物类型,
                            IF(a.low_price=0,'否','是') 是否低价,
                            a.product_id 产品id,
             		        gk_sale.product_name 产品名称,
            --              e.`name` 产品名称,
                            dim_cate.ppname 父级分类,
                            dim_cate.pname 二级分类,
                            dim_cate.name 三级分类,
                            dim_payment.pay_name 付款方式,
                            a.amount 价格,
                            a.addtime 下单时间,
                            a.verity_time 审核时间,
                            a.delivery_time 仓储扫描时间,
                            a.finish_status 完结状态,
                            a.endtime 完结状态时间,
                            a.salesRMB 价格RMB,
                            intervals.intervals 价格区间,
                            null 成本价,
                            a.logistics_cost 物流花费,
                            null 打包花费,
                            a.other_fee 其它花费,
                            a.weight 包裹重量,
                            a.volume 包裹体积,
                            a.ship_zip 邮编,
                            a.turn_purchase_time 添加物流单号时间,
                            a.del_reason 订单删除原因,
                            IF(dim_area.name = '精灵家族-品牌',IF(a.coll_id=1000000269,'饰品','内衣'),a.coll_id) 站点ID
                    FROM gk_order a
                            left join dim_area ON dim_area.id = a.area_id
                            left join dim_payment ON dim_payment.id = a.payment_id
	                        LEFT JOIN gk_sale ON gk_sale.id = a.sale_id
             		--		left join (SELECT * FROM gk_sale WHERE id IN (SELECT MAX(id) FROM gk_sale GROUP BY product_id ) ORDER BY id) gs ON gs.product_id = a.product_id
                            left join dim_trans_way ON dim_trans_way.id = a.logistics_id
                            left join dim_cate ON dim_cate.id = a.third_cate_id
                            left join intervals ON intervals.id = a.intervals
                            left join dim_currency_lang ON dim_currency_lang.id = a.currency_lang_id
                    WHERE  a.rq = '{0}' AND a.rq <= '{1}'
                        AND dim_area.name IN ({2});'''.format(last_month, yesterday, match[team])
            else:
                sql = '''SELECT a.id,
                            a.month 年月,
                            a.month_mid 旬,
                            a.rq 日期,
                            dim_area.name 团队,
                            a.region_code 区域,
                            dim_currency_lang.pname 币种,
                            a.beform 订单来源,
                            a.order_number 订单编号,
                            a.qty 数量,
                            a.ship_phone 电话号码,
                            UPPER(a.waybill_number) 运单编号,
                            a.order_status 系统订单状态id,
                            a.logistics_status 系统物流状态id,
                            IF(a.second=0,'直发','改派') 是否改派,
                            dim_trans_way.all_name 物流方式,
                            dim_trans_way.simple_name 物流名称,
                            dim_trans_way.remark 运输方式,
                            a.logistics_type 货物类型,
                            IF(a.low_price=0,'否','是') 是否低价,
                            a.product_id 产品id,
             		        gk_sale.product_name 产品名称,
            --              e.`name` 产品名称,
                            dim_cate.ppname 父级分类,
                            dim_cate.pname 二级分类,
                            dim_cate.name 三级分类,
                            dim_payment.pay_name 付款方式,
                            a.amount 价格,
                            a.addtime 下单时间,
                            a.verity_time 审核时间,
                            a.delivery_time 仓储扫描时间,
                            a.finish_status 完结状态,
                            a.endtime 完结状态时间,
                            a.salesRMB 价格RMB,
                            intervals.intervals 价格区间,
                            null 成本价,
                            a.logistics_cost 物流花费,
                            null 打包花费,
                            a.other_fee 其它花费,
                            a.weight 包裹重量,
                            a.volume 包裹体积,
                            a.ship_zip 邮编,
                            a.turn_purchase_time 添加物流单号时间,
                            a.del_reason 订单删除原因
                    FROM gk_order a
                            left join dim_area ON dim_area.id = a.area_id
                            left join dim_payment ON dim_payment.id = a.payment_id
	                        LEFT JOIN gk_sale ON gk_sale.id = a.sale_id
             		--		left join (SELECT * FROM gk_sale WHERE id IN (SELECT MAX(id) FROM gk_sale GROUP BY product_id ) ORDER BY id) gs ON gs.product_id = a.product_id
                            left join dim_trans_way ON dim_trans_way.id = a.logistics_id
                            left join dim_cate ON dim_cate.id = a.third_cate_id
                            left join intervals ON intervals.id = a.intervals
                            left join dim_currency_lang ON dim_currency_lang.id = a.currency_lang_id
                    WHERE  a.rq = '{0}' AND a.rq <= '{1}'
                        AND dim_area.name IN ({2});'''.format(last_month, yesterday, match[team])
            print('正在获取 ' + match[team] + last_month[5:7] + '-' + yesterday[8:10] + ' 号订单…………')
            df = pd.read_sql_query(sql=sql, con=self.engine2)
            sql = 'SELECT * FROM dim_order_status;'
            df1 = pd.read_sql_query(sql=sql, con=self.engine1)
            print('+++合并订单状态中…………')
            df = pd.merge(left=df, right=df1, left_on='系统订单状态id', right_on='id', how='left')
            sql = 'SELECT * FROM dim_logistics_status;'
            df1 = pd.read_sql_query(sql=sql, con=self.engine1)
            print('+++合并物流状态中…………')
            df = pd.merge(left=df, right=df1, left_on='系统物流状态id', right_on='id', how='left')
            df = df.drop(labels=['id', 'id_y', '系统订单状态id', '系统物流状态id'], axis=1)
            df.rename(columns={'id_x': 'id', 'name_x': '系统订单状态', 'name_y': '系统物流状态'}, inplace=True)
            print('++++++正在将 ' + yesterday[8:10] + ' 号订单写入数据库++++++')
            # 这一句会报错,需要修改my.ini文件中的[mysqld]段中的"max_allowed_packet = 1024M"
            try:
                df.to_sql('sl_order', con=self.engine1, index=False, if_exists='replace')
                sql = 'REPLACE INTO {}_order_list SELECT *, NOW() 记录时间 FROM sl_order; '.format(team)
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
            except Exception as e:
                print('插入失败：', str(Exception) + str(e))
            print('写入完成…………')
        return '写入完成'

    def creatMyOrderSlTWO(self, team):  # 最近两个月的更新订单信息
        match = {'slgat': '"神龙家族-港澳台"',
                 'slgat_hfh': '"火凤凰-港澳台"',
                 'slgat_hs': '"红杉家族-港澳台", "红杉家族-港澳台2"',
                 'sltg': '"神龙家族-泰国"',
                 'slxmt': '"神龙家族-新加坡", "神龙家族-马来西亚", "神龙家族-菲律宾"',
                 'slxmt_t': '"神龙-T新马菲"',
                 'slxmt_hfh': '"火凤凰-新加坡", "火凤凰-马来西亚", "火凤凰-菲律宾"',
                 'slrb': '"神龙家族-日本团队"',
                 'slrb_js': '"金狮-日本"',
                 'slrb_hs': '"红杉家族-日本", "红杉家族-日本666"',
                 'slrb_jl': '"精灵家族-日本", "精灵家族-韩国", "精灵家族-品牌"'}
        today = datetime.date.today().strftime('%Y.%m.%d')
        if team in ('sltg', 'slrb0', 'slrb_jl0', 'slrb_js0', 'slrb_hs0', 'slgat0', 'slgat_hfh0', 'slgat_hs0', 'slxmt0', 'slxmt_t0', 'slxmt_hfh0'):
            yy = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y'))
            mm = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%m'))
            begin = datetime.date(yy, mm, 1)
            print(begin)
            yy2 = int(datetime.datetime.now().strftime('%Y'))
            mm2 = int(datetime.datetime.now().strftime('%m'))
            dd2 = int(datetime.datetime.now().strftime('%d'))
            end = datetime.date(yy2, mm2, dd2)
            print(end)
        else:
            begin = datetime.date(2021, 3, 1)
            print(begin)
            end = datetime.date(2021, 4, 1)
            print(end)
        for i in range((end - begin).days):  # 按天循环获取订单状态
            day = begin + datetime.timedelta(days=i)
            # print(str(day))
            yesterday = str(day) + ' 23:59:59'
            last_month = str(day)
            sql = '''SELECT DISTINCT a.id,
                            a.rq 日期,
                            dim_currency_lang.pname 币种,
                            a.order_number 订单编号,
                            a.qty 数量,
                            a.ship_phone 电话号码,
                            UPPER(a.waybill_number) 运单编号,
                            a.order_status 系统订单状态id,
                            a.logistics_status 系统物流状态id,
                            IF(a.second=0,'直发','改派') 是否改派,
                            dim_trans_way.all_name 物流方式,
                            dim_trans_way.simple_name 物流名称,
                            a.logistics_type 货物类型,
                            a.verity_time 审核时间,
                            a.delivery_time 仓储扫描时间,
                            a.endtime 完结状态时间
                    FROM gk_order a
                            left join dim_area ON dim_area.id = a.area_id
                            left join dim_payment on dim_payment.id = a.payment_id
                            LEFT JOIN gk_sale ON gk_sale.id = a.sale_id
             		--	    left join (SELECT * FROM gk_sale WHERE id IN (SELECT MAX(id) FROM gk_sale GROUP BY product_id ) ORDER BY id) gs ON gs.product_id = a.product_id
                            left join dim_trans_way on dim_trans_way.id = a.logistics_id
                            left join dim_cate on dim_cate.id = a.third_cate_id
                            left join intervals on intervals.id = a.intervals
                            left join dim_currency_lang on dim_currency_lang.id = a.currency_lang_id
                    WHERE a.rq = '{0}' AND a.rq <= '{1}'
                        AND dim_area.name IN ({2});'''.format(last_month, yesterday, match[team])
            print('正在更新 ' + match[team] + last_month[5:7] + '-' + yesterday[8:10] + ' 号订单信息…………')
            df = pd.read_sql_query(sql=sql, con=self.engine2)
            sql = 'SELECT * FROM dim_order_status;'
            df1 = pd.read_sql_query(sql=sql, con=self.engine1)
            print('++++更新订单状态中…………')
            df = pd.merge(left=df, right=df1, left_on='系统订单状态id', right_on='id', how='left')
            sql = 'SELECT * FROM dim_logistics_status;'
            df1 = pd.read_sql_query(sql=sql, con=self.engine1)
            print('++++更新物流状态中…………')
            df = pd.merge(left=df, right=df1, left_on='系统物流状态id', right_on='id', how='left')
            df = df.drop(labels=['id', 'id_y', '系统订单状态id', '系统物流状态id'], axis=1)
            df.rename(columns={'id_x': 'id', 'name_x': '系统订单状态', 'name_y': '系统物流状态'}, inplace=True)
            print('+++++++正在将 ' + yesterday[8:10] + ' 号订单更新到数据库++++++')
            # 这一句会报错,需要修改my.ini文件中的[mysqld]段中的"max_allowed_packet = 1024M"
            df.to_sql('sl_order2', con=self.engine1, index=False, if_exists='replace')
            try:
                sql = '''update {0}_order_list a, sl_order2 b
                        set a.`币种`=b.`币种`,
                            a.`数量`=b.`数量`,
		                    a.`电话号码`=b.`电话号码` ,
		                    a.`运单编号`=b.`运单编号`,
		                    a.`系统订单状态`=b.`系统订单状态`,
		                    a.`系统物流状态`=b.`系统物流状态`,
		                    a.`是否改派`=b.`是否改派`,
		                    a.`物流方式`=b.`物流方式`,
		                    a.`物流名称`=b.`物流名称`,
		                    a.`货物类型`=b.`货物类型`,
		                    a.`审核时间`=b.`审核时间`,
		                    a.`仓储扫描时间`=b.`仓储扫描时间`,
		                    a.`完结状态时间`=b.`完结状态时间`
		                where a.`订单编号`=b.`订单编号`;'''.format(team)
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
            except Exception as e:
                print('插入失败：', str(Exception) + str(e))
            print('----更新完成----')
        return '更新完成'

    def connectOrder(self, team):
        match = {'slgat': '神龙-港台',
                 'slgat_hfh': '火凤凰-港台',
                 'slgat_hs': '红杉-港台',
                 'sltg': '神龙-泰国',
                 'slrb': '神龙-日本',
                 'slrb_jl': '精灵-日本',
                 'slrb_js': '金狮-日本',
                 'slrb_hs': '红杉-日本',
                 'slxmt': '神龙-新马',
                 'slxmt_t': '神龙-T新马',
                 'slxmt_hfh': '火凤凰-新马'}
        emailAdd = {'slgat': 'giikinliujun@163.com',
                    'slgat_hfh': 'giikinliujun@163.com',
                    'slgat_hs': 'giikinliujun@163.com',
                    'sltg': 'zhangjing@giikin.com',
                    'slxmt': 'zhangjing@giikin.com',
                    'slxmt_t': 'zhangjing@giikin.com',
                    'slxmt_hfh': 'zhangjing@giikin.com',
                    'slrb': 'sunyaru@giikin.com',
                    'slrb_js': 'sunyaru@giikin.com',
                    'slrb_hs': 'sunyaru@giikin.com',
                    'slrb_jl': 'sunyaru@giikin.com'}
        if team in ('sltg', 'slrb', 'slrb_jl', 'slrb_js', 'slrb_hs', 'slgat', 'slgat_hfh', 'slgat_hs', 'slxmt0', 'slxmt_t0', 'slxmt_hfh0'):
            month_last = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            month_yesterday = datetime.datetime.now().strftime('%Y-%m-%d')
            month_begin = (datetime.datetime.now() - relativedelta(months=3)).strftime('%Y-%m-%d')
            print(month_begin)
        else:
            month_last = '2021-04-01'
            month_yesterday = '2021-04-30'
            month_begin = '2020-01-01'
        token = '6e80eb95fb6aaefeed0d2fd5ba20fb4a'        # 补充查询产品信息需要
        if team == 'slgat':  # 港台查询函数导出
            self.d.productIdInfo(token, '订单号', team)   # 产品id详情更新   （参数一需要手动更换）
            self.d.cateIdInfo(token, team)  # 进入产品检索界面（参数一需要手动更换）
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                        IF(出货时间='1990-01-01 00:00:00' or 出货时间='1899-12-29 00:00:00' or 出货时间='1899-12-30 00:00:00' or 出货时间='0000-00-00 00:00:00', a.仓储扫描时间, 出货时间) 出货时间,
                        IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,IF(状态时间='1990-01-01 00:00:00' or 状态时间='1899-12-30 00:00:00' or 状态时间='0000-00-00 00:00:00', '', 状态时间) 状态时间,
                        IF(上线时间='1990-01-01 00:00:00' or 上线时间='1899-12-29 00:00:00' or 上线时间='1899-12-30 00:00:00' or 上线时间='0000-00-00 00:00:00', '', 上线时间) 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态,
                        IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        IF(是否改派='二次改派', '改派', 是否改派) 是否改派,物流方式,物流名称,运输方式,货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                        二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                        包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                        b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态, b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间,数量
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {0} WHERE id IN (SELECT MAX(id) FROM {0} WHERE {0}.添加时间 > '{1}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {0}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {0}_return d ON a.订单编号 = d.订单编号
                    WHERE a.日期 >= '{2}' AND a.日期 <= '{3}'
                        AND a.系统订单状态 IN ('已审核', '待发货', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)', '待发货转审核')
                    ORDER BY a.`下单时间`;'''.format(team, month_begin, month_last, month_yesterday)
        elif team in ('slgat_hfh', 'slgat_hs'):  # 新马物流查询函数导出
            self.d.productIdInfo(token, '订单号', team)
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                        IF(出货时间='1990-01-01 00:00:00' or 出货时间='1899-12-29 00:00:00' or 出货时间='1899-12-30 00:00:00' or 出货时间='0000-00-00 00:00:00', a.仓储扫描时间, 出货时间) 出货时间,
                        IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,IF(状态时间='1990-01-01 00:00:00' or 状态时间='1899-12-30 00:00:00' or 状态时间='0000-00-00 00:00:00', '', 状态时间) 状态时间,
                        IF(上线时间='1990-01-01 00:00:00' or 上线时间='1899-12-29 00:00:00' or 上线时间='1899-12-30 00:00:00' or 上线时间='0000-00-00 00:00:00', '', 上线时间) 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态,
                        IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        IF(是否改派='二次改派', '改派', 是否改派) 是否改派,物流方式,物流名称,运输方式,货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                        二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                        包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                        b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态, b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间,数量
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {1} WHERE id IN (SELECT MAX(id) FROM {1} WHERE {1}.添加时间 > '{2}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {1}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {1}_return d ON a.订单编号 = d.订单编号
                    WHERE a.日期 >= '{3}' AND a.日期 <= '{4}'
                        AND a.系统订单状态 IN ('已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)','已退货(物流)', '已退货(不拆包物流)')
                    ORDER BY a.`下单时间`;'''.format(team, 'slgat', month_begin, month_last, month_yesterday)
        elif team == 'slxmt':  # 新马物流查询函数导出
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                        IF(ISNULL(b.出货时间) or b.出货时间='1899-12-29 00:00:00' or b.出货时间='0000-00-00 00:00:00' or b.状态时间='1990-01-01 00:00:00', g.出货时间, b.出货时间) 出货时间, IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,
                        IF(b.状态时间='1990-01-01 00:00:00' or b.状态时间='1899-12-30 00:00:00' or b.状态时间='0000-00-00 00:00:00', '', b.状态时间) 状态时间, 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态, IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        是否改派,物流方式,物流名称,运输方式,货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                        二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                        包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                        b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态, b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间,数量, a.省洲
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {0} WHERE id IN (SELECT MAX(id) FROM {0} WHERE {0}.添加时间 > '{1}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {0}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {0}_return d ON a.订单编号 = d.订单编号
                        LEFT JOIN (SELECT * FROM {0}wl WHERE id IN (SELECT MAX(id) FROM {0}wl  WHERE {0}wl.添加时间 > '{1}' GROUP BY 运单编号) ORDER BY id) g ON a.运单编号 = g.运单编号
                    WHERE a.日期 >= '{2}' AND a.日期 <= '{3}'
                        AND a.系统订单状态 IN ('已审核', '待发货', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)', '待发货转审核')
                    ORDER BY a.`下单时间`;'''.format(team, month_begin, month_last, month_yesterday)
        elif team == 'slxmt_hfh' or team == 'slxmt_t':
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                        IF(ISNULL(b.出货时间) or b.出货时间='1899-12-29 00:00:00' or b.出货时间='0000-00-00 00:00:00' or b.状态时间='1990-01-01 00:00:00', g.出货时间, b.出货时间) 出货时间, IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,
                        IF(b.状态时间='1990-01-01 00:00:00' or b.状态时间='1899-12-30 00:00:00' or b.状态时间='0000-00-00 00:00:00', '', b.状态时间) 状态时间, 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态, IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        是否改派,物流方式,物流名称,运输方式,货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                        二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                        包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                        b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态, b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间,数量, a.省洲
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {1} WHERE id IN (SELECT MAX(id) FROM {1} WHERE {1}.添加时间 > '{2}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {1}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {1}_return d ON a.订单编号 = d.订单编号
                        LEFT JOIN (SELECT * FROM {1}wl WHERE id IN (SELECT MAX(id) FROM {1}wl  WHERE {1}wl.添加时间 > '{2}' GROUP BY 运单编号) ORDER BY id) g ON a.运单编号 = g.运单编号
                    WHERE a.日期 >= '{3}' AND a.日期 <= '{4}'
                        AND a.系统订单状态 IN ('已审核', '待发货', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)', '待发货转审核')
                    ORDER BY a.`下单时间`;'''.format(team, 'slxmt', month_begin, month_last, month_yesterday)
        elif team == 'sltg':
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                            IF(出货时间='1990-01-01 00:00:00' or 出货时间='1899-12-29 00:00:00' or 出货时间='1899-12-30 00:00:00' or 出货时间='0000-00-00 00:00:00', null, 出货时间) 出货时间,
                            IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,IF(状态时间='1990-01-01 00:00:00' or 状态时间='1899-12-30 00:00:00' or 状态时间='0000-00-00 00:00:00', '', 状态时间) 状态时间,
                            IF(上线时间='1990-01-01 00:00:00' or 上线时间='1899-12-29 00:00:00' or 上线时间='1899-12-30 00:00:00' or 上线时间='0000-00-00 00:00:00', '', 上线时间) 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态,
                            IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                            IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                            是否改派,物流方式,物流名称,运输方式,货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                            二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                            包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                            b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态, b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间, 数量, b.问题明细
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {0} WHERE id IN (SELECT MAX(id) FROM {0} WHERE {0}.添加时间 > '{1}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {0}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {0}_return d ON a.订单编号 = d.订单编号
                    WHERE a.日期 >= '{2}' AND a.日期 <= '{3}'
                        AND a.系统订单状态 IN ('已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)')
                    ORDER BY a.`下单时间`;'''.format(team, month_begin, month_last, month_yesterday)
        elif team in ('slrb_jl', 'slrb_js', 'slrb_hs'):
            self.d.productIdInfo(token, '订单号', team)   # 产品id详情更新   （参数一需要手动更换）
            self.d.cateIdInfo(token, team)  # 进入产品检索界面（参数一需要手动更换）
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                        IF(出货时间='1990-01-01 00:00:00' or 出货时间='1899-12-29 00:00:00' or 出货时间='1899-12-30 00:00:00' or 出货时间='0000-00-00 00:00:00', null, 出货时间) 出货时间,
                        IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,IF(状态时间='1990-01-01 00:00:00' or 状态时间='1899-12-30 00:00:00' or 状态时间='0000-00-00 00:00:00', '', 状态时间) 状态时间,
                        IF(上线时间='1990-01-01 00:00:00' or 上线时间='1899-12-29 00:00:00' or 上线时间='1899-12-30 00:00:00' or 上线时间='0000-00-00 00:00:00', '', 上线时间) 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态,
                        IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        是否改派,物流方式,物流名称,运输方式,'货到付款' AS 货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                        二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                        包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                        b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态,b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间, 数量, a.站点ID
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {1} WHERE id IN (SELECT MAX(id) FROM {1} WHERE {1}.添加时间 > '{2}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {1}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {1}_return d ON a.订单编号 = d.订单编号
                    WHERE a.日期 >= '{3}' AND a.日期 <= '{4}'
                        AND a.系统订单状态 IN ('已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)')
                    ORDER BY a.`下单时间`;'''.format(team, 'slrb', month_begin, month_last, month_yesterday)
        else:
            self.d.productIdInfo(token, '订单号', team)
            self.d.cateIdInfo(token, team)  # 进入产品检索界面（参数一需要手动更换）
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                        IF(出货时间='1990-01-01 00:00:00' or 出货时间='1899-12-29 00:00:00' or 出货时间='1899-12-30 00:00:00' or 出货时间='0000-00-00 00:00:00', null, 出货时间) 出货时间,
                        IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,IF(状态时间='1990-01-01 00:00:00' or 状态时间='1899-12-30 00:00:00' or 状态时间='0000-00-00 00:00:00', '', 状态时间) 状态时间,
                        IF(上线时间='1990-01-01 00:00:00' or 上线时间='1899-12-29 00:00:00' or 上线时间='1899-12-30 00:00:00' or 上线时间='0000-00-00 00:00:00', '', 上线时间) 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态,
                        IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        是否改派,物流方式,物流名称,运输方式,货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                        二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                        包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                        b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态,b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间,数量
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {0} WHERE id IN (SELECT MAX(id) FROM {0} WHERE {0}.添加时间 > '{1}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {0}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {0}_return d ON a.订单编号 = d.订单编号
                    WHERE a.日期 >= '{2}' AND a.日期 <= '{3}'
                    AND a.系统订单状态 IN ('已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)')
                    ORDER BY a.`下单时间`;'''.format(team, month_begin, month_last, month_yesterday)
        print('正在获取---' + match[team] + ' ---全部导出数据内容…………')
        df = pd.read_sql_query(sql=sql, con=self.engine1)
        # 备用临时缓存表
        print('正在写入---' + match[team] + ' ---临时缓存…………')
        if team == 'sl':
            df.to_sql('d1', con=self.engine1, index=False, if_exists='replace')
        else:
            df.to_sql('d1_{0}'.format(team), con=self.engine1, index=False, if_exists='replace')
        today = datetime.date.today().strftime('%Y.%m.%d')
        print('正在写入excel…………')
        df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}签收表.xlsx'.format(today, match[team]),
                    sheet_name=match[team], index=False)
        print('----已写入excel')
        filePath = ['D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}签收表.xlsx'.format(today, match[team])]
        print('输出文件成功…………')
        # 文件太大无法发送的
        if team in ('sltg', 'slrb_jl', 'slgat0'):
            print('---' + match[team] + ' 不发送邮件')
        else:
            self.e.send('{} {}签收表.xlsx'.format(today, match[team]), filePath,
                        emailAdd[team])
        # 导入签收率表中和输出物流时效（不包含全部的订单状态）
        print('正在打印' + match[team] + ' 物流时效…………')
        if team == 'sltg0':
            print('---' + match[team] + ' 不打印文件')
        else:
            self.data_wl(team)
        print('正在写入' + match[team] + ' 全部签收表中…………')
        if team == 'slrb':
            sql = 'REPLACE INTO {0}_zqsb_rb SELECT *, NOW() 更新时间 FROM d1_{0};'.format(team)
        else:
            sql = 'REPLACE INTO {0}_zqsb SELECT *, NOW() 更新时间 FROM d1_{0};'.format(team)
        pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
        print('----已写入' + match[team] + '全部签收表中')

        # 商城订单的获取---暂时使用的
        if team == 'slgat':  # IG和UP订单
            emailAdd2 = {'slgat': 'service@upiinmall.com'}
            today = datetime.date.today().strftime('%Y.%m.%d')
            month_last = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            month_yesterday = datetime.datetime.now().strftime('%Y-%m-%d')
            month_begin = (datetime.datetime.now() - relativedelta(months=3)).strftime('%Y-%m-%d')
            print(month_begin)
            # month_last = '2021-03-01'
            # month_yesterday = '2021-04-06'
            # month_begin = '2020-12-01'
            sql = '''SELECT 年月, 旬, 日期, 团队,币种, 区域, 订单来源, a.订单编号 订单编号, 电话号码, a.运单编号 运单编号,
                        IF(出货时间='1990-01-01 00:00:00' or 出货时间='1899-12-30 00:00:00' or 出货时间='0000-00-00 00:00:00', a.仓储扫描时间, 出货时间) 出货时间,
                        IF(ISNULL(c.标准物流状态), b.物流状态, c.标准物流状态) 物流状态, c.`物流状态代码` 物流状态代码,IF(状态时间='1990-01-01 00:00:00' or 状态时间='1899-12-30 00:00:00' or 状态时间='0000-00-00 00:00:00', '', 状态时间) 状态时间,
                        IF(上线时间='1990-01-01 00:00:00' or 上线时间='1899-12-30 00:00:00' or 上线时间='0000-00-00 00:00:00', '', 上线时间) 上线时间, 系统订单状态, IF(ISNULL(d.订单编号), 系统物流状态, '已退货') 系统物流状态,
                        IF(ISNULL(d.订单编号), NULL, '已退货') 退货登记,
                        IF(ISNULL(d.订单编号), IF(ISNULL(系统物流状态), IF(ISNULL(c.标准物流状态) OR c.标准物流状态 = '未上线', IF(系统订单状态 IN ('已转采购', '待发货'), '未发货', '未上线') , c.标准物流状态), 系统物流状态), '已退货') 最终状态,
                        IF(是否改派='二次改派', '改派', 是否改派) 是否改派,物流方式,物流名称,运输方式,货物类型,是否低价,付款方式,产品id,产品名称,父级分类,
                        二级分类,三级分类,下单时间,审核时间,仓储扫描时间,完结状态时间,价格,价格RMB,价格区间,
                        包裹重量,包裹体积,邮编,IF(ISNULL(b.运单编号), '否', '是') 签收表是否存在,
                        b.订单编号 签收表订单编号, b.运单编号 签收表运单编号, 原运单号, b.物流状态 签收表物流状态, b.添加时间, a.成本价, a.物流花费, a.打包花费, a.其它花费, a.添加物流单号时间,数量
                    FROM {0}_order_list a
                        LEFT JOIN (SELECT * FROM {0} WHERE id IN (SELECT MAX(id) FROM {0} WHERE {0}.添加时间 > '{1}' GROUP BY 运单编号) ORDER BY id) b ON a.`运单编号` = b.`运单编号`
                        LEFT JOIN {0}_logisitis_match c ON b.物流状态 = c.签收表物流状态
                        LEFT JOIN {0}_return d ON a.订单编号 = d.订单编号
                    WHERE a.日期 >= '{2}' AND a.日期 <= '{3}' and (a.订单编号 like 'UP%' or a.订单编号 like 'IG%')
                        AND a.系统订单状态 IN ('已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)','已退货(物流)', '已退货(不拆包物流)')
                    ORDER BY a.`下单时间`;'''.format(team, month_begin, month_last, month_yesterday)
            print('正在获取---' + match[team] + ' ---商城IG和UP订单数据内容…………')
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 商城-{}签收表.xlsx'.format(today, match[team]),
                        sheet_name=match[team], index=False)
            print('----已写入excel')
            filePath = ['D:\\Users\\Administrator\\Desktop\\输出文件\\{} 商城-{}签收表.xlsx'.format(today, match[team])]
            self.e.send('{} 商城-{}签收表.xlsx'.format(today, match[team]), filePath,
                        emailAdd2[team])

    # 物流时效
    def data_wl(self, team):  # 获取各团队近两个月的物流数据
        match = {'slgat': ['台湾', '香港'],
                 'slgat_hfh': ['台湾', '香港'],
                 'slgat_hs': ['台湾', '香港'],
                 'sltg': ['泰国'],
                 'slxmt': ['新加坡', '马来西亚', '菲律宾'],
                 'slxmt_t': ['新加坡', '马来西亚', '菲律宾'],
                 'slxmt_hfh': ['新加坡', '马来西亚', '菲律宾'],
                 'slrb': ['日本'],
                 'slrb_js': ['日本'],
                 'slrb_hs': ['日本'],
                 'slrb_jl': ['日本', '韩国']}
        match1 = {'slgat': ['台湾|神龙家族-港澳台', '香港|神龙家族-港澳台'],
                  'slgat_hfh': ['台湾|火凤凰-港澳台', '香港|火凤凰-港澳台'],
                  'slgat_hs': ['台湾|红杉-港澳台', '香港|红杉-港澳台'],
                  'sltg': ['泰国|神龙家族-泰国'],
                  'slxmt': ['新加坡|神龙家族-新加坡', '马来西亚|神龙家族-马来西亚', '菲律宾|神龙家族-菲律宾'],
                  'slxmt_t': ['新加坡|神龙-T新马菲', '马来西亚|神龙-T新马菲', '菲律宾|神龙-T新马菲'],
                  'slxmt_hfh': ['新加坡|火凤凰-新加坡', '马来西亚|火凤凰-马来西亚', '菲律宾|火凤凰-菲律宾'],
                  'slrb': ['日本|神龙家族-日本团队'],
                  'slrb_js': ['日本|金狮-日本'],
                  'slrb_hs': ['日本|红杉-日本'],
                  'slrb_jl': ['日本|精灵家族-日本', '韩国|精灵家族-韩国', '品牌|精灵家族-品牌']}
        emailAdd = {'台湾': 'giikinliujun@163.com',
                    '香港': 'giikinliujun@163.com',
                    '新加坡': 'zhangjing@giikin.com',
                    '马来西亚': 'zhangjing@giikin.com',
                    '菲律宾': 'zhangjing@giikin.com',
                    '泰国': 'zhangjing@giikin.com',
                    '日本': 'sunyaru@giikin.com',
                    '韩国': 'sunyaru@giikin.com',
                    '品牌': 'sunyaru@giikin.com'}
        if team in ('sltg', 'slrb', 'slrb_jl', 'slgat', 'slgat_hfh', 'slgat_hs', 'slxmt', 'slxmt_t', 'slxmt_hfh'):
            month_last = (datetime.datetime.now().replace(day=1)).strftime('%Y-%m-%d')
        else:
            pass
        for tem in match1[team]:
            tem1 = tem.split('|')[0]
            tem2 = tem.split('|')[1]
            filePath = []
            listT = []  # 查询sql的结果 存放池
            print('正在获取---' + tem2 + '---物流时效…………')
            # 总月
            sql = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IF(下单出库时 = 90,null,IFNULL(下单出库时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            IF(ISNULL(DATEDIFF(仓储扫描时间, 下单时间)), 90, DATEDIFF(仓储扫描时间, 下单时间))  AS 下单出库时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			            FROM  d1_{0} cx 
			            WHERE cx.`币种` = '{1}'	AND cx.`团队` = '{2}'
				            AND cx.`是否改派` = '直发' 
				            AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				        GROUP BY 年月,币种,物流方式,下单出库时 
				        ORDER BY 年月,币种,物流方式,下单出库时 
		                ) sl
		            GROUP BY 年月,币种,物流方式,下单出库时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);;'''.format(team, tem1, tem2)
            df = pd.read_sql_query(sql=sql, con=self.engine1)
            listT.append(df)
            sql2 = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IF(出库完成时 = 90,null,IFNULL(出库完成时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            IF(ISNULL(DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`仓储扫描时间`)), 90, DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`仓储扫描时间`))  AS 出库完成时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			                FROM  d1_{0} cx 
			                WHERE cx.`币种` = '{1}' AND cx.`团队` = '{2}' 
				                AND cx.`是否改派` = '直发' 
				                AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				            GROUP BY 年月,币种,物流方式,出库完成时 
				            ORDER BY 年月,币种,物流方式,出库完成时 
		                ) sl
		            GROUP BY 年月,币种,物流方式,出库完成时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);'''.format(team, tem1, tem2)
            df2 = pd.read_sql_query(sql=sql2, con=self.engine1)
            listT.append(df2)
            sql3 = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IF(下单完成时 = 90,null,IFNULL(下单完成时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            IF(ISNULL(DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`)), 90, DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`))  AS 下单完成时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			            FROM  d1_{0} cx 
			            WHERE cx.`币种` = '{1}' AND cx.`团队` = '{2}' 
				            AND cx.`是否改派` = '直发' 
				            AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				        GROUP BY 年月,币种,物流方式,下单完成时 
				        ORDER BY 年月,币种,物流方式,下单完成时 
		            ) sl
		            GROUP BY 年月,币种,物流方式,下单完成时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);'''.format(team, tem1, tem2)
            df3 = pd.read_sql_query(sql=sql3, con=self.engine1)
            listT.append(df3)
            sql4 = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IF(下单完成时 = 90,null,IFNULL(下单完成时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            IF(ISNULL(DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`)), 90, DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`))  AS 下单完成时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			            FROM  d1_{0} cx
			            WHERE cx.`币种` = '{1}' AND cx.`团队` = '{2}' 
				            AND cx.`是否改派` = '改派' 
				            AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				        GROUP BY 年月,币种,物流方式,下单完成时 
				        ORDER BY 年月,币种,物流方式,下单完成时 
		            ) sl
		            GROUP BY 年月,币种,物流方式,下单完成时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);'''.format(team, tem1, tem2)
            df4 = pd.read_sql_query(sql=sql4, con=self.engine1)
            listT.append(df4)
            # 分旬
            print('正在获取---' + tem + '---物流分旬时效…………')
            sql10 = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IFNULL(旬,'总计') as 旬,
			                IF(下单出库时 = 90,null,IFNULL(下单出库时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            旬,
					            IF(ISNULL(DATEDIFF(仓储扫描时间, 下单时间)), 90, DATEDIFF(仓储扫描时间, 下单时间))  AS 下单出库时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			            FROM  d1_{0} cx 
			            WHERE cx.`币种` = '{1}' AND cx.`团队` = '{2}' 
				            AND cx.`是否改派` = '直发' 
				            AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				        GROUP BY 年月,币种,物流方式,旬,下单出库时 
				        ORDER BY 年月,币种,物流方式,旬,下单出库时 
		            ) sl
		            GROUP BY 年月,币种,物流方式,旬,下单出库时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);'''.format(team, tem1, tem2)
            df10 = pd.read_sql_query(sql=sql10, con=self.engine1)
            listT.append(df10)
            sql20 = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IFNULL(旬,'总计') as 旬,
			                IF(出库完成时 = 90,null,IFNULL(出库完成时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            旬,
					            IF(ISNULL(DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`仓储扫描时间`)), 90, DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`仓储扫描时间`))  AS 出库完成时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			            FROM  d1_{0} cx 
			            WHERE cx.`币种` = '{1}' AND cx.`团队` = '{2}' 
				            AND cx.`是否改派` = '直发' 
				            AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				            GROUP BY 年月,币种,物流方式,旬,出库完成时 
				        ORDER BY 年月,币种,物流方式,旬,出库完成时 
		            ) sl
		            GROUP BY 年月,币种,物流方式,旬,出库完成时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);'''.format(team, tem1, tem2)
            df20 = pd.read_sql_query(sql=sql20, con=self.engine1)
            listT.append(df20)
            sql30 = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IFNULL(旬,'总计') as 旬,
			                IF(下单完成时 = 90,null,IFNULL(下单完成时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            旬,
					            IF(ISNULL(DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`)), 90, DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`))  AS 下单完成时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			            FROM  d1_{0} cx 
			            WHERE cx.`币种` = '{1}' AND cx.`团队` = '{2}' 
				            AND cx.`是否改派` = '直发' 
				            AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				        GROUP BY 年月,币种,物流方式,旬,下单完成时 
				        ORDER BY 年月,币种,物流方式,旬,下单完成时 
		            ) sl
		            GROUP BY 年月,币种,物流方式,旬,下单完成时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);'''.format(team, tem1, tem2)
            df30 = pd.read_sql_query(sql=sql30, con=self.engine1)
            listT.append(df30)
            sql40 = '''SELECT 年月,
			                IFNULL(币种,'总计') as 币种,
			                IFNULL(物流方式,'总计') as 物流方式,
			                IFNULL(旬,'总计') as 旬,
			                IF(下单完成时 = 90,null,IFNULL(下单完成时,'总计')) as 天数,
			                SUM(订单量) as 总计,
			                SUM(签收量) as 签收量,
			                SUM(完成量) as 完成量,				  
			                SUM(签收量) / SUM(完成量) AS '签收率完成',
			                SUM(签收量) / SUM(订单量)  AS '签收率总计',
			                null 累计完成占比
                    FROM( SELECT  年月,
					            币种,
					            物流方式,
					            旬,
					            IF(ISNULL(DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`)), 90, DATEDIFF(IFNULL(`完结状态时间`,`状态时间`),`下单时间`))  AS 下单完成时,
					            COUNT(订单编号) AS 订单量, 
					            SUM(最终状态 = '已签收') as 签收量, 
					            SUM(最终状态 in ('已签收','拒收','理赔','已退货')) as 完成量
			            FROM  d1_{0} cx 
			            WHERE cx.`币种` = '{1}' AND cx.`团队` = '{2}' 
				            AND cx.`是否改派` = '改派' 
				            AND cx.系统订单状态 IN ( '已审核', '已转采购', '已发货', '已收货', '已完成', '已退货(销售)', '已退货(物流)', '已退货(不拆包物流)' ) 
				        GROUP BY 年月,币种,物流方式,旬,下单完成时 
				        ORDER BY 年月,币种,物流方式,旬,下单完成时 
		            ) sl
		            GROUP BY 年月,币种,物流方式,旬,下单完成时 
			        with rollup
			        HAVING (`币种` IS NOT null  AND `年月` IS NOT null);'''.format(team, tem1, tem2)
            df40 = pd.read_sql_query(sql=sql40, con=self.engine1)
            listT.append(df40)
            print('正在写入excel…………')
            today = datetime.date.today().strftime('%Y.%m.%d')
            if team in ('slgat_hfh', 'slxmt_hfh'):
                file_path = 'D:\\Users\\Administrator\\Desktop\\输出文件\\{} 火凤凰-{}物流时效.xlsx'.format(today, tem1)
            elif team in ('slgat_hs', 'slrb_hs'):
                file_path = 'D:\\Users\\Administrator\\Desktop\\输出文件\\{} 红杉-{}物流时效.xlsx'.format(today, tem1)
            elif team == 'slxmt_t':
                file_path = 'D:\\Users\\Administrator\\Desktop\\输出文件\\{} 神龙T-{}物流时效.xlsx'.format(today, tem1)
            elif team == 'slrb_js':
                file_path = 'D:\\Users\\Administrator\\Desktop\\输出文件\\{} 金狮-{}物流时效.xlsx'.format(today, tem1)
            elif team == 'slrb_jl':
                file_path = 'D:\\Users\\Administrator\\Desktop\\输出文件\\{} 精灵-{}物流时效.xlsx'.format(today, tem1)
            else:
                file_path = 'D:\\Users\\Administrator\\Desktop\\输出文件\\{} 神龙-{}物流时效.xlsx'.format(today, tem1)
            sheet_name = ['下单出库时', '出库完成时', '下单完成时', '改派下单完成时', '下单出库(分旬)', '出库完成(分旬)', '下单完成(分旬)', '改派下单完成(分旬)']
            df0 = pd.DataFrame([])                       # 创建空的dataframe数据框
            df0.to_excel(file_path, index=False)         # 备用：可以向不同的sheet写入数据（创建新的工作表并进行写入）
            writer = pd.ExcelWriter(file_path, engine='openpyxl')  # 初始化写入对象
            book = load_workbook(file_path)              # 可以向不同的sheet写入数据（对现有工作表的追加）
            writer.book = book                           # 将数据写入excel中的sheet2表,sheet_name改变后即是新增一个sheet
            for i in range(len(listT)):
                listT[i]['签收率完成'] = listT[i]['签收率完成'].fillna(value=0)
                listT[i]['签收率总计'] = listT[i]['签收率总计'].fillna(value=0)
                listT[i]['签收率完成'] = listT[i]['签收率完成'].apply(lambda x: format(x, '.2%'))
                listT[i]['签收率总计'] = listT[i]['签收率总计'].apply(lambda x: format(x, '.2%'))
                listT[i].to_excel(excel_writer=writer, sheet_name=sheet_name[i], index=False)
            if 'Sheet1' in book.sheetnames:              # 删除新建文档时的第一个工作表
                del book['Sheet1']
            writer.save()
            writer.close()
            print('正在运行宏…………')
            app = xl.App(visible=False, add_book=False)  # 运行宏调整
            app.display_alerts = False
            wbsht = app.books.open('D:/Users/Administrator/Desktop/新版-格式转换(工具表).xlsm')
            wbsht1 = app.books.open(file_path)
            wbsht.macro('sltem物流时效')()
            wbsht1.save()
            wbsht1.close()
            wbsht.close()
            app.quit()
            print('----已写入excel ')
            filePath.append(file_path)
            if team in ('slgat_hfh', 'slxmt_hfh'):
                self.e.send('{} 火凤凰-{}物流时效.xlsx'.format(today, tem1), filePath,
                            emailAdd[tem1])
            elif team in ('slgat_hs', 'slrb_hs'):
                self.e.send('{} 红杉-{}物流时效.xlsx'.format(today, tem1), filePath,
                            emailAdd[tem1])
            elif team == 'slxmt_t':
                self.e.send('{} 神龙T-{}物流时效.xlsx'.format(today, tem1), filePath,
                            emailAdd[tem1])
            elif team == 'slrb_js':
                self.e.send('{} 金狮-{}物流时效.xlsx'.format(today, tem1), filePath,
                            emailAdd[tem1])
            elif team == 'slrb_jl':
                self.e.send('{} 精灵-{}物流时效.xlsx'.format(today, tem1), filePath,
                            emailAdd[tem1])
            elif team == 'sltg':
                print('---' + tem1 + ' 不发送邮件')
            else:
                self.e.send('{} 神龙-{}物流时效.xlsx'.format(today, tem1), filePath,
                            emailAdd[tem1])

    # 无运单号查询
    def noWaybillNumber(self, team):
        match1 = {'slgat': '神龙-港台',
                  'slgat_hfh': '火凤凰-港台',
                  'sltg': '神龙-泰国',
                  'slxmt': '神龙-新马',
                  'slxmt_t': '神龙T-新马',
                  'slxmt_hfh': '火凤凰-新马',
                  'slrb': '神龙-日本',
                  'slrb_jl': '精灵-日本'}
        match = {'slgat': '"神龙家族-港澳台"',
                 'slgat_hfh': '"火凤凰-港澳台"',
                 'sltg': '"神龙家族-泰国"',
                 'slxmt': '"神龙家族-新加坡", "神龙家族-马来西亚", "神龙家族-菲律宾"',
                 'slxmt_t': '"神龙-T新马菲"',
                 'slxmt_hfh': '"火凤凰-新加坡", "火凤凰-马来西亚", "火凤凰-菲律宾"',
                 'slrb': '"神龙家族-日本团队"',
                 'slrb_jl': '"精灵家族-日本"'}
        emailAdd = {'slgat': 'giikinliujun@163.com',
                    'slgat_hfh': 'giikinliujun@163.com',
                    'sltg': 'zhangjing@giikin.com',
                    'slxmt': 'zhangjing@giikin.com',
                    'slxmt_t': 'zhangjing@giikin.com',
                    'slxmt_hfh': 'zhangjing@giikin.com',
                    'slrb': 'sunyaru@giikin.com',
                    'slrb_jl': 'sunyaru@giikin.com'}
        emailAdd2 = {'sltg': 'libin@giikin.com'}
        print('正在查询{}无运单订单列表…………'.format(match[team]))
        yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
        last_month = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
        sql = '''SELECT a.rq 日期,
                        dim_area.name 团队,
                        a.order_number 订单编号,
                        a.waybill_number 运单编号,
                        a.order_status 系统订单状态id,
                        IF(a.second=0,'直发','改派') 是否改派,
                        dim_trans_way.all_name 物流方式,
                        a.addtime 下单时间,
                        a.verity_time 审核时间
                FROM gk_order a
                    left join dim_area ON dim_area.id = a.area_id
                    left join dim_trans_way on dim_trans_way.id = a.logistics_id
                WHERE
                    a.rq >= '{}' AND a.rq <= '{}'
                    AND dim_area.name IN ({})
                    AND ISNULL(waybill_number)
                    AND order_status NOT IN (1, 8, 11, 14, 16)
                ORDER BY a.rq;'''.format(last_month, yesterday, match[team])
        df = pd.read_sql_query(sql=sql, con=self.engine2)
        sql = 'SELECT * FROM dim_order_status;'
        df1 = pd.read_sql_query(sql=sql, con=self.engine1)
        print('正在合并订单状态…………')
        df = pd.merge(left=df, right=df1, left_on='系统订单状态id', right_on='id', how='left')
        df = df.drop(labels=['id', '系统订单状态id', ], axis=1)
        df.rename(columns={'name': '系统订单状态'}, inplace=True)
        today = datetime.date.today().strftime('%Y.%m.%d')
        df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}无运单号列表.xlsx'.format(today, match1[team]),
                    sheet_name=match1[team], index=False)
        filePath = ['D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}无运单号列表.xlsx'.format(today, match1[team])]
        print('输出文件成功…………')
        self.e.send('{} {}无运单号列表.xlsx'.format(today, match1[team]), filePath,
                    emailAdd[team])
        if team == 'sltg':
            self.e.send('{} {}无运单号列表.xlsx'.format(today, match1[team]), filePath,
                        emailAdd2[team])

    # 产品花费表
    def orderCost(self, team):
        if datetime.datetime.now().day >= 9:
            endDate = datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')
            startDate = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            endDate = [endDate, datetime.datetime.now().strftime('%Y-%m-%d')]
            startDate = [startDate, datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')]
        else:
            endDate = datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')
            startDate = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            endDate = [endDate]
            startDate = [startDate]
        match = {'SG': '新加坡',
                 'MY': '马来西亚',
                 'PH': '菲律宾',
                 'JP': '日本',
                 'HK': '香港',
                 'TW': '台湾',
                 'TH': '泰国'}
        match2 = {'SG': 'slxmt_zqsb',
                  'MY': 'slxmt_zqsb',
                  'PH': 'slxmt_zqsb',
                  'JP': 'slrb_zqsb_rb',
                  'HK': 'slgat_zqsb',
                  'TW': 'slgat_zqsb',
                  'TH': 'sltg_zqsb'}
        emailAdd = {'SG': 'zhangjing@giikin.com',
                    'MY': 'zhangjing@giikin.com',
                    'PH': 'zhangjing@giikin.com',
                    'JP': 'sunyaru@giikin.com',
                    'HK': 'giikinliujun@163.com',
                    'TW': 'giikinliujun@163.com',
                    'TH': 'zhangjing@giikin.com'}
        # filePath = []
        for i in range(len(endDate)):
            print('正在查询' + match[team] + '产品花费表…………')
            sql = '''SELECT s1.`month` AS '月份',
                         s1.area AS '地区',
                         s1.leader AS '负责人',
                         s1.pid AS '产品ID',
                         s1.pname AS '产品名称',
                         s1.cate1 AS '一级品类',
                         s1.cate2 AS '二级品类',
                         s1.cate3 AS '三级品类',
                         s1.orders AS '订单量',
                         s1.orders - s1.gps AS '直发订单量',
                         s1.gps AS '改派订单量',
                         s1.salesRMB / s1.orders AS '客单价',
                         s1.salesRMB / s1.adcost AS 'ROI',
                         s1.orders / s2.orders AS '订单品类占比',
                         s1.cgcost_zf / s1.salesRMB AS '直发采购/销售额',
                         s1.adcost / s1.salesRMB AS '花费占比',
                         s1.wlcost / s1.salesRMB AS '运费占比',
                         s1.qtcost / s1.salesRMB AS '手续费占比',
                         ( s1.cgcost_zf + s1.adcost + s1.wlcost + s1.qtcost ) / s1.salesRMB AS '总成本占比',
                         s1.salesRMB_yqs / (s1.salesRMB_yqs + s1.salesRMB_yjs) AS '金额签收/完成',
                         s1.salesRMB_yqs / s1.salesRMB AS '金额签收/总计',
                         (s1.salesRMB_yqs + s1.salesRMB_yjs) / s1.salesRMB AS '金额完成占比',
                         s1.yqs / (s1.yqs + s1.yjs) AS '数量签收/完成',
                         (s1.yqs + s1.yjs) / s1.orders AS '数量完成占比',
                         s3.orders AS '昨日订单量'
            FROM (  SELECT DISTINCT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     a.product_id AS pid,
                                     e.`product_name` AS pname,
            --                       e.`name` AS pname,
                                     d.ppname AS cate1,
                                     d.pname AS cate2,
                                     d.`name` AS cate3,
            -- 						 GROUP_CONCAT(DISTINCT a.low_price) AS low_price,
                                     SUM(a.orders) AS orders,
                                     SUM(a.yqs) AS yqs,
                                     SUM(a.yjs) AS yjs,
                                     SUM(a.salesRMB) AS salesRMB,
                                     SUM(a.salesRMB_yqs) AS salesRMB_yqs,
                                     SUM(a.salesRMB_yjs) AS salesRMB_yjs,
                                     SUM(a.gps) AS gps,
            --                       SUM(a.cgcost_zf) AS cgcost_zf,
            --                       SUM(a.adcost) AS adcost,
                                     null cgcost_zf,
                                     null adcost,
                                     SUM(a.wlcost) AS wlcost,
            --                       SUM(a.qtcost) AS qtcost
                                     null qtcost
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                            LEFT JOIN dim_cate d on d.id = a.third_cate_id
                          	LEFT JOIN gk_sale e ON e.id = a.sale_id
                        WHERE a.rq >= '{startDate}'
                            AND a.rq < '{endDate}'
                            AND b.pcode = '{team}'
                            AND c.uname = '王冰'
                            AND a.beform <> 'mf'
                            AND c.uid <> 10099  -- 过滤翼虎
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, a.product_id
                     ) s1
                      LEFT JOIN
                     (  SELECT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     d.ppname AS cate1,
                                     SUM(a.orders) AS orders
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                            LEFT JOIN dim_cate d on d.id = a.third_cate_id
                        WHERE a.rq >= '{startDate}'
                            AND a.rq < '{endDate}'
                            AND b.pcode = '{team}'
                            AND c.uname = '王冰'
                            AND a.beform <> 'mf'
                            AND c.uid <> 10099
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, d.ppname
                     ) s2  ON s1.`month`=s2.`month` AND s1.area=s2.area AND s1.leader=s2.leader AND s1.cate1=s2.cate1
                      LEFT JOIN
                     (  SELECT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     a.product_id AS pid,
                                     SUM(a.orders) AS orders
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                        WHERE a.rq = DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                                    AND b.pcode = '{team}'
                                    AND c.uname = '王冰'
                                    AND a.beform <> 'mf'
                                    AND c.uid <> 10099
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, a.product_id
                     ) s3 ON s1.area=s3.area AND s1.leader=s3.leader AND s1.pid=s3.pid
            WHERE s1.orders > 0
            ORDER BY s1.orders DESC'''.format(team=team, startDate=startDate[i], endDate=endDate[i])
            df = pd.read_sql_query(sql=sql, con=self.engine2)
            print('正在输出' + match[team] + '产品花费表…………')
            columns = ['订单品类占比', '直发采购/销售额', '花费占比', '运费占比', '手续费占比', '总成本占比',
                       '金额签收/完成', '金额签收/总计', '金额完成占比', '数量签收/完成', '数量完成占比']
            for column in columns:
                df[column] = df[column].fillna(value=0)
                df[column] = df[column].apply(lambda x: format(x, '.2%'))
            today = datetime.date.today().strftime('%Y.%m.%d')
            if i == 0:
                df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}上月产品花费表.xlsx'.format(today, match[team]),
                            sheet_name=match[team], index=False)
                # filePath.append('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}上月产品花费表.xlsx'.format(today, match[team]))
            else:
                df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}本月产品花费表.xlsx'.format(today, match[team]),
                            sheet_name=match[team], index=False)
                # filePath.append('D:\\Users\\Administrator\\Desktop\\输出文件\\{} {}本月产品花费表.xlsx'.format(today, match[team]))
        # self.e.send(match[team] + '产品花费表', filePath,
        #             emailAdd[team])
        self.d.sl_tem_cost(match2[team], match[team])
    def orderCostHFH(self, team):
        if datetime.datetime.now().day >= 9:
            endDate = datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')
            startDate = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            endDate = [endDate, datetime.datetime.now().strftime('%Y-%m-%d')]
            startDate = [startDate, datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')]
        else:
            endDate = datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')
            startDate = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            endDate = [endDate]
            startDate = [startDate]
        match = {'SG': '新加坡',
                 'MY': '马来西亚',
                 'PH': '菲律宾',
                 'JP': '日本',
                 'HK': '香港',
                 'TW': '台湾',
                 'TH': '泰国'}
        match2 = {'SG': 'slxmt_hfh_zqsb',
                  'MY': 'slxmt_hfh_zqsb',
                  'PH': 'slxmt_hfh_zqsb',
                  'JP': 'slrb_zqsb_rb',
                  'HK': 'slgat_hfh_zqsb',
                  'TW': 'slgat_hfh_zqsb',
                  'TH': 'sltg_zqsb'}
        emailAdd = {'SG': 'zhangjing@giikin.com',
                    'MY': 'zhangjing@giikin.com',
                    'PH': 'zhangjing@giikin.com',
                    'JP': 'sunyaru@giikin.com',
                    'HK': 'giikinliujun@163.com',
                    'TW': 'giikinliujun@163.com',
                    'TH': 'zhangjing@giikin.com'}
        # filePath = []
        for i in range(len(endDate)):
            print('正在查询' + match[team] + '产品花费表…………')
            sql = '''SELECT s1.`month` AS '月份',
                         s1.area AS '地区',
                         s1.leader AS '负责人',
                         s1.pid AS '产品ID',
                         s1.pname AS '产品名称',
                         s1.cate1 AS '一级品类',
                         s1.cate2 AS '二级品类',
                         s1.cate3 AS '三级品类',
                         s1.orders AS '订单量',
                         s1.orders - s1.gps AS '直发订单量',
                         s1.gps AS '改派订单量',
                         s1.salesRMB / s1.orders AS '客单价',
                         s1.salesRMB / s1.adcost AS 'ROI',
                         s1.orders / s2.orders AS '订单品类占比',
                         s1.cgcost_zf / s1.salesRMB AS '直发采购/销售额',
                         s1.adcost / s1.salesRMB AS '花费占比',
                         s1.wlcost / s1.salesRMB AS '运费占比',
                         s1.qtcost / s1.salesRMB AS '手续费占比',
                         ( s1.cgcost_zf + s1.adcost + s1.wlcost + s1.qtcost ) / s1.salesRMB AS '总成本占比',
                         s1.salesRMB_yqs / (s1.salesRMB_yqs + s1.salesRMB_yjs) AS '金额签收/完成',
                         s1.salesRMB_yqs / s1.salesRMB AS '金额签收/总计',
                         (s1.salesRMB_yqs + s1.salesRMB_yjs) / s1.salesRMB AS '金额完成占比',
                         s1.yqs / (s1.yqs + s1.yjs) AS '数量签收/完成',
                         (s1.yqs + s1.yjs) / s1.orders AS '数量完成占比',
                         s3.orders AS '昨日订单量'
            FROM (  SELECT DISTINCT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     a.product_id AS pid,
                                     e.`product_name` AS pname,
            --                       e.`name` AS pname,
                                     d.ppname AS cate1,
                                     d.pname AS cate2,
                                     d.`name` AS cate3,
            -- 						 GROUP_CONCAT(DISTINCT a.low_price) AS low_price,
                                     SUM(a.orders) AS orders,
                                     SUM(a.yqs) AS yqs,
                                     SUM(a.yjs) AS yjs,
                                     SUM(a.salesRMB) AS salesRMB,
                                     SUM(a.salesRMB_yqs) AS salesRMB_yqs,
                                     SUM(a.salesRMB_yjs) AS salesRMB_yjs,
                                     SUM(a.gps) AS gps,
            --                       SUM(a.cgcost_zf) AS cgcost_zf,
            --                       SUM(a.adcost) AS adcost,
                                     null cgcost_zf,
                                     null adcost,
                                     SUM(a.wlcost) AS wlcost,
            --                       SUM(a.qtcost) AS qtcost
                                     null qtcost
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                            LEFT JOIN dim_cate d on d.id = a.third_cate_id
                            LEFT JOIN gk_sale e ON e.id = a.sale_id
                        WHERE a.rq >= '{startDate}'
                            AND a.rq < '{endDate}'
                            AND b.pcode = '{team}'
                            AND c.uname = '罗超源'
                            AND a.beform <> 'mf'
                            AND c.uid <> 10099  -- 过滤翼虎
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, a.product_id
                     ) s1
                      LEFT JOIN
                     (  SELECT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     d.ppname AS cate1,
                                     SUM(a.orders) AS orders
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                            LEFT JOIN dim_cate d on d.id = a.third_cate_id
                        WHERE a.rq >= '{startDate}'
                            AND a.rq < '{endDate}'
                            AND b.pcode = '{team}'
                            AND c.uname = '罗超源'
                            AND a.beform <> 'mf'
                            AND c.uid <> 10099
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, d.ppname
                     ) s2  ON s1.`month`=s2.`month` AND s1.area=s2.area AND s1.leader=s2.leader AND s1.cate1=s2.cate1
                      LEFT JOIN
                     (  SELECT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     a.product_id AS pid,
                                     SUM(a.orders) AS orders
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                        WHERE a.rq = DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                                    AND b.pcode = '{team}'
                                    AND c.uname = '罗超源'
                                    AND a.beform <> 'mf'
                                    AND c.uid <> 10099
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, a.product_id
                     ) s3 ON s1.area=s3.area AND s1.leader=s3.leader AND s1.pid=s3.pid
            WHERE s1.orders > 0
            ORDER BY s1.orders DESC'''.format(team=team, startDate=startDate[i], endDate=endDate[i])
            df = pd.read_sql_query(sql=sql, con=self.engine2)
            print('正在输出' + match[team] + '产品花费表…………')
            columns = ['订单品类占比', '直发采购/销售额', '花费占比', '运费占比', '手续费占比', '总成本占比',
                       '金额签收/完成', '金额签收/总计', '金额完成占比', '数量签收/完成', '数量完成占比']
            for column in columns:
                df[column] = df[column].fillna(value=0)
                df[column] = df[column].apply(lambda x: format(x, '.2%'))
            today = datetime.date.today().strftime('%Y.%m.%d')
            if i == 0:
                df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 火凤凰-{}上月产品花费表.xlsx'.format(today, match[team]),
                            sheet_name=match[team], index=False)
                # filePath.append('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 火凤凰{}上月产品花费表.xlsx'.format(today, match[team]))
            else:
                df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 火凤凰-{}本月产品花费表.xlsx'.format(today, match[team]),
                            sheet_name=match[team], index=False)
                # filePath.append('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 火凤凰{}本月产品花费表.xlsx'.format(today, match[team]))
        # self.e.send(match[team] + '产品花费表', filePath,
        #             emailAdd[team])
        self.d.sl_tem_costHFH(match2[team], match[team])
    def orderCostT(self, team):
        if datetime.datetime.now().day >= 9:
            endDate = datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')
            startDate = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            endDate = [endDate, datetime.datetime.now().strftime('%Y-%m-%d')]
            startDate = [startDate, datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')]
        else:
            endDate = datetime.datetime.now().replace(day=1).strftime('%Y-%m-%d')
            startDate = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
            endDate = [endDate]
            startDate = [startDate]
        match = {'SG': '新加坡',
                 'MY': '马来西亚',
                 'PH': '菲律宾'}
        match2 = {'SG': 'slxmt_t_zqsb',
                  'MY': 'slxmt_t_zqsb',
                  'PH': 'slxmt_t_zqsb'}
        emailAdd = {'SG': 'zhangjing@giikin.com',
                    'MY': 'zhangjing@giikin.com',
                    'PH': 'zhangjing@giikin.com'}
        # filePath = []
        for i in range(len(endDate)):
            print('正在查询' + match[team] + '产品花费表…………')
            sql = '''SELECT s1.`month` AS '月份',
                         s1.area AS '地区',
                         s1.leader AS '负责人',
                         s1.pid AS '产品ID',
                         s1.pname AS '产品名称',
                         s1.cate1 AS '一级品类',
                         s1.cate2 AS '二级品类',
                         s1.cate3 AS '三级品类',
                         s1.orders AS '订单量',
                         s1.orders - s1.gps AS '直发订单量',
                         s1.gps AS '改派订单量',
                         s1.salesRMB / s1.orders AS '客单价',
                         s1.salesRMB / s1.adcost AS 'ROI',
                         s1.orders / s2.orders AS '订单品类占比',
                         s1.cgcost_zf / s1.salesRMB AS '直发采购/销售额',
                         s1.adcost / s1.salesRMB AS '花费占比',
                         s1.wlcost / s1.salesRMB AS '运费占比',
                         s1.qtcost / s1.salesRMB AS '手续费占比',
                         ( s1.cgcost_zf + s1.adcost + s1.wlcost + s1.qtcost ) / s1.salesRMB AS '总成本占比',
                         s1.salesRMB_yqs / (s1.salesRMB_yqs + s1.salesRMB_yjs) AS '金额签收/完成',
                         s1.salesRMB_yqs / s1.salesRMB AS '金额签收/总计',
                         (s1.salesRMB_yqs + s1.salesRMB_yjs) / s1.salesRMB AS '金额完成占比',
                         s1.yqs / (s1.yqs + s1.yjs) AS '数量签收/完成',
                         (s1.yqs + s1.yjs) / s1.orders AS '数量完成占比',
                         s3.orders AS '昨日订单量'
            FROM (  SELECT DISTINCT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     a.product_id AS pid,
                                     e.`product_name` AS pname,
            --                       e.`name` AS pname,
                                     d.ppname AS cate1,
                                     d.pname AS cate2,
                                     d.`name` AS cate3,
            -- 						 GROUP_CONCAT(DISTINCT a.low_price) AS low_price,
                                     SUM(a.orders) AS orders,
                                     SUM(a.yqs) AS yqs,
                                     SUM(a.yjs) AS yjs,
                                     SUM(a.salesRMB) AS salesRMB,
                                     SUM(a.salesRMB_yqs) AS salesRMB_yqs,
                                     SUM(a.salesRMB_yjs) AS salesRMB_yjs,
                                     SUM(a.gps) AS gps,
            --                       SUM(a.cgcost_zf) AS cgcost_zf,
            --                       SUM(a.adcost) AS adcost,
                                     null cgcost_zf,
                                     null adcost,
                                     SUM(a.wlcost) AS wlcost,
            --                       SUM(a.qtcost) AS qtcost
                                     null qtcost
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                            LEFT JOIN dim_cate d on d.id = a.third_cate_id
            --              LEFT JOIN gk_product e on e.id = a.product_id
                          LEFT JOIN (SELECT * FROM gk_sale WHERE id IN (SELECT MAX(id) FROM gk_sale GROUP BY product_id ) ORDER BY id) e on e.product_id = a.product_id
                        WHERE a.rq >= '{startDate}'
                            AND a.rq < '{endDate}'
                            AND b.pcode = '{team}'
                            AND c.uname = '王冰'
                            AND a.beform <> 'mf'
                            AND c.uid <> 10099  -- 过滤翼虎
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, a.product_id
                     ) s1
                      LEFT JOIN
                     (  SELECT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     d.ppname AS cate1,
                                     SUM(a.orders) AS orders
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                            LEFT JOIN dim_cate d on d.id = a.third_cate_id
                        WHERE a.rq >= '{startDate}'
                            AND a.rq < '{endDate}'
                            AND b.pcode = '{team}'
                            AND c.uname = '王冰'
                            AND a.beform <> 'mf'
                            AND c.uid <> 10099
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, d.ppname
                     ) s2  ON s1.`month`=s2.`month` AND s1.area=s2.area AND s1.leader=s2.leader AND s1.cate1=s2.cate1
                      LEFT JOIN
                     (  SELECT EXTRACT(YEAR_MONTH FROM a.rq) AS `month`,
                                     b.pname AS area,
                                     c.uname AS leader,
                                     a.product_id AS pid,
                                     SUM(a.orders) AS orders
                        FROM gk_order_day a
                            LEFT JOIN dim_currency_lang b ON a.currency_lang_id = b.id
                            LEFT JOIN dim_area c on c.id = a.area_id
                        WHERE a.rq = DATE_SUB(CURDATE(), INTERVAL 1 DAY)
                                    AND b.pcode = '{team}'
                                    AND c.uname = '王冰'
                                    AND a.beform <> 'mf'
                                    AND c.uid <> 10099
                        GROUP BY EXTRACT(YEAR_MONTH FROM a.rq), b.pname, c.uname, a.product_id
                     ) s3 ON s1.area=s3.area AND s1.leader=s3.leader AND s1.pid=s3.pid
            WHERE s1.orders > 0
            ORDER BY s1.orders DESC'''.format(team=team, startDate=startDate[i], endDate=endDate[i])
            df = pd.read_sql_query(sql=sql, con=self.engine2)
            print('正在输出' + match[team] + '产品花费表…………')
            columns = ['订单品类占比', '直发采购/销售额', '花费占比', '运费占比', '手续费占比', '总成本占比',
                       '金额签收/完成', '金额签收/总计', '金额完成占比', '数量签收/完成', '数量完成占比']
            for column in columns:
                df[column] = df[column].fillna(value=0)
                df[column] = df[column].apply(lambda x: format(x, '.2%'))
            today = datetime.date.today().strftime('%Y.%m.%d')
            if i == 0:
                df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 神龙T-{}上月产品花费表.xlsx'.format(today, match[team]),
                            sheet_name=match[team], index=False)
                # filePath.append('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 火凤凰{}上月产品花费表.xlsx'.format(today, match[team]))
            else:
                df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 神龙T-{}本月产品花费表.xlsx'.format(today, match[team]),
                            sheet_name=match[team], index=False)
                # filePath.append('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 火凤凰{}本月产品花费表.xlsx'.format(today, match[team]))
        # self.e.send(match[team] + '产品花费表', filePath,
        #             emailAdd[team])
        self.d.sl_tem_costT(match2[team], match[team])


if __name__ == '__main__':
    #  messagebox.showinfo("提示！！！", "当前查询已完成--->>> 请前往（ 输出文件 ）查看")200
    m = MysqlControl()
    start = datetime.datetime.now()

    # 更新产品id的列表
    m.update_gk_product()

    for team in ['slrb', 'slxmt', 'slxmt_t', 'slxmt_hfh']:  # 无运单号查询200
        m.noWaybillNumber(team)

    match = {'SG': '新加坡',
             'MY': '马来西亚',
             'PH': '菲律宾',
             'JP': '日本'}
    # match = {'HK': '香港',
    #          'TW': '台湾'}
    for team in match.keys():  # 产品花费表200
        if team == 'JP':
            m.orderCost(team)
        elif team in ('HK', 'TW'):
            m.orderCost(team)
            m.orderCostHFH(team)
        else:
            m.orderCost(team)
            m.orderCostHFH(team)
            m.orderCostT(team)


    # sm = SltemMonitoring()  # 成本查询
    # for team in ['菲律宾', '新加坡', '马来西亚', '日本', '香港', '台湾']:
    #     sm.costWaybill(team)

    # 测试物流时效
    # team = 'sltg'
    # m.data_wl(team)
    # for team in ['slgat', 'slgat_hfh', 'slrb', 'slrb_jl', 'sltg', 'slxmt', 'slxmt_hfh']:
    # for team in ['slgat']:
    #     m.data_wl(team)



