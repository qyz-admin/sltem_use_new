from openpyxl import Workbook, load_workbook
from excelControl import ExcelControl
from mysqlControl import MysqlControl
from settings import Settings
from sqlalchemy import create_engine
import datetime
import pandas as pd
import os
import xlwings
import numpy as np

class orderControl(Settings):
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
    def readsheet(self):
        start: datetime = datetime.datetime.now()
        team = 'sltem'
        match = {'sltem': r'D:\Users\Administrator\Desktop\需要用到的文件\A查询导表'}
        path = match[team]
        dirs = os.listdir(path=path)
        # ---读取execl文件---
        for dir in dirs:
            filePath = os.path.join(path, dir)
            print(filePath)
            if dir[:2] != '~$':
                fileType = os.path.splitext(filePath)[1]
                app = xlwings.App(visible=False, add_book=False)
                app.display_alerts = False        # 不显示excel窗口
                if 'xls' in fileType:
                    wb = app.books.open(filePath, update_links=False, read_only=True)
                    for sht in wb.sheets:
                        try:
                            file = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                            if file.empty or sht.name == '宅配':
                                lst = sht.used_range.value
                                file = pd.DataFrame(lst[1:], columns=lst[0])
                            # print(file)
                            if file is not None and len(file) > 0:
                                print('++++正在写入临时表：......')
                                file.to_sql('d1', con=self.engine1, index=False, if_exists='replace')
                            else:
                                print('----读取数据为空！！！')
                        except Exception as e:
                            print('xxxx读取失败：' + sht.name, str(Exception) + str(e))
                    wb.close()
                    app.quit()
        print('获取耗时：', datetime.datetime.now() - start)

    def creatReadSheet(self, team):  # 最近五天的全部订单信息
        match = {'slgat': '"神龙家族-港澳台"',
                 'slgat_hfh': '"火凤凰-港澳台"',
                 'sltg': '"神龙家族-泰国"',
                 'slxmt': '"神龙家族-新加坡", "神龙家族-马来西亚", "神龙家族-菲律宾"',
                 'slxmt_t': '"神龙-T新马菲"',
                 'slxmt_hfh': '"火凤凰-新加坡", "火凤凰-马来西亚", "火凤凰-菲律宾"',
                 'slrb': '"神龙家族-日本团队"',
                 'slrb_jl': '"精灵家族-日本", "精灵家族-韩国", "精灵家族-品牌"'}
        start = datetime.datetime.now()
        month_last = '2021-01-01'
        month_yesterday = '2021-04-20'
        print('正在获取需要查询的订单编号......')
        sql = '''SELECT id, sl.`订单编号`  FROM d1;'''
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        ordersDict = ', '.join(ordersDict)

        print('正在获取订单详情......')
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
             		        gs.product_name 产品名称,
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
             				left join (SELECT * FROM gk_sale WHERE id IN (SELECT MAX(id) FROM gk_sale GROUP BY product_id ) ORDER BY id) gs ON gs.product_id = a.product_id
                            left join dim_trans_way ON dim_trans_way.id = a.logistics_id
                            left join dim_cate ON dim_cate.id = a.third_cate_id
                            left join intervals ON intervals.id = a.intervals
                            left join dim_currency_lang ON dim_currency_lang.id = a.currency_lang_id
                    WHERE  a.order_number IN ({0}) AND dim_area.name IN ({1})
                        AND a.rq = '{2}' AND a.rq <= '{3}';'''.format(ordersDict, match[team], month_last, month_yesterday)
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
        print('++++++正在写入数据库++++++')
        try:
            df.to_sql('d0_sl', con=self.engine1, index=False, if_exists='replace')
            sql = 'REPLACE INTO d0_sl_list SELECT *, NOW() 记录时间 FROM d0_sl; '.format(team)
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
        except Exception as e:
            print('插入失败：', str(Exception) + str(e))
        print('写入完成…………')
        return '写入完成'