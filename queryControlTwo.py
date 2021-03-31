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
            ord = ', '.join(orderId[n:n + 10])
            print(ord)
            n = n + 10
            self.productIdquery(tokenid, ord, searchType, team)

    def productIdquery(self, tokenid, orderId, searchType, team):  # 进入订单检索界面，
        start = datetime.datetime.now()
        url = r'http://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
        data = {'phone': None,
                'email': None,
                'ip': None,
                'page': 1,
                'pageSize': 300,
                '_token': tokenid}
        if searchType == '订单号':
            data.update({'orderPrefix': orderId,
                         'shippingNumber': None})
        elif searchType == '运单号':
            data.update({'order_number': None,
                         'shippingNumber': orderId})
        r_header = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36',
            'Referer': 'http://gimp.giikin.com/front/orderToolsServiceQuery'}
        req = self.session.post(url=url, headers=r_header, data=data)
        print(req)
        print('已成功发送请求++++++')
        print('正在处理json数据…………')
        req = json.loads(req.text)  # json类型数据转换为dict字典
        print('正在转化数据为dataframe…………')
        print(req)
        # print(req)
        ordersDict = []
        for result in req['data']['list']:
            # print(result)
            # 添加新的字典键-值对，为下面的重新赋值用
            result['productId'] = 0
            result['saleName'] = 0
            result['saleProduct'] = 0
            result['spec'] = 0
            result['link'] = 0
            # print(result['specs'])
            # spe = ''
            # spe2 = ''
            # spe3 = ''
            # spe4 = ''
            # # 产品详细的获取
            # for ind, re in enumerate(result['specs']):
            #     print(ind)
            #     print(re)
            #     print(result['specs'][ind])
            #     spe = spe + ';' + result['specs'][ind]['saleName']
            #     spe2 = spe2 + ';' + result['specs'][ind]['saleProduct']
            #     spe3 = spe3 + ';' + result['specs'][ind]['spec']
            #     spe4 = spe4 + ';' + result['specs'][ind]['link']
            #     spe = spe + ';' + result['specs'][ind]['saleProduct'] + result['specs'][ind]['spec'] + result['specs'][ind]['link'] + result['specs'][ind]['saleName']
            # result['specs'] = spe
            # # del result['specs']             # 删除多余的键值对
            # result['saleName'] = spe
            # result['saleProduct'] = spe2
            # result['spec'] = spe3
            # result['link'] = spe4
            result['saleName'] = result['specs'][0]['saleName']
            result['saleProduct'] = result['specs'][0]['saleProduct']
            result['spec'] = result['specs'][0]['spec']
            result['link'] = result['specs'][0]['link']
            result['productId'] = (result['specs'][0]['saleProduct']).split('#')[1]
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
            self.q.put(result)
        # print(len(req['data']['list']))
        for i in range(len(req['data']['list'])):
            ordersDict.append(self.q.get())
        data = pd.json_normalize(ordersDict)
        df = data[['orderNumber', 'wayBillNumber', 'logisticsName', 'logisticsStatus', 'orderStatus', 'isSecondSend',
                   'currency', 'area', 'currency', 'shipInfo.shipPhone', 'quantity', 'productId']]
        print('正在写入缓存中......')
        try:
            df.to_sql('d1_cp', con=self.engine1, index=False, if_exists='replace')
            if team == 'slgat' or team == 'slgat_hfh':
                sql = '''SELECT orderNumber,
                    logisticsName 物流方式,
                    IF(logisticsName LIKE '%易速配%', '易速配', IF(logisticsName LIKE '%速派%','速派国际', IF(logisticsName LIKE '%森鸿%', '森鸿国际', IF(logisticsName LIKE '%壹加壹%', '壹加壹', IF(logisticsName LIKE '%立邦%','立邦国际', IF(logisticsName LIKE '%天马%全家%','天马运通', IF(logisticsName LIKE '%天马%711%','天马运通', IF(logisticsName LIKE '%天马%新竹%','天马运通', IF(logisticsName LIKE '%天马%顺丰%','天马物流', logisticsName))))))))) 物流名称,
                    productId,
                    dp.`name`,
                    dc.ppname cate,
                    dc.pname second_cate,
                    dc.`name` third_cate
                FROM d1_cp
                LEFT JOIN dim_product dp ON  d1_cp.productId = dp.id
                LEFT JOIN dim_cate dc ON  dc.id = dp.third_cate_id;'''
                df = pd.read_sql_query(sql=sql, con=self.engine1)
            elif team == 'slrb':
                sql = '''SELECT orderNumber,
                        logisticsName 物流方式,
                        IF(logisticsName LIKE '%捷浩通%', '捷浩通', IF(logisticsName LIKE '%翼通达%','翼通达', IF(logisticsName LIKE '%博佳图%', '博佳图', IF(logisticsName LIKE '%保辉达%', '保辉达物流', IF(logisticsName LIKE '%万立德%','万立德', logisticsName))))) 物流名称,
    					productId,
    					dp.`name`,
    					dc.ppname cate,
    					dc.pname second_cate,
    					dc.`name` third_cate
    				FROM d1_cp
    				LEFT JOIN dim_product dp ON  d1_cp.productId = dp.id
    				LEFT JOIN dim_cate dc ON  dc.id = dp.third_cate_id;'''
                df = pd.read_sql_query(sql=sql, con=self.engine1)
            print(df)
            df.to_sql('tem_product_id', con=self.engine1, index=False, if_exists='replace')
            print('正在更新产品详情…………')
            sql = '''update {0}_order_list a, tem_product_id b
    		                        set a.`物流方式`=b.`物流方式`,
    		                            a.`物流名称`=b.`物流名称`,
    		                            a.`产品id`=b.`productId`,
    		                            a.`产品名称`=b.`name`,
    				                    a.`父级分类`=b.`cate` ,
    				                    a.`二级分类`=b.`second_cate`,
    				                    a.`三级分类`=b.`third_cate`
    				                where a.`订单编号`=b.`orderNumber`;'''.format(team)
            pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
        except Exception as e:
            print('更新失败：', str(Exception) + str(e))
        print('更新成功…………')
        print('更新耗时：', datetime.datetime.now() - start)

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
    token = '1b4a0d9c0f62c43e4a97f8cf250b06a8'
    m.productIdquery(token, 'NJ104012041487412,NJ104012042191134,NJ104012042233462,NJ104012043202092,NJ104012043306768,NJ210401204335042260,NJ104012044169468,NJ104012046077587,NJ104012046214215,NJ210401204622012232,NJ104012048022897,NJ104012048046694,NJ104012049037982,NJ104012049243726,NJ210401205103392080,NJ104012051196726,NJ104012051336710,NJ104012052171571,NJ104012052557314,NJ104012053207128,NJ104012054122863,NJ104012054528869,NJ104012054546698,NJ104012055102659,NJ210401205511857056,NJ104012055372986,NJ104012056187616,NJ104012056183870,NJ210401205648295986,NJ210401205728209001,NJ104012057318200,NJ104012057344838,NJ104012058233978,NJ104012058537969,NJ104012059105793,NJ104012100546139,NJ104012101206707,NJ104012101343249,NJ104012102081104,NJ104012102466369,NJ104012102514840,NJ104012102548635,NJ104012103526553,NJ104012104386982,NJ210401210718354149,NJ104012107372978,NJ104012108478460,NJ104012108552411,NJ104012110025023,NJ104012110072085,NJ104012110166896,NJ104012110238739,NJ104012111115043,NJ210401211113912385,NJ104012112104614,NJ210401211214801688,NJ210401211242295114,NJ210401211352621585,NJ104012115054687,NJ210401211509099792,NJ104012115323266,NJ210401211538327706,NJ210401211624190213,NJ210401211633878301,NJ210401211644953867,NJ104012116483953,NJ104012117181518,NJ210401211850157523,NJ104012118511799,NJ104012118576680,NJ210401212148976930,NJ210401212153934608,NJ104012122419537,NJ210401212400852944,NJ104012124097899,NJ104012124333731,NJ104012124528146,NJ104012125078688,NJ104012125315068,NJ104012126036935,NJ104012126054264,NJ104012126182745,NJ104012126281099,NJ104012126505322,NJ210401212728453396,NJ104012127365704,NJ104012127587212,NJ210401212816304576,NJ104012128208312,NJ104012129193170,NJ104012129585909,NJ104012131489475,NJ104012131494388,NJ104012132314424,NJ104012132347372,NJ104012132398046,NJ210401213315829229,NJ104012133536643,NJ104012134273590,NJ104012135049824,NJ210401213533747737,NJ104012135394410,NJ104012135547858,NJ104012136179832,NJ104012137254585,NJ210401213746758553,NJ104012138494947,NJ210401214027789124,NJ104012141151092,NJ104012141277858,NJ104012141389470,NJ104012142389520,NJ210401214247379144,NJ210401214247884834,NJ104012143388277,NJ104012144081415,NJ104012144364333,NJ210401214440226286,NJ104012146431796,NJ104012146569618,NJ210401214716549574,NJ210401214728814672,NJ210401214740555948,NJ104012147428252,NJ104012148362295,NJ104012150009299,NJ104012150249289,NJ104012150532290,NJ104012151287339,NJ104012151533882,NJ104012151533216,NJ104012154303351,NJ104012155032675,NJ104012155151190,NJ104012156112837,NJ104012156149551,NJ104012157302239,NJ210401215746903019,NJ104012157562191,NJ104012157574052,NJ104012158181562,NJ104012158487634,NJ210401215932220974,NJ104012200391759,NJ104012200417265,NJ210401220204484825,NJ104012202188663,NJ104012202373212,NJ104012202429501,NJ210401220305868425,NJ104012203173437,NJ104012203346009,NJ104012204135810,NJ104012204241030,NJ210401220447064457,NJ104012205113648,NJ104012205234592,NJ104012205413510,NJ104012205593812,NJ104012206065979,NJ104012207146683,NJ104012207546140,NJ104012208225132,NJ210401220859445120,NJ210401220942141617,NJ104012210006767,NJ104012210166307,NJ104012210596339,NJ104012212153796,NJ104012213451597,NJ104012214046936,NJ210401221548559250,NJ210401221556481956,NJ104012216526492,NJ104012217226008,NJ104012217394111,NJ104012217476139,NJ104012217518753,NJ104012218344137,NJ104012218573305,NJ104012219274446,NJ104012219388670,NJ104012220157928,NJ104012220152090,NJ104012221597414,NJ210401222241456496,NJ104012225113570,NJ104012225299319,NJ104012225315691,NJ104012225322198,NJ210401222651804713,NJ104012227517856,NJ210401222845324604,NJ104012229194107,NJ104012229237498,NJ104012229349919,NJ104012229425643,NJ210401222951300659,NJ104012230149821,NJ104012230506620,NJ104012230529030,NJ104012231089628,NJ104012231404614,NJ104012232147134,NJ104012233192198,NJ104012233219450,NJ210401223321378383,NJ104012233274202,NJ104012234144067,NJ104012235597627,NJ210401223600478847,NJ104012236164795,NJ104012236413850,NJ104012236596577,NJ104012238251290,NJ104012240021418,NJ104012241534014,NJ104012242299756,NJ210401224336244366,NJ104012243401558,NJ210401224508471703,NJ104012245495061,NJ104012246208407,NJ104012246277001,NJ210401224700940570,NJ104012247117036,NJ104012249031680,NJ104012251112756,NJ104012253449730,NJ104012253547493,NJ104012254421180,NJ104012257323402,NJ210401225735981930,NJ104012257443816,NJ210401225948534664,NJ210401225952590700,NJ210401230043941317,NJ104012300592227,NJ210401230135119168,NJ104012302191474,NJ104012302225610,NJ104012303298841,NJ104012303413655,NJ104012304093711,NJ104012304334254,NJ210401230635409652,NJ104012306421914,NJ210401230654416402,NJ104012307119468,NJ210401230727523969,NJ104012307576851,NJ104012308042560,NJ104012309377832,NJ104012311528283,NJ104012312151731,NJ210401231217336420,NJ104012313288954,NJ104012313336106,NJ104012313466772,NJ104012313484069,NJ104012314129950,NJ104012315286146,NJ104012315323182,NJ104012316029857,NJ104012316348169,NJ104012317107370,NJ104012319146700,NJ104012319395153,NJ104012322071149,NJ104012323239016,NJ104012325252619,NJ104012325559321,NJ104012326134309,NJ104012329115264,NJ104012329203362,NJ104012332291364,NJ104012333483750,NJ104012335572679,NJ104012336383425,NJ104012337414832,NJ104012339292137,NJ104012340044501,NJ104012340507237,NJ210401234126081986,NJ210401234326681508,NJ104012343487935,NJ104012343554443,NJ210401234442880893,NJ104012344451749,NJ104012345436527,NJ104012345434451,NJ104012348481743,NJ104012351269931,NJ104012351342748,NJ104012351438763,NJ210401235352951283,NJ104012357242873,NJ104012358081607,NJ104012358284931,', '订单号', team)
    # m.productIdInfo(token, '订单号', team)