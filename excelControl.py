import pandas as pd
import os
import math
import xlwings
import numpy as np
import datetime
import zhconv          # transform2_zh_hant：转为繁体;transform2_zh_hans：转为简体
from mysqlControl import MysqlControl
from settings import Settings
from sqlalchemy import create_engine
from 查询_订单检索 import QueryOrder
class ExcelControl(Settings):
    '''
    excel的导入和整理
    '''
    # __slots__ = ['filePath', 'team', 'sql']
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
        self.engine4 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql4['user'],
                                                                                    self.mysql4['password'],
                                                                                    self.mysql4['host'],
                                                                                    self.mysql4['port'],
                                                                                    self.mysql4['datebase']))
        self.sql = MysqlControl()
        '''用来和数据库通信'''
    def readExcel(self, filePath, team):
        '''
        读取指定团队的指定文件，获取的数据整理格式之后，导入数据库
        :param filePath: excel文件路径
        :param team: 团队的代码
        :return: 无返回值
        '''
        # 文件扩展名
        fileType = os.path.splitext(filePath)[1]
        # 用xlwings来读取excel，因为pd.excel_read总是失败
        app = xlwings.App(visible=False, add_book=False)
        # 不显示excel窗口
        app.display_alerts = False

        if 'xls' in fileType:
            # 打开excel文件，只读方式，不更新公式，连接
            wb = app.books.open(filePath, update_links=False, read_only=True)
            # 遍历每个sheet
            for sht in wb.sheets:
                # 部分不可能是明细表的sheet直接排除，因为这些表的字段结构和明细表可能相似。
                if '问题件' not in sht.name and '录单' not in sht.name and '退件问题' not in sht.name and \
                        '历史' not in sht.name and '取件' not in sht.name and \
                        '异常' not in sht.name and sht.api.Visible == -1:
                    try:    # 可能会读取sheet内容失败，所以写了这个
                        db = None
                        # 读取sht的所有已使用单元格内容，并转换为pd的DateFrame格式
                        file = sht.used_range.options(pd.DataFrame, header=1, numbers=int, index=False).value
                        # print(file)
                        # 如果xlwings的直接转换失败的话。读取单元格值，并转换为db.DateFrame格式
                        if file.empty or sht.name == '宅配':
                            lst = sht.used_range.value
                            file = pd.DataFrame(lst[1:], columns=lst[0])
                            # print(file)
                        db = self.isRightSheet(file, team, sht.name)
                    except Exception as e:
                        print('xxxx查看失败：' + sht.name, str(Exception) + str(e))
                    if db is not None and len(db) > 0:
                        print('++++正在导入：' + sht.name + ' 共：' + str(len(db)) + '行', 'sheet共：' + str(
                            sht.used_range.last_cell.row) + '行')
                        # 将返回的dateFrame导入数据库的临时表
                        self.sql.writeSqlReplace(db)
                        print('++++正在更新：' + sht.name + '--->>>到总订单')
                        # print(db)
                        # 将数据库的临时表替换进指定的总表
                        self.sql.replaceInto(team, list(db.columns))
                        print('++++----->>>' + sht.name + '：订单更新完成++++')
                    else:
                        print('----------不用导入(五项条件不足)：' + sht.name)
                else:
                    print('----不用导入：' + sht.name)
            wb.close()              # 关闭excel文件
            app.quit()              # 退出excel app
    def date(para):
        delta = pd.Timedelta(str(para) + 'days')
        time = pd.to_datetime('1899-12-30') + delta
        return time
    def isRightSheet(self, df, team, shtName):
        '''
        根据团队，判断DateFrame是否是正确的明细表，并整理数据，输出需要的格式
        :param df: 要判断的DateFrame
        :param team: 哪个团队的数据
        :param shtName: sht名字，有些sheet需要特殊处理
        :return: 可以直接导入数据库的DateFrame
        '''
        math = {'slrb': {'出货时间': [True, ['预报日期', '出货日期', '出库日期', '发表时间', '日期', '出货日', '发货日期', '业务日期', '物流发货日期', '提货日期', '到货日期'], []],
                         '订单编号': [True, ['原单号', '顾客管理号码', '订单编号', '订单号', '内单号', '单号', '参考号'], []],
                         '运单编号': [True, ['渠道转单号', '系统单号', '系统运单号', '转单号', 'BJT转单号', '系统运单号', '运单编号', '跟踪号',
                                         '运单号', '跟踪单号', '改派单号', '国内物流运单号（订单号）'], []],
                         '物流状态': [True, ['状况', '物流状态', '状态', '轨迹状态'], []],
                         '状态时间': [True, ['轨迹日期', '状态时间', '时间', '末条信息日期时间', '出货预定日', '签收时间'], []],
                         '航班时间': [False, ['航班起飞时间'], []],
                         '清关时间': [False, ['日本清关时间', '清关时间'], []],
                         '上线时间': [False, ['''上线时间
（即货交地派时间）''', '上线时间（即货交地派时间）', '日本清关时间', '发送操作时间', '清关通行时间', '预计出港日期'], ['即货交地派时间']],
                         '原运单号': [False, ['原包裹运单号(可含多个)', '原运单号', '转单号'], []]},
                'sl_rb': {'出货时间': [True, ['出货时间', '预报日期', '出货日期', '出库日期', '发表时间', '日期', '出货日', '发货日期', '业务日期', '物流发货日期', '提货日期', '到货日期'], []],
                        '订单编号': [True, ['原单号', '顾客管理号码', '订单编号', '订单号', '内单号', '单号', '参考号'], []],
                        '运单编号': [True, ['渠道转单号', '系统单号', '系统运单号', '转单号', 'BJT转单号', '系统运单号', '运单编号', '跟踪号',
                                        '运单号', '跟踪单号', '改派单号', '国内物流运单号（订单号）'], []],
                        '物流状态': [True, ['状况', '物流状态', '状态', '标准状态', '轨迹状态'], []],
                        '状态时间': [True, ['轨迹日期', '状态时间', '时间', '末条信息日期时间', '出货预定日', '签收时间'], []],
                        '航班时间': [False, ['航班起飞时间'], []],
                        '清关时间': [False, ['日本清关时间', '清关时间'], []],
                        '上线时间': [False, ['''上线时间
（即货交地派时间）''', '上线时间（即货交地派时间）', '日本清关时间', '发送操作时间', '清关通行时间', '预计出港日期'], ['即货交地派时间']],
                        '原运单号': [False, ['原包裹运单号(可含多个)', '原运单号', '转单号'], []]},
                'sltg': {'出货时间': [True, ['提货日期', '发货日期', '接收订单资料日期', '出货时间', '揽收日期', '日期', 'Inbound Date'], []],
                         '订单编号': [True, ['订单号', '新订单号'], []],
                         '运单编号': [True, ['运单号', '新运单号', '转单号', '跟踪单号', 'Tracking ID'], []],
                         '物流状态': [True, ['物流状态', '订单状态', '最新扫描类型', 'Status'], []],
                         '状态时间': [True, ['状态时间', '時間', '最新扫描时间', 'Last Update/Scan'], []],
                         '问题明细': [False, ['问题件明细'], []],
                         '航班时间': [False, ['航班起飞时间', '''国内清关时间
（或航班起飞时间）'''], ['航班起飞时间']],
                         '清关时间': [False, ['''泰国清关时间
（可用到泰国时间代替）''', '泰国清关时间'], ['泰国清关时间']],
                         '上线时间': [False, ['''上线时间
（即货交地派时间）''', '上线时间'], ['货交地派时间'], 'Inbound Date'],
                         '原运单号': [False, ['原包裹运单号(可含多个)', '原运单号'], []]},
                'slsc': {'出货时间': [True, ['提货日期', '录单时间', '提交时间', '出货日期', '安排日期', '出库日期', '创建日期', '核重时间', '下单时间'], []],
                         '订单编号': [True, ['订单号', '原单号', '订单号码', '订单编号'], []],
                         '运单编号': [True, ['系统单号', '转单号', '运单编号', 'BJT转单号', '渠道转单号', '新单号', '提单号', '运单号', '承运单号'], []],
                         '物流状态': [True, ['物流状态', '状态', '订单状态', '运单最新状态', '货态'], []],
                         '状态时间': [True, ['状态时间', '日期',  '签收时间', '轨迹日期', '运单最新状态时间', '最新状态', '最终状态时间', '最新状态时间', '末条时间', '最终货态时间'],[]],
                         '问题明细': [False, ['问题件明细'], []],
                         '航班时间': [False, ['航班起飞时间'], ['航班起飞时间']],
                         '清关时间': [False, ['清关时间', '日本清关时间'], []],
                         '上线时间': [False, ['''上线时间
（即货交地派时间）''', '出库时间', '上线时间'], ['货交地派时间'], 'Inbound Date'],
                         '原运单号': [False, ['原包裹运单号(可含多个)', '原运单号', '原单号'], []]},
                'slgat': {'出货时间': [True, ['出货日期', '出货时间', '核重时间', '出库日期', '重出日期', '安排日期', '發貨日期',
                                          '收件日期', '业务日期', '出库时间', '发货日期', '下单时间'], []],
                          '订单编号': [True, ['订单编号', '订单号', '订单号码', '客户单号', '内部单号', '原始订单号',
                                          '件號', '件号'], []],
                          '运单编号': [True, ['运单号', '新单号', '提单号', '查件单号', '重出單號', '重出单号', '重出新單號', '單號',
                                          '重出新单号', '承运单号', '运单编号', '转单号码', 'SF转单号', '转单号', '转单'], []],
                          '物流状态': [True, ['物流状态', '状态', '运单最新状态', '貨態', '货态', '货态内容', '狀態',
                                          '新单号货态'], []],
                          '状态时间': [True, ['最新状态时间', '最新货态日期', '末条时间', '运单最新状态时间', '状态时间', '簽收日期',
                                          '最终状态时间', '最终货态时间', '新货态日期', '最新状态', '签收时间',
                                          '时间'], []],
                          '航班时间': [False, ['航班起飞时间', '''国内清关时间
（或航班起飞时间）''', '起飞时间'], ['航班起飞时间']],
                          '清关时间': [False, ['''清关时间
''', '清关时间'], ['泰国清关时间']],
                          '上线时间': [False, ['''上线时间
（即货交地派时间）''', '上线时间', '新竹上线时间', '到港日期'], ['货交地派时间']],
                          '原运单号': [False, ['原单号', '原單號', '原始顺丰订单号'], []]},
                'gat': {'出货时间': [True, ['出货日期', '出货时间', '核重时间', '出库日期', '重出日期', '安排日期', '转寄安排日期', '發貨日期', '出貨日期', '資料日期',
                                        '收件日期', '业务日期', '出库时间', '发货日期', '下单时间', '交货日期'], []],
                        '订单编号': [True, ['订单编号', '订单号', '订单号码', '客户单号', '客人订单号', '内部单号', '原始订单号', '運單號碼',
                                        '件號', '件号'], []],
                        '运单编号': [True, ['运单号', '新单号', '提单号', '查件单号', '重出單號', '重出单号', '重出新單號', '單號', '配送單號',
                                        '重出新单号', '承运单号', '运单编号', '转单号码', '顺丰单号', 'SF转单号', '转单号', '转单'], []],
                        '物流状态': [True, ['物流状态', '状态', '运单最新状态', '貨態', '货态', '货态内容', '狀態', '貨態分類',
                                        '新单号货态'], []],
                        '状态时间': [True, ['最新状态时间', '最新货态日期', '末条时间', '运单最新状态时间', '状态时间', '货态日期', '簽收日期', '最新發生時間',
                                        '最终状态时间', '最终货态时间', '新货态日期', '最新状态', '签收时间', '最後操作日期', '时间'], []],
                        '航班时间': [False, ['航班起飞时间', '''国内清关时间
（或航班起飞时间）''', '起飞时间'], ['航班起飞时间']],
                        '清关时间': [False, ['''清关时间''', '''清关时间
'''], []],
                        '原运单号': [False, ['原单号', '原單號', '原始顺丰订单号','原运单号'], []],
                        '上线时间': [False, ['上线时间', '''上线时间
（即货交地派时间）''', '货交地派时间', '到港日期', '''清关时间
''', '''新竹上线时间'''], []]},
                'slxmt': {'出货时间': [True, ['出货时间', 'Inbound Datetime'], []],
                        '订单编号': [True, ['订单号', '订单编号', 'Shipper Order Number', 'Shipper Reference Number'], []],
                        '运单编号': [True, ['转单号', '运单号', '运单编号', 'Tracking ID', 'Tracking Id ', 'tracking_id',
                                        'Tracking ID', 'Waybill Number'], []],
                        '物流状态': [True, ['状态', '订单状态', 'Granular Status', 'Status', 'status', 'Order Status'], []],
                        '状态时间': [True, ['Last Update/Scan', 'Last Delivery Date', 'Last Delivey Date', '日期',
                                        'Latest Service End Time', 'Last Valid Delivery Attempt Datetime', '签收时间',
                                        'Last Valid Delivery Attempt Date', 'Last Valid', 'SigningTime'], []],
                        '航班时间': [False, [], []],
                        '清关时间': [False, [], []],
                        '上线时间': [False, ['提取时间', 'Inbound Date'], []],
                        '原运单号': [False, [], []]}
                }
        necessary = 0
        # 初始化字段是否是必须的字段计数
        unnecessary = 0
        # 初始化字段是否是非必须的字段计数
        needDrop = []
        columns = list(df.columns)
        # print(columns)
        if '协来运-新竹' in shtName or '协来运-黑貓' in shtName or '协来运-重出' in shtName:
            if '單號' in df.columns:
                df.drop(labels=['單號'], axis=1, inplace=True)
        if '圓通-轉單表' in shtName:
            df.insert(0, '订单编号', '')  # 圆通直发
            # df.drop(labels=['發貨日期'], axis=1, inplace=True)  # 圆通改派
            # df.rename(columns={'到港日期': '發貨日期'}, inplace=True)
        if '圓通-國內單號' in shtName:
            if '單號' in df.columns and '發貨日期' in df.columns and '到港日期' in df.columns and '狀態' in df.columns and '最後操作日期' in df.columns:
                df.insert(0, '订单编号', '')        # 圆通直发
        if team == 'slgat' or team == 'gat':
            if '运单号' in columns and '查件单号' in columns and '订单编号' in columns and '换单号' in columns:
                df.drop(labels=['查件单号'], axis=1, inplace=True)     # 速派7-11的去掉多余的查件单号
            if '运单编号' in columns and '客户单号' in columns and '转单号' in columns:
                df.drop(labels=['转单号'], axis=1, inplace=True)     # 顺航的去掉多余的转单号
            if '运单编号' in columns and '件号' in columns and '转单号' in columns:
                df.drop(labels=['运单编号'], axis=1, inplace=True)   # 立邦的去掉多余的运单编号
            if '新单号' in columns and '承运单号' in columns:
                df.drop(labels=['承运单号'], axis=1, inplace=True)   # 天马的去掉多余的承运单号
            if '加' in columns and '原单号' in columns and '改派失败原因' in columns:
                df.drop(labels=['原单号'], axis=1, inplace=True)   # 龟山的去掉多余的承运单号
            if '客人订单号' in columns and '内部单号' in columns and '转单号码' in columns and '架位' in columns:
                df.rename(columns={'内部单号': '吉客印内部单号'}, inplace=True)
            if '交货日期' in columns and '出货时间' in columns and '件号' in columns:
                df.drop(labels=['交货日期'], axis=1, inplace=True)   # 立邦直发的去掉多余的交货日期
            # print(df)
            # print(df.columns)
        if team == 'slxmt':
            if '出货时间' not in df:
                df.insert(0, '出货时间', '')
            if '订单编号' not in df:
                df.insert(0, '订单编号', '')
            # df['状态时间'] = pd.to_datetime(df['状态时间'])

        if team == 'slsc':
            if '订单号' not in df:
                df.insert(0, '订单号', '')
            if '原单号' in columns and '转单号' in columns and '渠道转单号' in columns:
                df.drop(labels=['转单号'], axis=1, inplace=True)   # 天马的去掉多余的承运单号
            if '出货时间' in columns and '提货日期' in columns and '渠道' in columns:
                df.drop(labels=['出货时间'], axis=1, inplace=True)   # 天马的去掉多余的承运单号
            if '改派失败原因' in columns and '原单号' in columns and '订单编号' in columns:
                df.drop(labels=['原单号'], axis=1, inplace=True)   # 龟山的去掉多余的承运单号
        if team == 'slrb' or team == 'sl_rb':
            if '内单号' in columns and '转单号' in columns and '原单号' in columns:  # 吉客印神龙直发签收表JP使用
                df.drop(labels=['转单号'], axis=1, inplace=True)
                df.rename(columns={'内单号': '运单编号'}, inplace=True)
            if '单号' in columns and '转单号' in columns and '改派单号' in columns:  # 返品改派签收表使用
                df.drop(labels=['转单号'], axis=1, inplace=True)
            if '运单号' in columns and '转单号' in columns:  # 返品表使用
                # df.drop(labels=['转单号'], axis=1, inplace=True)
                df.rename(columns={'转单号': '原运单号'}, inplace=True)
            if 'BJT转单号' in columns and '跟踪单号' in columns:
                df.drop(labels=['跟踪单号'], axis=1, inplace=True)
                # print(df.columns)
            if '参考号' in columns and '跟踪号' in columns and '渠道转单号' in columns:   # 义达表
                df.drop(labels=['渠道转单号'], axis=1, inplace=True)
                # print(df.columns)
            if '订单号' in columns and '原单号' in columns:
                df.drop(labels=['原单号'], axis=1, inplace=True)
                # print(df.columns)
            if '运单号' in columns:
                df['运单号'] = df['运单号'].str.strip()  # 去掉运单号中的前后空字符串
            if '订单编号' not in df:
                df.insert(0, '订单编号', '')
            if '发表时间' in columns and '国内物流运单号（订单号）' in columns and '日期' in columns:  # 物流轨迹去掉多余的日期
                df.drop(labels=['日期'], axis=1, inplace=True)
            # print(columns)
        if team == 'sltg':
            if '订单号' not in df:
                df.insert(0, '订单号', '')
        if shtName == '宅配':
            # 宅配的原单编号可能出现重复，重复的话，会在A字段列，出现一个新的单号，以 A原运单号 的形式出现，所以如果A字
            # 段列有内容的话，直接替换掉原有的运单编号的列就能得到符合后台的运单编号了。
            df['重出新单号'] = np.where(df['A'].isnull(), df['重出新单号'], df['A'])
        elif shtName == '总明细':
            # 有一个叫做总明细的sheet，里面会重复出现字段，且前面的字段有内容，所以要丢弃后面无用的字段。避免冲突
            columns = list(df.columns)
            for index, column in enumerate(columns):
                if column == '状态' and index > 10:
                    columns[index] = column + str(index)
                    df.columns = columns
                    break
            df.drop(columns=df.columns[10:], axis=1, inplace=True)
            if team == 'slgat' or team == 'gat':
                if '状态' in columns and '货态' in columns:
                    df.drop(labels=['货态'], axis=1, inplace=True)
        elif shtName == 'LIST':
            # 有个叫LIST的sheet，在系统里所有的运单号和订单编号一样，所以把签收表里面的运单编号，替换成订单编号
            df['渠道转单号'] = df['内单号']
        columns = list(df.columns)
        # 保留一个列名，后面要用
        for index, column in enumerate(columns):
            if not column:
                # 如果列名为空，肯定不是需要的列，起一个名字，标记，后面要删除
                columns[index] = 'needDrop' + str(index)
                column = 'needDrop' + str(index)
            for k, v in math[team].items():
                # 遍历字段匹配字典
                if column in v[1]:
                    # 如果列名完全匹配需要的字段，则，字段重命名为标准字段名
                    columns[index] = k
                    if k in columns[:index]:
                        # 如果这个需要的字段，之前出现过，则添加到需要删除的列表里面
                        tem = k + str(columns.index(k, 0, index))
                        columns[columns.index(k, 0, index)] = tem
                        needDrop.append(tem)
                        if v[0]:
                            necessary -= 1
                    break
                else:
                    for vs in v[2]:
                        # 模糊匹配，因为担心出错，所以模糊匹配的关键字，没有写。所以这一段，应该不会生效。
                        if vs in column:
                            columns[index] = k
                            if k in columns[:index]:
                                tem = k + str(columns.index(k, 0, index))
                                columns[columns.index(k, 0, index)] = tem
                                needDrop.append(tem)
                                if v[0]:
                                    necessary -= 1
                            break
            if k != columns[index]:
                needDrop.append(columns[index])
            else:
                if v[0]:
                    necessary += 1
                else:
                    unnecessary += 1
        # print(df.columns)
        print(88)
        # print(needDrop)
        # print(necessary)   # 五项条件不足 信息

        if necessary >= 5:
            # print(df['订单编号'])
            df.columns = columns
            df.drop(labels=needDrop, axis=1, inplace=True)
            df.dropna(axis=0, subset=['运单编号'], inplace=True)
            # print(df.columns)
            # if team == 'gat':
            #     print(88)
            #     print(df['物流状态'])
            #     df['物流状态'] = zhconv.convert(df['物流状态'], 'zh-hans')
            if team == 'slrb' or team == 'sl_rb':
                try:
                    # df['状态时间'] = df['状态时间'].str.strip()
                    df['状态时间'] = df['状态时间'].replace(to_replace=0, value=datetime.datetime(1990, 1, 1, 0, 0))
                    df['状态时间'] = df['状态时间'].replace(to_replace=' ', value=datetime.datetime(1990, 1, 1, 0, 0))
                    df['状态时间'] = df['状态时间'].replace(to_replace='客人联系保管', value=(datetime.datetime.now() - datetime.timedelta(days=1)))
                    #df['上线时间'] = df['上线时间'].replace(to_replace=0, value=datetime.datetime(1990, 1, 1, 0, 0))
                    df['状态时间'] = df['状态时间'].fillna(value=datetime.datetime(1990, 1, 1, 0, 0))
                    df['物流状态'] = df['物流状态'].fillna(value='未上线')
                    df['订单编号'] = df['订单编号'].str.replace('原内单号：', '')
                    df['订单编号'] = df['订单编号'].str.replace('原单:', '')
                    # df['出货时间'] = df['出货时间'].astype(str)
                    # print(df['出货时间'])
                    # print(df['运单编号'])
                except Exception as e:
                    print('----修改状态时间失败：', str(Exception) + str(e))
            if team == 'slsc':
                df['状态时间'] = df['状态时间'].replace(to_replace=0, value=datetime.datetime(1990, 1, 1, 0, 0))
                df['状态时间'] = df['状态时间'].replace(to_replace='取消', value=datetime.datetime(1990, 1, 1, 0, 0))
                df['状态时间'] = df['状态时间'].replace(to_replace='-', value=datetime.datetime(1990, 1, 1, 0, 0))
                df['状态时间'] = df['状态时间'].replace(to_replace=' ', value=datetime.datetime(1990, 1, 1, 0, 0))
                df['状态时间'] = df['状态时间'].replace(to_replace='通知取消', value=datetime.datetime(1990, 1, 1, 0, 0))
                df['状态时间'] = df['状态时间'].replace(to_replace='10.20退回吉客印仓库', value=datetime.datetime(1990, 1, 1, 0, 0))
                df['状态时间'] = df['状态时间'].fillna(value=datetime.datetime(1990, 1, 1, 0, 0))
                df['物流状态'] = df['物流状态'].fillna(value='未上线')
            if shtName in ['新竹']:
                df['订单编号'] = df['订单编号'].str.replace('原单:', '')
            if shtName in ['全家']:
                df['订单编号'] = df['订单编号'].str.replace('原单:', '')
                df['运单编号'] = df['订单编号']
            df['运单编号'] = df['运单编号'].astype(str)
            df['运单编号'] = df['运单编号'].replace(to_replace=r'\.0$', regex=True, value=r'')
            if '原运单号' in columns:
                df['原运单号'] = df['原运单号'].astype(str)
                df['原运单号'] = df['原运单号'].replace(to_replace=r'\.0$', regex=True, value=r'')
            if team == 'slgat' or team == 'gat':
                df['状态时间'] = df['状态时间'].replace(to_replace=0, value=datetime.datetime(1990, 1, 1, 0, 0))
                df['状态时间'] = df['状态时间'].fillna(value=datetime.datetime(1990, 1, 1, 0, 0))
                df['订单编号'] = df['订单编号'].astype(str)
                # df = df[~(df['订单编号'].str.contains('TW|XM'))]
                df.reset_index(drop=True, inplace=True)
            elif team == 'sltg':
                df['订单编号'] = df['订单编号'].astype(str)
                df = df[~(df['订单编号'].str.contains('BJ|GK|KD|NB|NR|TG|TR|XM'))]
                df.reset_index(drop=True, inplace=True)

            if '7-11' in shtName or '711' in shtName or ('新竹' in shtName and '协来运' not in shtName):  # 添加上线时间新竹
                df.insert(0, '上线时间', '')
                if '清关时间' in df.columns or '''清关时间
        ''' in df.columns:
                    df['上线时间'] = df['清关时间'].copy()
                else:
                    df['上线时间'] = df['出货时间'].copy()
                # print(11)
                # print(df['上线时间'])
                # print(df['清关时间'])
            if '协来运' in shtName:  # 添加上线时间  协来运
                df.insert(0, '上线时间', '')
                df['上线时间'] = df['出货时间'].copy()
            if '重出明细' in shtName or '重出明细' in shtName:                       # 添加上线时间  天马改派
                df.insert(0, '上线时间', '')
                df['上线时间'] = df['出货时间'].copy()
            if '转寄表' in shtName:                          # 添加上线时间  立邦 改派
                df.insert(0, '上线时间', '')
                df['上线时间'] = df['出货时间'].copy()
            if '立邦-明细' in shtName:                          # 添加上线时间  立邦 改派
                df.insert(0, '上线时间', '')
                df['上线时间'] = df['出货时间'].copy()
            if '月明细' in shtName:                          # 添加上线时间  易速配
                df.insert(0, '上线时间', '')
                df['上线时间'] = df['清关时间'].copy()
            if '月重出' in shtName or '月总表' in shtName or '备货转寄' in shtName:       # 添加上线时间  易速配 改派
                df.insert(0, '上线时间', '')
                df['上线时间'] = df['出货时间'].copy()
            if '圓通-轉單表' in shtName:
                df['上线时间'] = df['出货时间'].copy()
            # print(df['上线时间'])
            return df
        else:
            return None

    def readReturnOrder(self, team):
        import os
        # 用pandas读取excel，结果日期自动的变为了数字
        def date(para):
            if type(para) == int:
                delta = pd.Timedelta(str(int(para)) + 'days')
                time = pd.to_datetime('1899-12-30') + delta
                return time
            else:
                return para

        path = r'D:\Users\Administrator\Desktop\需要用到的文件\A退货'
        dirs = os.listdir(path=path)
        for dir in dirs:
            filePath = os.path.join(path, dir)
            print(filePath)
            if dir[:2] != '~$':
                df = pd.read_excel(filePath)
                columns_value = list(df.columns)  # 获取数据的标题名，转为列表
                # for column_val in columns_value:
                #     if '订单编号' != column_val:
                #         df.drop(labels=[column_val], axis=1, inplace=True)  # 去掉多余的旬列表
                print(df.columns)
                if '吉客印新竹退款明细' in filePath:
                    df.insert(0, '物流', '速派')
                else:
                    df.insert(0, '物流', '天马')
                df.rename(columns={'订单号': '物流订单编号','单号': '运单编号', '日期 ': '日期'}, inplace=True)
                df['日期'] = df['日期'].apply(date)
                df = df[['日期', '物流订单编号', '运单编号', '取件单号', '代收金额', '处理结果', '物流']]
                df.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
                print('正在写入退货表 订单信息…………')
                sql = '''INSERT IGNORE INTO {}_return (日期, 订单编号, 物流订单编号, 联系电话, 运单编号, 取件单号, 代收金额, 处理结果, 发货时间, 物流状态, 完成时间, 上线时间, 物流, 添加时间)
                                                SELECT 日期, null 订单编号, 物流订单编号, null 联系电话, 运单编号, 取件单号, 代收金额, 处理结果, 
                                                null 发货时间, null 物流状态, null 完成时间, null 上线时间, 物流, NOW() 添加时间 FROM customer; '''.format(team)
                pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
                self.readReturnOrder_Update(team, filePath)

    def readReturnOrder_Update(self, team, filePath):
        print('正在检查 需补充的 订单信息…………')
        sql = '''SELECT 运单编号 FROM {}_return r WHERE r.`订单编号` IS NULL OR r.`订单编号` = "";'''.format(team)
        ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
        print(ordersDict)

        if ordersDict.empty:
            print(' ****** 没有要补充的订单信息; ****** ')
        else:
            print('！！！ 请再次更新订单编号数据！！！')
            proxy_handle = '代理服务器0'
            proxy_id = '192.168.13.89:37467'                            # 输入代理服务器节点和端口
            handle = '手0动'
            login_TmpCode = '517e55c6fb6c34ca99a69874aaf5ec25'          # 输入登录口令Tkoen
            js = QueryOrder('+86-18538110674', 'qyz04163510.', login_TmpCode, handle, proxy_handle, proxy_id)

            orders_Dict = list(ordersDict['运单编号'])
            max_count = len(orders_Dict)
            if max_count > 0:
                in_count = math.ceil(max_count / 500)
                df = pd.DataFrame([])
                dlist = []
                n = 0
                while n < in_count:  # 这里用到了一个while循环，穿越过来的
                    print('查询第 ' + str(n) + ' 页中，剩余次数' + str(in_count - n))
                    n1 = n * 500
                    n2 = (n + 1) * 500
                    ord = ','.join(orders_Dict[n1:n2])
                    print(ord)
                    data = js.orderInfoQuery(ord, '运单号', proxy_id, proxy_handle)
                    # data = js.orderInfo_pople(ord, '运单号', proxy_id, proxy_handle)
                    print(data)
                    dlist.append(data)
                    n = n + 1
                dp = df.append(dlist, ignore_index=True)
                if dp.empty:
                    print(' ****** 没有要补充的订单信息; ****** ')
                else:
                    print('正在写入......')
                    dp = dp[['orderNumber','wayBillNumber', 'shipInfo.shipPhone', 'deliveryTime','logisticsStatus', 'onlineTime','finishTime']]
                    dp.to_sql('customer', con=self.engine1, index=False, if_exists='replace')
                    sql = '''update `gat_return` a, customer b
                                SET a.`订单编号` = IF(b.`orderNumber` = '', NULL, b.`orderNumber`),
                                    a.`联系电话` = IF(b.`shipInfo.shipPhone` = '', NULL, b.`shipInfo.shipPhone`),
                                    a.`发货时间` = IF(b.`deliveryTime` = '', NULL, b.`deliveryTime`),
                                    a.`物流状态` = IF(b.`logisticsStatus` = '', NULL, b.`logisticsStatus`),
                                    a.`上线时间` = IF(b.`onlineTime` = '', NULL, b.`onlineTime`),
                                    a.`完成时间` = IF(b.`finishTime` = '', NULL, b.`finishTime`)
                            where a.`运单编号`=b.`wayBillNumber`;'''
                    pd.read_sql_query(sql=sql, con=self.engine1, chunksize=10000)
            else:
                print('无需补充数据')

        if '吉客印新竹退款明细' in filePath or '现退明细' in filePath:
            rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
            new_path = "F:\\神龙签收率\\" + (datetime.datetime.now()).strftime('%m.%d') + "\\导状态\\{0} 导退货状态-临时.xlsx".format(rq)
            print('正在输出 退货退款 导状态表…………')
            sql = '''SELECT 订单编号, 联系电话, 运单编号 AS 回执单号, 发货时间,  '退货退款' AS 物流状态,  4 AS 完成状态,  完成时间 AS 更新时间, 上线时间 
                    FROM {0}_return r 
                    WHERE DATE_FORMAT(r.添加时间, '%Y-%m-%d') = DATE_FORMAT(CURDATE() , '%Y-%m-%d')
                         AND r.`订单编号` IS NOT NULL AND r.`订单编号` <>  "";'''.format(team)
            ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
            print(ordersDict)
            if ordersDict.empty:
                print(' ****** 无需生成导状态表 ****** ')
            else:
                ordersDict.to_excel(new_path, sheet_name='查询', index=False,engine='xlsxwriter')
        else:
            print('无需生成导状态表')

        print('退货更新文件成功…………')
        os.remove(filePath)
        print('已清除退货文件…………')

if __name__ == '__main__':
    e = ExcelControl()
    match = {'slrb': r'D:\Users\Administrator\Desktop\需要用到的文件\退货',
             'sltg': r'D:\Users\Administrator\Desktop\需要用到的文件\退货',
             'slgat': r'D:\Users\Administrator\Desktop\需要用到的文件\退货',
             'slxmt': r'D:\Users\Administrator\Desktop\需要用到的文件\退货'}
    e.readReturnOrder('slgat')                   # 先退货退款导入 后进行签收表的计算