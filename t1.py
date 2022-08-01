import os
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from excelControl import ExcelControl
from mysqlControl import MysqlControl
from wlMysql import WlMysql
from wlExecl import WlExecl
from sso_updata import Query_sso_updata
from gat_update2 import QueryUpdate
import datetime
from dateutil.relativedelta import relativedelta
start: datetime = datetime.datetime.now()
team = 'gat'
match1 = {'gat': '港台',
          'slsc': '品牌'}
match = {'slgat': r'D:\Users\Administrator\Desktop\需要用到的文件\A港台签收表',
         'slgat_hfh': r'D:\Users\Administrator\Desktop\需要用到的文件\A港台签收表',
         'slgat_hs': r'D:\Users\Administrator\Desktop\需要用到的文件\A港台签收表',
         'slrb': r'D:\Users\Administrator\Desktop\需要用到的文件\A日本签收表',
         'slrb_jl': r'D:\Users\Administrator\Desktop\需要用到的文件\A日本签收表',
         'slrb_js': r'D:\Users\Administrator\Desktop\需要用到的文件\A日本签收表',
         'slrb_hs': r'D:\Users\Administrator\Desktop\需要用到的文件\A日本签收表',
         'slsc': r'D:\Users\Administrator\Desktop\需要用到的文件\品牌',
         'gat': r'D:\Users\Administrator\Desktop\需要用到的文件\A港台签收表',
         'gat_upload': r'D:\Users\Administrator\Desktop\需要用到的文件\A港台签收表 - 单独上传',
         'sltg': r'D:\Users\Administrator\Desktop\需要用到的文件\A泰国签收表',
         'slxmt': r'D:\Users\Administrator\Desktop\需要用到的文件\A新马签收表',
         'slxmt_t': r'D:\Users\Administrator\Desktop\需要用到的文件\A新马签收表',
         'slxmt_hfh': r'D:\Users\Administrator\Desktop\需要用到的文件\A新马签收表'}
'''    msyql 语法:      show processlist（查看当前进程）;  
                        set global event_scheduler=0;（关闭定时器）;
备注：  港台 需整理的表：香港立邦>(明细再copy一份保存) ； 台湾龟山改派>(copy保存为xlsx格式);
说明：  日本 需整理的表：1、吉客印神龙直发签收表=密码：‘JKTSL’>(明细再copy保存；改派明细不需要);2、直发签收表>(明细再copy保存；3、状态更新需要copy保存);
'''
# 初始化时间设置
if team in ('slsc', 'slrb', 'slrb_jl', 'slrb_js', 'slrb_hs', 'gat'):
    # 更新时间
    yy = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y'))
    mm = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%m'))
    begin = datetime.date(yy, mm, 1)
    print(begin)
    yy2 = int(datetime.datetime.now().strftime('%Y'))
    mm2 = int(datetime.datetime.now().strftime('%m'))
    dd2 = int(datetime.datetime.now().strftime('%d'))
    end = datetime.date(yy2, mm2, dd2)
    print(end)
    # 导出时间
    month_last = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
    month_yesterday = datetime.datetime.now().strftime('%Y-%m-%d')
    month_begin = (datetime.datetime.now() - relativedelta(months=3)).strftime('%Y-%m-%d')
    print(month_begin)
else:
    # 更新时间
    begin = datetime.date(2021, 8, 1)
    print(begin)
    end = datetime.date(2021, 9, 1)
    print(end)
    # 导出时间
    month_last = '2021-08-01'
    month_yesterday = '2021-09-01'
    month_begin = '2021-07-01'

# 库的引用
if team == 'gat':
    path = match['gat_upload']
    dirs = os.listdir(path=path)
else:
    path = match[team]
    dirs = os.listdir(path=path)

e = ExcelControl()
m = MysqlControl()
w = WlMysql()
we = WlExecl()
qu = QueryUpdate()

# 上传退货
# e.readReturnOrder(team)
print('退货导入耗时：', datetime.datetime.now() - start)

# ---读取execl文件---
for dir in dirs:
    filePath = os.path.join(path, dir)
    print(filePath)
    if 'xlsx' not in filePath:
        wbsheet = filePath
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = excel.Workbooks.Open(filePath)
        wb.SaveAs(filePath + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
        wb.Close()  # FileFormat = 56 is for .xls extension
        excel.Application.Quit()
        filePath = filePath + "x"
        print(filePath)
        print('****** 已成功将 xls 转换成 xlsx 格式 ******')
        os.remove(wbsheet)
    if dir[:2] != '~$':
        wb_start = datetime.datetime.now()
        wb = load_workbook(filePath, data_only=True)
        wb.save(filePath)
        print('+++处理表格公式-耗时：', datetime.datetime.now() - wb_start)
        if dir[:6] == 'GIIKIN' or dir[:6] == 'Giikin':
            print('98')
            we.logisitis(filePath, team)
        else:
            print('02')
            e.readExcel(filePath, team)
        print('单表+++导入-耗时：', datetime.datetime.now() - wb_start)
print('导入耗时：', datetime.datetime.now() - start)


# ---获取 改派未发货 execl文件---
mkpath = r"F:\神龙签收率\(未发货) 改派-物流"
dirs_gp = os.listdir(path=mkpath)
wb = ''
for dir in dirs_gp:
    filePath = os.path.join(mkpath, dir)
    if (datetime.datetime.now()).strftime('%Y.%m.%d') in filePath:
        wb = '改派未发货 已导出'
if wb == '改派未发货 已导出':
    print(wb)
else:
    print('正在获取 改派未发货 中')
    handle = '手动'
    token = '3f9a3410b45035a180743c4a13093a05'
    sso = Query_sso_updata('+86-18538110674', 'qyz04163510.', '1343', token, handle)
    sso.gp_order()

print('*' * 50)

# TODO---数据库分段读取---
# m.creatMyOrderSl(team)  # 最近五天的全部订单信息
#
# print('------------更新部分：---------------------')
# if team in ('slsc', 'slrb', 'slrb_jl', 'slrb_js', 'slrb_hs'):
#     m.creatMyOrderSlTWO(team, begin, end)   # 最近两个月的更新订单信息
#     print('处理耗时：', datetime.datetime.now() - start)
#
#     print('------------导出部分：---------------------')
#     m.connectOrder(team, month_last, month_yesterday, month_begin)  # 最近两个月的订单信息导出
#     print('输出耗时：', datetime.datetime.now() - start)
#
# elif team in ('gat'):
#     sso = QueryTwo('+86-18538110674', 'qyz04163510')
#     print(datetime.datetime.now())
#     print('++++++正在获取 ' + match1[team] + ' 信息++++++')
#     tem = '{0}_order_list'.format(team)     # 获取单号表
#     tem2 = '{0}_order_list'.format(team)    # 更新单号表
#     for i in range((end - begin).days):  # 按天循环获取订单状态
#         day = begin + datetime.timedelta(days=i)
#         yesterday = str(day) + ' 23:59:59'
#         last_month = str(day)
#         print('正在更新 ' + match1[team] + last_month + ' 号订单信息…………')
#         searchType = '订单号'      # 运单号，订单号   查询切换
#         sso.orderInfo(searchType, tem, tem2, last_month)
#     print('更新耗时：', datetime.datetime.now() - start)
#
#     print('------------导出部分：---------------------')
#     # m.connectOrder(team, month_last, month_yesterday, month_begin)  # 最近两个月的订单信息导出
#     qu.EportOrder(team, month_last, month_yesterday, month_begin)     # 最近两个月的更新信息导出
#     print('输出耗时：', datetime.datetime.now() - start)





# 输出签收率表、(备用)
# tem = '泰国'
# w.OrderQuan(team, tem)
# print('导出耗时：', datetime.datetime.now() - start)


'''
    IDE很多技巧:
    1,  `ctrl + alt + L`，格式化代码
    2,  双击`shift`搜索一切，不管是IDE功能、文件、方法、变量……都能搜索
    3,  `alt+enter`万能键
    4,  `shift+enter`向下换行
    5,  `shift+ctrl`向上换行
    6,  `ctrl+space` 万能提示键，PyCharm的会根据上下文提供补全
    7,  `ctrl+shift+f10`运行当前文件
    8,  `ctrl+w`扩展选取和`ctrl+shift+w`缩减选区, `ctrl+alt+shift+T`重构选区
    9,  `ctrl+q`查注释
'''