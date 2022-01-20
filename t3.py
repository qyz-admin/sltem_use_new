import os
import win32api,win32con
import win32com.client as win32
from openpyxl import Workbook, load_workbook
from excelControl import ExcelControl
from mysqlControl import MysqlControl
from wlMysql import WlMysql
from wlExecl import WlExecl
from sso_updata import QueryTwo
from gat_update2 import QueryUpdate
import datetime
from dateutil.relativedelta import relativedelta
start: datetime = datetime.datetime.now()
team = 'gat'
match1 = {'gat': '港台',
          'slsc': '品牌'}
match = {'sl_rb': r'D:\Users\Administrator\Desktop\需要用到的文件\A日本签收表',
         'slsc': r'D:\Users\Administrator\Desktop\需要用到的文件\品牌',
         'gat': r'D:\Users\Administrator\Desktop\需要用到的文件\A港台签收表',
         'gat_upload': r'D:\Users\Administrator\Desktop\需要用到的文件\A港台签收表 - 单独上传'}
'''    msyql 语法:      show processlist（查看当前进程）;  
                        set global event_scheduler=0;（关闭定时器）;
备注：  港台 需整理的表：香港立邦>(明细再copy一份保存) ； 台湾龟山改派>(copy保存为xlsx格式);
说明：  日本 需整理的表：1、吉客印神龙直发签收表=密码：‘JKTSL’>(明细再copy保存；改派明细不需要);2、直发签收表>(明细再copy保存；3、状态更新需要copy保存);
'''
# 初始化时间设置
if team in ('slsc', 'slrb', 'slrb_jl', 'slrb_js', 'slrb_hs', 'ga9t', 'sl_rb'):
    # 更新时间
    yy = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y'))
    mm = int((datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%m'))
    begin = datetime.date(yy, mm, 1)
    yy2 = int(datetime.datetime.now().strftime('%Y'))
    mm2 = int(datetime.datetime.now().strftime('%m'))
    dd2 = int(datetime.datetime.now().strftime('%d'))
    end = datetime.date(yy2, mm2, dd2)
    # 导出时间
    month_last = (datetime.datetime.now().replace(day=1) - datetime.timedelta(days=1)).strftime('%Y-%m') + '-01'
    month_yesterday = datetime.datetime.now().strftime('%Y-%m-%d')
    month_begin = (datetime.datetime.now() - relativedelta(months=3)).strftime('%Y-%m-%d')
else:
    # 更新时间
    begin = datetime.date(2021, 11, 1)
    end = datetime.date(2021, 12, 7)
    # 导出时间
    month_last = '2021-11-01'
    month_yesterday = '2021-12-07'
    month_begin = '2021-10-01'
print('****** 更新起止时间：' + begin.strftime('%Y-%m-%d') + ' - ' + end.strftime('%Y-%m-%d') + ' ******')
print('****** 导出起止时间：' + begin.strftime('%Y-%m-%d') + ' - ' + end.strftime('%Y-%m-%d') + ' ******')
# 库的引用
path = match[team]
dirs = os.listdir(path=path)
e = ExcelControl()
m = MysqlControl()
w = WlMysql()
we = WlExecl()
qu = QueryUpdate()

# 上传退货
e.readReturnOrder(team)
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

# TODO---数据库分段读取---
m.creatMyOrderSlTWO(team, begin, end)   # 最近两个月的 部分内容 更新信息
m.creatMyOrderSl(team)  # 最近五天的全部订单信息
print('获取-更新 耗时：', datetime.datetime.now() - start)

print('------------更新部分：---------------------')
if team in ('ga99t', 'slsc', 'slrb', 'slrb_jl', 'slrb_js', 'slrb_hs', 'sl_rb'):
    m.creatMyOrderSlTWO(team, begin, end)   # 最近两个月的更新订单信息
    print('处理耗时：', datetime.datetime.now() - start)

    print('------------导出部分：---------------------')
    m.connectOrder(team, month_last, month_yesterday, month_begin)  # 最近两个月的订单信息导出
    print('输出耗时：', datetime.datetime.now() - start)

elif team in ('gat'):
    sso = QueryTwo('+86-18538110674', 'qyz04163510', '1343')
    print(datetime.datetime.now())
    print('++++++正在获取 ' + match1[team] + ' 信息++++++')
    tem = '{0}_order_list'.format(team)     # 获取单号表
    tem2 = '{0}_order_list'.format(team)    # 更新单号表
    searchType = '订单号'
    print('正在更新 >>>昨日<<< 的最新订单信息......')
    last_month = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
    now_month = (datetime.datetime.now()).strftime('%Y-%m-%d')
    sso.orderInfo_th(searchType, tem, tem2, last_month, now_month)

    for i in range((end - begin).days):  # 按天循环获取订单状态
        print(i)
        last_month = begin + datetime.timedelta(days=5 * i)
        now_month = begin + datetime.timedelta(days=(i+1) * 5)
        if end >= now_month:
            print('正在更新 ' + str(last_month) + ' 号 --- ' + str(now_month) + ' 号信息…………')
            sso.orderInfo_th(searchType, tem, tem2, last_month, now_month)
        else:
            now_month = last_month + datetime.timedelta(days=(end - last_month).days)
            print('正在更新 ' + str(last_month) + ' 号 --- ' + str(now_month) + ' 号信息…………')
            sso.orderInfo_th(searchType, tem, tem2, last_month, now_month)
            break
    print('更新耗时：', datetime.datetime.now() - start)

    print('------------导出部分：---------------------')
    # m.connectOrder(team, month_last, month_yesterday, month_begin)  # 最近两个月的订单信息导出
    qu.EportOrder(team, month_last, month_yesterday, month_begin)     # 最近两个月的更新信息导出
    print('输出耗时：', datetime.datetime.now() - start)




win32api.MessageBox(0, "注意:>>>    程序运行结束， 请查看表  ！！！", "提 醒",win32con.MB_OK)
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