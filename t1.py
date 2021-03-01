import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from excelControl import ExcelControl
from mysqlControl import MysqlControl
from wlMysql import WlMysql
from wlExecl import WlExecl
# from orderQuery import OrderQuery
import datetime
start: datetime = datetime.datetime.now()
team = 'slxmt'
match = {'slrb': r'D:\Users\Administrator\Desktop\需要用到的文件\日本签收表',
         'sltg': r'D:\Users\Administrator\Desktop\需要用到的文件\泰国签收表',
         'slgat': r'D:\Users\Administrator\Desktop\需要用到的文件\港台签收表',
         'slxmt': r'D:\Users\Administrator\Desktop\需要用到的文件\新马签收表'}
'''    msyql 语法:      show processlist;
备注：  港台 需整理的表：香港立邦>(明细再copy一份保存) ； 台湾龟山改派>(copy保存为xlsx格式);
说明：  日本 需整理的表：1、吉客印神龙直发签收表=密码：‘JKTSL’>(明细再copy保存；改派明细不需要);2、直发签收表>(明细再copy保存；3、状态更新需要copy保存);
'''
path = match[team]
dirs = os.listdir(path=path)
e = ExcelControl()
m = MysqlControl()
w = WlMysql()
we = WlExecl()
# qo = OrderQuery()
# 上传退货
e.readReturnOrder(team)
print('退货导入耗时：', datetime.datetime.now() - start)

# ---读取execl文件---
for dir in dirs:
    filePath = os.path.join(path, dir)
    print(filePath)
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
m.creatMyOrderSl(team)      # 最近五天的全部订单信息
print('------------更新部分：---------------------')
# m.creatMyOrderSlTWO(team)   # 最近两个月的更新订单信息
print('处理耗时：', datetime.datetime.now() - start)
# m.connectOrder(team)      # 最近两个月的订单信息导出
print('输出耗时：', datetime.datetime.now() - start)





# ---数据库分段读取---
# m.creatMyOrder(team)   # 备用获取最近两个月全部订单信息
# 输出签收率表、(备用)
# tem = '泰国'
# w.OrderQuan(team, tem)
# print('导出耗时：', datetime.datetime.now() - start)
