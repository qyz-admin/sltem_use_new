import pandas as pd
import os
import datetime
import time
import xlwings
import requests
import json
from queue import Queue
from settings_sso import Settings_sso
from apscheduler.schedulers.blocking import BlockingScheduler

# -*- coding:utf-8 -*-
class QueryTwo(Settings_sso):
    def __init__(self, userMobile, password):
        Settings_sso.__init__(self)
        self.session = requests.session()  # 实例化session，维持会话,可以让我们在跨请求时保存某些参数
        self.q = Queue()  # 多线程调用的函数不能用return返回值，用来保存返回值
        self.userMobile = userMobile
        self.password = password
        # self.sso__online_auto()

    #  物流分配1-单点系统（一）
    def doSaveRuleStatus(self, wayid, isReal):  # 进入 物流分配1 界面
        print('设置中......')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.logistics&action=doSaveRuleStatus'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/logisticsTrajectory'}
        data = {'id': wayid,
                'status': isReal
                }
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)          # json类型数据转换为dict字典
        print(req)
        check = req['data']['success']
        if check == 'true':
            print(check)
        print('已经设置+++')
        return check

    def check_doSaveRuleStatus(self, wayid):  # 进入 物流分配1 界面
        print('正在检查中......')
        rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
        url = r'https://gimp.giikin.com/service?service=gorder.logistics&action=getRuleList'
        r_header = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
                    'origin': 'https: // gimp.giikin.com',
                    'Referer': 'https://gimp.giikin.com/front/logisticsTrajectory'}
        data = {'page': 1,
                'pageSize': 10,
                'id': wayid,
                'name': '',
                'area_id': '',
                'country_code': '',
                'logistics_id': '',
                'logistics_style': '',
                'status': '',
                'currency_id': ''
                }
        proxy = '39.105.167.0:40005'  # 使用代理服务器
        proxies = {'http': 'socks5://' + proxy,
                   'https': 'socks5://' + proxy}
        # req = self.session.post(url=url, headers=r_header, data=data, proxies=proxies)
        req = self.session.post(url=url, headers=r_header, data=data)
        # print('+++已成功发送请求......')
        req = json.loads(req.text)          # json类型数据转换为dict字典
        # print(req)
        check = req['data']['list'][0]['status']
        if check == 0:
            print('设置失败')
        elif check == 1:
            print('设置成功')
        return check

    def main(self):
        self.sso__online_auto()
        # for wayid in [451, 452, 428, 429]:
        for wayid in [451, 452, 641]:
            print('正在打开：' + str(wayid) + ' :' + match[wayid] + ' 的物流分配规则')
            isReal = 1
            m.doSaveRuleStatus(wayid, isReal)  # 导入；，更新--->>数据更新切换
            m.check_doSaveRuleStatus(wayid)

if __name__ == '__main__':
    m = QueryTwo('+86-18538110674', 'qyz04163510.')
    start: datetime = datetime.datetime.now()
    match = {451: '台湾-铱熙无敌-新竹普货',
             452: '台湾-铱熙无敌-新竹特货',
             428: '台湾-铱熙无敌-新竹普货',
             429: '台湾-铱熙无敌-新竹特货',
             641: '台湾-铱熙无敌-711'}
    '''
    # -----------------------------------------------手动导入状态运行（一）-----------------------------------------
    # 1、 正在按订单查询；2、正在按时间查询；--->>数据更新切换
    # isReal: 0 查询后台保存的运单轨迹； 1 查询物流的实时运单轨迹 
    '''
    print('开始：' + str(start))
    sched = BlockingScheduler()
    sched.add_job(m.main, 'cron', hour=23, minute=59)
    # sched.add_job(m.my_job2, 'interval', seconds=10, misfire_grace_time=10,id = 'my_job')
    sched.start()
    print('结束：' + str(datetime.datetime.now()))

    # for wayid in [451, 452, 428, 429]:
    # for wayid in [451, 452]:
    #     print('正在打开：' + str(wayid)+ ' :' + match[wayid] + ' 的物流分配规则')
    #     isReal = 1
    #     m.doSaveRuleStatus(wayid,isReal)       # 导入；，更新--->>数据更新切换
    #     m.check_doSaveRuleStatus(wayid)

    # m.order_bind_status('2022-01-01', '2022-01-02')
    # m._order_bind_status('7449201841')

    print('查询耗时：', datetime.datetime.now() - start)