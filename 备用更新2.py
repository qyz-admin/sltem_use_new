import requests
from bs4 import BeautifulSoup # 抓标签里面元素的方法
import os
import xlwings
import pandas as pd
import datetime
import time

from dateutil.relativedelta import relativedelta
from settings import Settings
from sqlalchemy import create_engine
from queue import Queue
from threading import Thread #  使用 threading 模块创建线程
class BpsControl99(Settings):
	def __init__(self, userName, password):
		Settings.__init__(self)
		self.userName = userName
		self.password = password
		self.session = requests.session()   #	实例化session，维持会话,可以让我们在跨请求时保存某些参数
		self.__load()
		self.q = Queue()    				# 多线程调用的函数不能用return返回值，用来保存返回值
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
		self.engine3 = create_engine('mysql+mysqlconnector://{}:{}@{}:{}/{}'.format(self.mysql3['user'],
																					self.mysql3['password'],
																					self.mysql3['host'],
																					self.mysql3['port'],
																					self.mysql3['datebase']))
	def __load(self):  # 登录系统保持会话状态
		url = r'https://goms.giikin.com/admin/login/index.html'
		# url = r'http://gsso.giikin.com/admin/login/index.html'
		data = {'username': self.userName,
				'password': self.password,
				'remember': '1'}
		r_header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36'}
		req = self.session.post(url=url, headers=r_header, data=data)
		# print(req)
		# print(req.text)
		print('------  成功登陆系统后台  -------')

	def newOrderInfo(self, orderId, searchType):                  # 进入老后台查询界面
		url = 'https://goms.giikin.com/admin/order/orderquery.html'
		# url = 'http://gimp.giikin.com/service?service=gorder.customer&action=getOrderList'
		data = {'phone': None,
				'ship_email': None,
				'ip': None}
		if searchType == '订单号':
			data.update({'order_number': orderId,
						'waybill_number': None})
		elif searchType == '运单号':
			data.update({'order_number': None,
						'waybill_number': orderId})
		req = self.session.post(url=url, data=data)
		print(req)
		# print(req.text)
		print('-------已成功发送请求++++++')
		orderInfo = self.new_parseDate(req)   			# 获取订单简单信息
		# return orderInfo

	def new_parseDate(self, req):  					# 对返回的response 进行处理； 处理订单简单信息
		soup = BeautifulSoup(req.text, 'lxml') 		# 创建 beautifulsoup 对象
		orderInfo = {}
		labels = soup.find_all('th')   				# 获取行标签的th值
		vals = soup.find_all('td')     				# 获取表格的td的值
		if len(labels) > len(vals) or len(labels) < len(vals):
			print('查询失败！！！')
		else:
			for i in range(len(labels)):
				orderInfo[str(labels[i]).replace("<th>", "").replace("</th>", "").strip()] = str(vals[i]).replace("<td>", "").replace("</td>", "").strip()
		print('-------已处理订单简单信息---------')
		orderInfo = pd.json_normalize(orderInfo)
		# print(orderInfo)
		orderInfo = orderInfo[['订单号', '姓名']]
		print(orderInfo)
		if orderInfo is not None and len(orderInfo) > 0:
			# orderInfo.to_sql('备用', con=self.engine1, index=False, if_exists='replace')
			# sql = '''update sheet1_iphone a, 备用 b set a.`姓名`=b.`姓名` where a.`订单编号`=b.`订单号`;'''
			sql = '''update sheet1_iphone a set a.`姓名`='{1}' where a.`订单编号`='{0}';'''.format(orderInfo['订单号'][0], orderInfo['姓名'][0])
			pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
		return orderInfo


	def newUser(self, team, searchType):  # ----主线程的执行（多线程函数）
		print("======== 开始订单详情查询 ======")
		start = datetime.datetime.now()
		rq = datetime.datetime.now().strftime('%Y%m%d.%H%M%S')
		sql = '''SELECT `订单编号`  FROM sheet1_iphone;'''
		ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
		if ordersDict.empty:
			print('无需要更新的信息！！！')
			return
		print(ordersDict)
		ordersDict = ordersDict['订单编号'].values.tolist()
		print('获取耗时：', datetime.datetime.now() - start)
		df = self.newOrderInfo(ordersDict[0], '订单号')
		dlist = []
		max_count = len(ordersDict)
		for order in ordersDict:
			data = self.newOrderInfo(order, '订单号')
			# dlist.append(data)
		# dp = df.append(dlist, ignore_index=True)
		# dp.to_excel('G:\\输出文件\\产品检索-查询{}.xlsx'.format(rq), sheet_name='查询', index=False, engine='xlsxwriter')


if __name__ == '__main__':                    # 以老后台的简单查询为主，
	start = datetime.datetime.now()
	# s = BpsControl99('qiyuanzhang@jikeyin.com', 'qiyuanzhang123.')  老后台密码
	s = BpsControl99('qiyuanzhang@jikeyin.com', 'qiyuanzhang123.0') #新后台密码
	# s = BpsControl99('gupeiyu@giikin.com', 'gu19931209*')
	# s = BpsControl99('zhangjing@giikin.com', 'Giao2020..0')
	# s.newOrderInfo('MY102010000268232', '订单号')
	team = 'slxmt'
	searchType = '订单号'
	s.newUser(team, searchType)
