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
		data = {'username': self.userName,
				'password': self.password,
				'remember': '1'}
		r_header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36'}
		req = self.session.post(url=url, headers=r_header, data=data)
		print('------  成功登陆系统后台  -------')

	def newOrderInfo(self, orderId, searchType):                  # 进入老后台查询界面
		url = 'https://goms.giikin.com/admin/order/orderquery.html'
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
		print('-------已成功发送请求++++++')
		orderInfo = self.new_parseDate(req)   			# 获取订单简单信息
		# print(orderInfo)
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
		try:
			self.q.put(orderInfo)
		except Exception as e:
			print('放入失败---：', str(Exception) + str(e))
		return orderInfo

	def newUser(self, team, searchType, last_month): # ----主线程的执行（多线程函数）
		match = {'slgat': '港台',
				'sltg': '泰国',
				'slxmt': '新马',
				'slrb': '日本'}
		match2 = {'slgat': '台币',
				'sltg': '泰铢',
				'slxmt': '新加坡（英语）',
				'slrb': '日币'}
		match3 = {'slgat': '台湾',
				'sltg': '泰国',
				'slxmt': '新加坡',
				'slrb': '日本'}
		match4 = {'slgat': '台湾',
				'sltg': '泰国',
				'slxmt': '马来西亚（繁体）',
				'slrb': '日本'}
		match5 = {'slgat': '台湾',
				'sltg': '泰国',
				'slxmt': '马来西亚',
				'slrb': '日本'}
		print("======== 开始订单详情查询 ======")
		month_begin = (datetime.datetime.now() - relativedelta(months=4)).strftime('%Y-%m-%d')
		start = datetime.datetime.now()
		sql = '''SELECT id,`订单编号`  FROM {0}_order_list sl  where sl.日期 ='{1}';'''.format(team, last_month)
		ordersDict = pd.read_sql_query(sql=sql, con=self.engine1)
		if ordersDict.empty:
			print('无需要更新的产品id信息！！！')
			return
		print(ordersDict)
		ordersDict = ordersDict['订单编号'].values.tolist()
		print('获取耗时：', datetime.datetime.now() - start)
		print('主线程开始执行……………………')
		threads = []  				 # 多线程用线程池--
		for order in ordersDict:     # 注意前后数组的取值长度一致
			threads.append(Thread(target=self.newOrderInfo, args=(order, searchType)))    #  -----也即是子线程
		print('子线程分配完成++++++')
		if threads:                  # 当所有的线程都分配完成之后，通过调用每个线程的start()方法再让他们开始。
			print(len(threads))
			for th in threads:
				th.start()           # print ("开启子线程…………")
			for th in threads:
				th.join()            # print ("退出子线程")
		else:
			print("没有需要运行子线程！！！")
		print('主线程执行结束---------')
		results = []
		for i in range(len(ordersDict)):   # print(i)
			try:
				results.append(self.q.get())
			except Exception as e:
				print('取出失败---：', str(Exception) + str(e))
		print('-----执行结束---------')
		print('         V           ')
		pf = pd.DataFrame(list(results))  # 将字典列表转换为DataFrame
		pf = pf[['订单号', '订单状态', '物流单号', '商品名称', '是否二次改派', '数量', '币种', '下单时间', '电话', '物流状态']]
		print('正在写入缓存中......')
		try:
			pf.to_sql('备用', con=self.engine1, index=False, if_exists='replace')
			sql = '''update {0}_order_list a, 备用 b
						set a.`币种`=IF(IF(b.`币种` = '{1}', '{2}', b.`币种`) = '{3}', '{4}', b.`币种`),
							a.`数量`=b.`数量`,
							a.`电话号码`=b.`电话` ,
							a.`运单编号`=b.`物流单号`,
							a.`系统订单状态`=b.`订单状态`,	
							a.`系统物流状态`=b.`物流状态`,
							a.`是否改派`=IF(b.`是否二次改派`='二次改派', '改派', b.`是否二次改派`)
					where a.`订单编号`=b.`订单号`;'''.format(team, match2[team], match3[team], match4[team], match5[team])
			pd.read_sql_query(sql=sql, con=self.engine1, chunksize=1000)
		except Exception as e:
			print('更新失败：', str(Exception) + str(e))
		print('更新成功…………')
		return ordersDict

if __name__ == '__main__':                    # 以老后台的简单查询为主，
	start = datetime.datetime.now()
	# s = BpsControl99('qiyuanzhang@jikeyin.com', 'qiyuanzhang123.0')
	# s = BpsControl99('gupeiyu@giikin.com', 'gu19931209*')
	s = BpsControl99('zhangjing@giikin.com', 'Giao2020..0')
	begin = datetime.date(2021, 2, 1)
	print(begin)
	end = datetime.date(2021, 2, 2)
	print(end)
	for i in range((end - begin).days):  		# 按天循环获取订单状态
		day = begin + datetime.timedelta(days=i)
		yesterday = str(day) + ' 23:59:59'
		last_month = str(day)
		print('正在更新 ' + last_month + ' 号订单信息…………')
		team = 'slxmt'
		searchType = '订单号'  					# 运单号，订单号   查询切换
		s.newUser(team, searchType, last_month)