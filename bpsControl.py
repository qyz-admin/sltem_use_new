import requests
from bs4 import BeautifulSoup # 抓标签里面元素的方法
import os
import xlwings
import pandas as pd
import datetime
import time

# from mysqlControl import MysqlControl
from settings import Settings
from sqlalchemy import create_engine
from queue import Queue
from threading import Thread #  使用 threading 模块创建线程
class BpsControl(Settings):
	def __init__(self, userName, password):
		Settings.__init__(self)
		self.userName = userName
		self.password = password
		self.session = requests.session()  #实例化session
		self.__load()
		self.q = Queue()    # 多线程调用的函数不能用return返回值，用来保存返回值
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
		# requests.session():维持会话,可以让我们在跨请求时保存某些参数
		r_header = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.83 Safari/537.36'}
		r = self.session.post(url=url, headers=r_header, data=data)
		print('------  成功登陆系统后台  -------')
		print('               v           ')
	def getOrderInfo(self, orderId, searchType):                  # 进入查询界面
		url = 'https://goms.giikin.com/admin/order/orderquery.html'
		data = {'phone': None,
			'ship_email': None,
			'ip': None
				}
		if searchType == '订单号':
			data.update({'order_number': orderId,
						'waybill_number': None})
		elif searchType == '运单号':
			data.update({'order_number': None,
						'waybill_number': orderId})
		req = self.session.post(url=url, data=data)
		# print('-------已成功发送请求++++++')
		orderInfo = self._parseDate(req)   			# 获取订单简单信息
		# print(orderInfo)
		return orderInfo
	def _parseDate(self, req):  					# 对返回的response 进行处理
		# print('-------正在处理订单简单信息---------')
		soup = BeautifulSoup(req.text, 'lxml') 		# 创建 beautifulsoup 对象
		orderInfo = {}
		# print(soup)
		# print(soup.a['href'])
		labels = soup.find_all('th')   # 获取行标签的th值
		vals = soup.find_all('td')     # 获取表格的td的值
		# print('-------正在获取查询值..........')
		# print(labels)
		# print(vals)
		if len(labels) > len(vals) or len(labels) < len(vals):
			print('查询失败！！！')
		else:
			for i in range(len(labels)):
				orderInfo[str(labels[i]).replace("<th>", "").replace("</th>", "").strip()] = str(vals[i]).replace("<td>", "").replace("</td>", "").strip()
		# print('-------已处理订单简单信息---------')
		try:
			self.q.put(orderInfo)
		except Exception as e:
			print('放入失败---：', str(Exception) + str(e))
		# print(orderInfo)
		return orderInfo
	def getNumberT(self, team, searchType): # ----主线程的执行（多线程函数）
		match = {'slgat': '港台',
				'sltg': '泰国',
				'slxmt': '新马',
				'slzb': '直播团队',
				'slyn': '越南',
				'slrb': '日本'}
		print("========开始第一阶段查询（近6天）======")
		now_yesterday = (datetime.datetime.now()).strftime('%Y-%m-%d') + ' 23:59:59'
		last_yesterday = (datetime.datetime.now() - datetime.timedelta(days=4)).strftime('%Y-%m-%d') + ' 00:00:00'
		print(now_yesterday)
		print(last_yesterday)
		print('-----正在获取工作表的订单编号++++++')
		start = datetime.datetime.now()
		sql = '''SELECT order_number FROM 全部订单_{0} WHERE 全部订单_{0}.addtime>= '{1}' AND 全部订单_{0}.addtime<= '{2}';'''.format(team, last_yesterday, now_yesterday)
		ordersDict = pd.read_sql_query(sql=sql, con=self.engine3)
		print(ordersDict)
		ordersDict = ordersDict['order_number'].values.tolist()
		# print(ordersDict)
		print('获取耗时：', datetime.datetime.now() - start)
		print('------正在查询单个订单的详情++++++')
		print('主线程开始执行……………………')
		threads = []  # 多线程用线程池--
		for order in ordersDict:     # 注意前后数组的取值长度一致
			# print (order)   # print (ordersDict)
			threads.append(Thread(target=self.getOrderInfo, args=(order, searchType)))    #  -----也即是子线程
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
		print(self.q.qsize())
		print(self.q.empty())
		print(self.q.full())
		for i in range(len(ordersDict)):   # print(i)
			try:
				results.append(self.q.get())
			except Exception as e:
				print('取出失败---：', str(Exception) + str(e))
		print('-----执行结束---------')
		print('         V           ')
		# print(results)
		# pf = pd.DataFrame(list(results))  # 将字典列表转换为DataFrame
		pf = pd.DataFrame(results)
		pf.insert(0, '應付金額', '')
		pf.insert(0, '支付方式', '')
		pf.rename(columns={'规格': '规格中文'}, inplace=True)
		pf.dropna(subset=['订单号'],inplace=True)
		pf = pf[['订单号', '订单状态', '物流单号', '下单时间', '币种', '物流状态', '應付金額', '支付方式', '规格中文']]
		pf = pf.loc[:, ['订单号', '订单状态', '物流单号', '下单时间', '币种', '物流状态', '應付金額', '支付方式', '规格中文']]
		try:
			print('正在写入缓存表中…………')
			pf.to_sql('规格缓存_sltg', con=self.engine3, index=False, if_exists='replace')
			print('正在写入总订单表中…………')
			sql = 'REPLACE INTO 全部订单规格_sltg SELECT *, NOW() 添加时间  FROM 规格缓存_sltg;'
			#  sql = 'UPDATE 全部订单_sltg r INNER JOIN (SELECT 订单号,规格中文 FROM 规格缓存_sltg) t ON r.order_number= t.`订单号` SET r.op_id = t.`规格中文`;'
			pd.read_sql_query(sql=sql, con=self.engine3, chunksize=100)
		except Exception as e:
			print('缓存---：', str(Exception) + str(e))
		pf = pf.astype(str)   # dataframe的类型为dtype: object无法导入mysql中，需要转换为str类型
		print('------写入成功------')
		today = datetime.date.today().strftime('%Y.%m.%d')
		pf.to_excel('F:\\查询\\查询输出\\{} {} 订单查询.xlsx'.format(today, match[team]),
					sheet_name=match[team], index=False)
		print('------输出文件成功------')
		return ordersDict
	def getNumberAdd(self, team, searchType):   # ----主线程的执行（多线程函数）
		match = {'slgat': '港台',
				'sltg': '泰国',
				'slxmt': '新马',
				'slzb': '直播团队',
				'slyn': '越南',
				'slrb': '日本'}
		print("========开始第二阶段查询（补充）======")
		now_yesterday = (datetime.datetime.now() - datetime.timedelta(days=5)).strftime('%Y-%m-%d') + ' 00:00:00'
		last_yesterday = (datetime.datetime.now() - datetime.timedelta(days=8)).strftime('%Y-%m-%d') + ' 00:00:00'
		print('-------正在获取工作表的订单编号++++++')
		start = datetime.datetime.now()
		sql = '''SELECT order_number FROM 全部订单_{0} WHERE 全部订单_{0}.op_id= '' And 全部订单_{0}.addtime>= '{1}' AND 全部订单_{0}.addtime<= '{2}';'''.format(team, last_yesterday, now_yesterday)
		ordersDict = pd.read_sql_query(sql=sql, con=self.engine3)
		print(ordersDict)
		ordersDict = ordersDict['order_number'].values.tolist()
		# print(ordersDict)
		print('获取耗时：', datetime.datetime.now() - start)
		print('--------正在查询单个订单的详情++++++')
		print('主线程开始执行……………………')
		threads = []  # 多线程用线程池--
		for order in ordersDict:     # 注意前后数组的取值长度一致
			# print (order)   # print (ordersDict)
			threads.append(Thread(target=self.getOrderInfo, args=(order, searchType)))    #  -----也即是子线程
		print('子线程分配完成++++++')
		if threads:                  # 当所有的线程都分配完成之后，通过调用每个线程的start()方法再让他们开始。
			print(len(threads))
			for th in threads:
				th.start()           # print ("开启子线程…………")
			for th in threads:
				th.join()            # print ("退出子线程")
		else:
			print("没有需要运行子线程！！！")
		print('子线程运行结束---------')
		results = []
		print(self.q.qsize())
		print(self.q.empty())
		print(self.q.full())
		for i in range(len(ordersDict)):    # print(i)
			try:
				results.append(self.q.get())
			except Exception as e:
				print('取出失败---：', str(Exception) + str(e))
		print('-----订单获取执行结束---------')
		print('         V           ')
		# print(results)
		# pf = pd.DataFrame(list(results))  # 将字典列表转换为DataFrame
		pf = pd.DataFrame(results)
		pf.insert(0, '應付金額', '')
		pf.insert(0, '支付方式', '')
		pf.rename(columns={'规格': '规格中文'}, inplace=True)
		pf.dropna(subset=['订单号'],inplace=True)
		pf = pf[['订单号', '订单状态', '物流单号', '下单时间', '币种', '物流状态', '應付金額', '支付方式', '规格中文']]
		pf = pf.loc[:, ['订单号', '订单状态', '物流单号', '下单时间', '币种', '物流状态', '應付金額', '支付方式', '规格中文']]
		try:
			print('正在写入缓存表中…………')
			pf.to_sql('规格缓存_sltg', con=self.engine3, index=False, if_exists='replace')
			sql = 'REPLACE INTO 全部订单规格_sltg SELECT *, NOW() 添加时间  FROM 规格缓存_sltg;'
			# sql = 'UPDATE 全部订单_sltg r INNER JOIN (SELECT 订单号,规格中文 FROM 规格缓存_sltg) t ON r.order_number= t.`订单号` SET r.op_id = t.`规格中文`;'
			pd.read_sql_query(sql=sql, con=self.engine3, chunksize=100)
		except Exception as e:
			print('缓存---：', str(Exception) + str(e))
		pf = pf.astype(str)   # dataframe的类型为dtype: object无法导入mysql中，需要转换为str类型
		print('------缓存成功------')
		today = datetime.date.today().strftime('%Y.%m.%d')
		pf.to_excel('F:\\查询\\查询输出\\{} {} 订单补充查询.xlsx'.format(today, match[team]),
					sheet_name=match[team], index=False)
		print('------输出文件成功------')
		return ordersDict

	# 获取泰国海外仓
	def sltg_HaiWaiCang(self, house):
		match = {'shifeng': 'Tracking Number',
				'chaoshidai': '运单号',
				'bojiatu': '上架单号', }
		match1 = {'shifeng': '海外仓库存_时丰',
				'chaoshidai': '海外仓库存_超时代在库',
				'bojiatu': '海外仓库存_博佳图', }
		today = datetime.date.today().strftime('%Y.%m.%d')
		sql = '''SELECT 订单编号,
						运单号,
						产品id,
						产品名称,
						规格中文,
						数量,
						qb.订单状态 
				FROM
					(SELECT 
						a.order_number '订单编号', 
						a.waybill_number '运单号',
						a.goods_id '产品id',
						a.goods_name '产品名称',
						a.op_id '规格',
						a.quantity '数量',
						a.order_status '订单状态' 
					FROM 
						全部订单_sltg a 
					INNER JOIN 
						(SELECT DISTINCT 
							upper(`{0}`) 'Tracking Number'
						FROM {1}) b 
						ON a.waybill_number = b.`Tracking Number`) qb 
					INNER JOIN 全部订单规格_sltg b 
						ON qb.`订单编号` = b.`订单号`;'''.format(match[house], match1[house])
		print('正在查询' + match1[house] + '订单…………')
		df = pd.read_sql_query(sql=sql, con=self.engine3)
		# print(df)
		print('正在写入excel…………')
		df.to_excel('D:\\Users\\Administrator\\Desktop\\查询\\{} {}.xlsx'.format(today, match1[house]),
					sheet_name=match1[house], index=False)
		print('输出文件成功…………')

	# 各团队(泰国)全部订单表-函数（停用）
	def tgOrderQuan(self, team):  # 3天内的
		match1 = {'slgat': '港台',
			  	'sltg': '泰国',
			  	'slxmt': '新马',
			  	'slzb': '直播团队',
			  	'slyn': '越南',
			  	'slrb': '日本'}
		match = {'slgat': '"神龙家族-港澳台"',
			 	'sltg': '"神龙家族-泰国"',
			 	'slxmt': '"神龙家族-新加坡", "神龙家族-马来西亚", "神龙家族-菲律宾"',
			 	'slzb': '"神龙家族-直播团队"',
			 	'slyn': '"神龙家族-越南"',
			 	'slrb': '"神龙家族-日本团队"'}
		print('正在获取' + match1[team] + '最近 10 天订单…………')
		yesterday = (datetime.datetime.now()).strftime('%Y-%m-%d')
		# yesterday = (datetime.datetime.now().replace(month=1, day=10)).strftime('%Y-%m-%d')
		print(yesterday)
		last_month = (datetime.datetime.now() - datetime.timedelta(days=10)).strftime('%Y-%m-%d')
		# last_month = (datetime.datetime.now().replace(month=1, day=5)).strftime('%Y-%m-%d')
		print(last_month)
		sql = '''SELECT a.id,
                    a.订单编号 order_number,
                    a.团队 area_id,
                    '' main_id,
                    a.电话号码 ship_phone,
                    a.邮编 ship_zip,
                    a.价格 amount,
                    a.系统订单状态 order_status,
                    UPPER(a.运单编号) waybill_number,
                    a.付款方式 pay_type,
                    a.下单时间 addtime,
                    a.审核时间 update_time,
                    a.产品id goods_id, 
                    '' quantity,
                    a.物流方式 logistics_id,
                    '' op_id,
                    CONCAT(a.产品id,'#' ,a.产品名称) goods_name, 
                    a.是否改派 secondsend_status,
                    a.是否低价 low_price
            FROM {}_order_list a 
            WHERE a.日期 >= '{}' AND a.日期 <= '{}';'''.format(team, last_month, yesterday)
		try:
			df = pd.read_sql_query(sql=sql, con=self.engine1)
			print(df)
			print('正在写入缓存表中…………')
			df.to_sql('tem_sl', con=self.engine3, index=False, if_exists='replace')
			print('++++更新缓存完成++++')
		except Exception as e:
			print('更新缓存失败：', str(Exception) + str(e))
		print('正在写入 ' + match1[team] + ' 全部订单表中…………')
		sql = 'REPLACE INTO 全部订单_{} SELECT *, NOW() 添加时间 FROM tem_sl;'.format(team)
		pd.read_sql_query(sql=sql, con=self.engine3, chunksize=100)
		# 获取订单明细（泰国）
		print('======正在启动查询订单程序>>>>>')
		b = BpsControl('gupeiyu@giikin.com', 'gu19931209*')
		match = {'slgat': '港台',
			 'sltg': '泰国',
			 'slxmt': '新马',
			 'slzb': '直播团队',
			 'slyn': '越南',
			 'slrb': '日本'}
		team = 'sltg'
		searchType = '订单号'  # 运单号，订单号；查询切换
		b.getNumberT(team, searchType)
		print('查询耗时：', datetime.datetime.now() - start)
		time.sleep(10)
		b.getNumberAdd(team, searchType)
		print('补充耗时：', datetime.datetime.now() - start)
	# 修改样式（备用）
	def xiugaiyangshi(self, filePath, sheetname):
		print('正在修改样式…………')
		wb = load_workbook(filePath)
		print(wb.sheetnames)
		sheet = wb[sheetname]
		for i in range(4, 5):
			for j in range(2, sheet.max_row):
				if sheet.cell(j, i).value == '合计' and sheet.cell(j, i + 1).value == '合计' and sheet.cell(j,
																										i + 2).value == '合计':
					for c in range(1, sheet.max_column + 1):
						sheet.cell(j, c).fill = PatternFill(patternType='solid', fgColor='1874CD')
		for i in range(5, 6):
			for j in range(2, sheet.max_row):
				if sheet.cell(j, i - 1).value != '合计' and sheet.cell(j, i).value == '合计':
					for c in range(1, sheet.max_column + 1):
						sheet.cell(j, c).fill = PatternFill(patternType='solid', start_color='FFFF00',
															end_color='FFFF00')
		for i in range(6, 7):
			for j in range(2, sheet.max_row):
				if sheet.cell(j, i).value == '合计' and sheet.cell(j, i + 1).value != '合计' and sheet.cell(j,
																										i - 1).value != '合计':
					for c in range(1, sheet.max_column + 1):
						sheet.cell(j, c).font = Font(color='00FF0000')
		print('----已完成样式修改----')
	# 团队品类签收率（停用）
	def OrderQuan(self, team, tem):
		match1 = {'slgat': '港台',
				  'sltg': '泰国',
				  'slxmt': '新马',
				  'slzb': '直播团队',
				  'slyn': '越南',
				  'slrb': '日本'}
		match = {'slgat': '"神龙家族-港澳台"',
				 'sltg': '"神龙家族-泰国"',
				 'slxmt': '"神龙家族-新加坡", "神龙家族-马来西亚"',
				 'slzb': '"神龙家族-直播团队"',
				 'slyn': '"神龙家族-越南"',
				 'slrb': '"神龙家族-日本团队"'}
		# yesterday = (datetime.datetime.now()).strftime('%Y-%m-%d') + ' 23:59:59'
		# yesterday = (datetime.datetime.now().replace(month=5, day=31)).strftime('%Y-%m-%d')
		# yesterday = '2020-08-25'
		# print(yesterday)
		# last_month = (datetime.datetime.now().replace(day=1)).strftime('%Y-%m-%d')
		# last_month = (datetime.datetime.now().replace(month=5, day=27)).strftime('%Y-%m-%d')
		# last_month = '2020-08-15'
		# print(last_month)
		# -*- coding:utf-8 -*-
		sql = ''' SELECT IFNULL(ql.币种,'合计') 币种,IFNULL(ql.年月,'合计') 年月,IFNULL(ql.是否改派,'合计') 是否改派,IFNULL(ql.父级分类,'合计') 父级分类,IFNULL(ql.产品名称,'合计') 产品名称,IFNULL(ql.物流方式,'合计') 物流方式,IFNULL(ql.旬,'合计') 旬,签收,拒收,在途,未发货,未上线,已退货,理赔,自发头程丢件,已发货,已完成,全部,ql.签收 / ql.已完成 AS 完成签收, ql.签收 / ql.全部 AS 总计签收, ql.已完成 / ql.全部 AS 完成占比 ,ql.已发货 / ql.已完成 AS '已完成/已发货' , ql.已退货 / ql.全部 AS 退货率,'' 已发货占比,'' 已完成占比,'' 全部占比 FROM
	    (SELECT qq.币种,qq.年月,qq.是否改派,qq.父级分类,qq.产品名称,qq.物流方式, qq.旬,sum(签收) 签收,sum(拒收) 拒收,sum(在途) 在途,sum(未发货) 未发货,sum(未上线) 未上线,sum(已退货) 已退货,sum(理赔) 理赔,sum(自发头程丢件) 自发头程丢件,sum(已发货) 已发货,sum(已完成) 已完成,sum(全部) 全部 FROM
	    (SELECT q.币种,q.年月,q.是否改派,q.父级分类,q.产品名称,q.物流方式, q.旬,已签收 签收,拒收,在途,未发货,未上线,已退货,理赔,自发头程丢件,已发货,已完成 FROM
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 已签收 FROM sl_tem sl				
	    WHERE sl.最终状态 IN ('已签收')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) q
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 拒收 FROM sl_tem sl				
	    WHERE sl.最终状态 IN ('拒收')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) j
	    ON q.`币种` = j.`币种` AND q.`年月` = j.`年月` AND q.`产品名称` = j.`产品名称` AND q.`物流方式` = j.`物流方式` AND q.`旬` = j.`旬` AND q.`父级分类` = j.`父级分类` AND q.`是否改派` = j.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 在途 FROM sl_tem sl 				
	    WHERE sl.最终状态 IN ('在途')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) zz
	    ON  q.`币种` = zz.`币种` AND q.`年月` = zz.`年月`AND q.`产品名称` = zz.`产品名称`AND q.`物流方式` = zz.`物流方式`AND q.`旬` = zz.`旬` AND q.`父级分类` = zz.`父级分类` AND q.`是否改派` = zz.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 未发货 FROM sl_tem sl 				
	    WHERE sl.最终状态 IN ('未发货')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) wf
	    ON  wf.`币种` = q.`币种` AND wf.`年月` = q.`年月`AND wf.`产品名称` = q.`产品名称`AND wf.`物流方式` = q.`物流方式`AND wf.`旬` = q.`旬` AND q.`父级分类` = wf.`父级分类` AND q.`是否改派` = wf.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 未上线 FROM sl_tem sl 				
	    WHERE sl.最终状态 IN ('未上线') 
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) ws
	    ON  ws.`币种` = q.`币种` AND ws.`年月` = q.`年月`AND ws.`产品名称` = q.`产品名称`AND ws.`物流方式` = q.`物流方式`AND ws.`旬` = q.`旬`AND q.`父级分类` = ws.`父级分类` AND q.`是否改派` = ws.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 已退货 FROM sl_tem sl 				
	    WHERE sl.最终状态 IN ('已退货')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) th
	    ON  q.`币种` = th.`币种` AND q.`年月` = th.`年月`AND q.`产品名称` = th.`产品名称`AND q.`物流方式` = th.`物流方式`AND q.`旬` = th.`旬`AND q.`父级分类` = th.`父级分类`AND q.`是否改派` = th.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 理赔 FROM sl_tem sl				
	    WHERE sl.最终状态 IN ('理赔')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) lp
	    ON  lp.`币种` = q.`币种` AND lp.`年月` = q.`年月`AND lp.`产品名称` = q.`产品名称`AND lp.`物流方式` = q.`物流方式`AND lp.`旬` = q.`旬`AND q.`父级分类` = lp.`父级分类`AND q.`是否改派` = lp.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 自发头程丢件 FROM sl_tem sl				
	    WHERE sl.最终状态 IN ('自发头程丢件') 
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) zf
	    ON  zf.`币种` = q.`币种` AND zf.`年月` = q.`年月`AND zf.`产品名称` = q.`产品名称`AND zf.`物流方式` = q.`物流方式`AND zf.`旬` = q.`旬`AND q.`父级分类` = zf.`父级分类`AND q.`是否改派` = zf.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 已发货 FROM sl_tem sl				
	    WHERE sl.最终状态 IN ('已签收','拒收','理赔','已退货','在途','未上线')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) fh
	    ON  fh.`币种` = q.`币种` AND fh.`年月` = q.`年月`AND fh.`产品名称` = q.`产品名称`AND fh.`物流方式` = q.`物流方式`AND fh.`旬` = q.`旬`AND q.`父级分类` = fh.`父级分类`AND q.`是否改派` = fh.`是否改派`
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 已完成 FROM sl_tem sl				
	    WHERE sl.最终状态 IN ('已签收','拒收','理赔','已退货')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) wc
	    ON  wc.`币种` = q.`币种` AND wc.`年月` = q.`年月`AND wc.`产品名称` = q.`产品名称`AND wc.`物流方式` = q.`物流方式`AND wc.`旬` = q.`旬`AND q.`父级分类` = wc.`父级分类`AND q.`是否改派` = wc.`是否改派`
	    ORDER BY 币种,年月) qq
	    LEFT JOIN
	    (SELECT 币种,年月,是否改派,父级分类,CONCAT(产品id, 产品名称) 产品名称,物流方式, 旬,COUNT(最终状态) 全部 FROM sl_tem sl				
	    WHERE sl.最终状态 IN ('已签收','拒收','理赔','已退货','在途','未上线','未发货')  
	    GROUP BY 币种,年月,是否改派,父级分类,产品名称,物流方式,旬
	    ORDER BY 币种,年月) qb
	    ON  qb.`币种` = qq.`币种` AND qb.`年月` = qq.`年月`AND qb.`产品名称` = qq.`产品名称`AND qb.`物流方式` = qq.`物流方式`AND qb.`旬` = qq.`旬`AND qq.`父级分类` = qb.`父级分类`AND qq.`是否改派` = qb.`是否改派`
	    GROUP BY 年月,是否改派,父级分类,产品名称,物流方式,旬
	    with rollup) ql;'''.format(team, tem)
		print('正在获取-' + match1[team] + '-品类签收率…………')
		df = pd.read_sql_query(sql=sql, con=self.engine1)
		print('----已获' + match1[team] + '-品类签收率')
		columns = list(df.columns)  # 获取数据的标题名，转为列表
		columns_value = ['退货率', '完成签收', '总计签收', '完成占比', '已完成/已发货']
		for column_val in columns_value:
			if column_val in columns:
				df[column_val] = df[column_val].fillna(value=0)
				df[column_val] = df[column_val].apply(lambda x: format(x, '.2%'))
		df.loc['全部'] = df.apply(lambda x: x.sum())
		df.drop(df.index[len(df) - 1], inplace=True)
		print(df)
		print('正在写入EXECL中…………')
		today = datetime.date.today().strftime('%Y.%m.%d')
		df.to_excel('D:\\Users\\Administrator\\Desktop\\输出文件\\{} 神龙{}---2200签收率.xlsx'.format(today, match1[team]),
					sheet_name=match[team], index=False)
		print('----已写入excel')
		#  https://www.cnblogs.com/liming19680104/p/11648048.html 修改表格样式
		filePath = r'D:\\Users\\Administrator\\Desktop\\输出文件\\{} 神龙{}---2200签收率.xlsx'.format(today, match1[team])
		app = xl.App(visible=False, add_book=False)
		wb = app.books.open(filePath, update_links=False, read_only=True)
		print(match[team])
		sht = wb.sheets[match[team]]

		rng = sht.range('a1')
		rng.color = (233, 233, 235)
		sht.range("b2:g4").columns.autofit()
		print(wb.sheets[match[team]].range('d4').value)
		wb.close()
		app.quit()
	# 各团队全部订单表-函数（停用）
	def tgOrderQuanTT(self, team):
		match1 = {'slgat': '港台',
				  'sltg': '泰国',
				  'slxmt': '新马',
				  'slzb': '直播团队',
				  'slyn': '越南',
				  'slrb': '日本'}
		match = {'slgat': '"神龙家族-港澳台"',
				 'sltg': '"神龙家族-泰国"',
				 'slxmt': '"神龙家族-新加坡", "神龙家族-马来西亚"',
				 'slzb': '"神龙家族-直播团队"',
				 'slyn': '"神龙家族-越南"',
				 'slrb': '"神龙家族-日本团队"'}
		yesterday = (datetime.datetime.now() + datetime.timedelta(days=1)).strftime('%Y-%m-%d')
		# yesterday = '2020-08-25'
		print(yesterday)
		last_month = (datetime.datetime.now() - datetime.timedelta(days=2)).strftime('%Y-%m-%d')
		# last_month = '2020-08-15'
		print(last_month)
		sql = '''SELECT a.id,
	                    a.order_number order_number,
	                    dim_area.name area_id,
	                    a.sale_id main_id,
	                    a.ship_phone ship_phone,
	                    a.ship_zip ship_zip,
	                    a.amount amount,
	                    tg_order_status.name order_status,
	                    UPPER(a.waybill_number) waybill_number,
	                    dim_payment.pay_name pay_type,
	                    a.addtime addtime,
	                    a.uptime update_time,
	                    a.product_id goods_id, 
	                    a.qty quantity,
	                    dim_trans_way.all_name logistics_id,
	                    '' op_id,
	                    CONCAT(gk_sale.product_id, gk_sale.product_name) goods_name, 
	                    IF(a.second=0,'直发','改派') secondsend_status,
	                    IF(a.low_price=0,'否','是') low_price
	            FROM gk_order a 
	                left join dim_area ON dim_area.id = a.area_id 
	                left join dim_payment on dim_payment.id = a.payment_id
	                left join gk_sale on gk_sale.product_id = a.product_id 
	                left join dim_trans_way on dim_trans_way.id = a.logistics_id
	                left join tg_order_status on tg_order_status.id = a.order_status
	            WHERE a.rq >= '{}' AND a.rq <= '{}'
	                AND dim_area.name IN ({});'''.format(last_month, yesterday, match[team])
		print('正在获取最近 3 天订单…………')
		try:
			df = pd.read_sql_query(sql=sql, con=self.engine2)
			print('----已获取近 3 天订单')
			# print(df)
			print('正在写入缓存表中…………')
			df.to_sql('tem_tg', con=self.engine3, index=False, if_exists='replace')
		except Exception as e:
			print('更新缓存失败：', str(Exception) + str(e))
		print('++++更新缓存完成++++')
		print('正在写入全部订单表中…………')
		sql = 'REPLACE INTO 全部订单_{} SELECT *, NOW() 添加时间 FROM tem_tg;'.format(team)
		pd.read_sql_query(sql=sql, con=self.engine3, chunksize=100)
		print('----已写入全部订单表中')

if __name__ == '__main__':                    # 以老后台的简单查询为主，
	start = datetime.datetime.now()	
	print('======正在启动查询订单程序>>>>>')
	print('               v           ')
	# s = Bds('qiyuanzhang@jikeyin.com', 'qiyuanzhang123.')
	s = BpsControl('nixiumin@giikin.com', 'nixiumin123@.')
	# # s.getOrderInfo("NR010230026492511", '订单号')
	# s.getOrderInfo("TH009281245118873", '订单号')

	# 获取全部订单表（各团队）
	match = {'slgat': '港台',
		'sltg': '泰国',
		'slxmt': '新马',
		'slzb': '直播团队',
		'slyn': '越南',
		'slrb': '日本'}
	s.tgOrderQuan('sltg')

	# 获取泰国海外仓
	house = 'shifeng'
	match0 = {'shifeng': 'Tracking Number',
			 'chaoshidai': '运单号',
			'bojiatu': '上架单号', }
	s.sltg_HaiWaiCang(house)


