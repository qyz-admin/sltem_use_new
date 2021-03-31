import requests
import json
r = requests.get("http://gimp.giikin.com/service?service=gorder.customer&action=getProductList&page=1&pageSize=10&productId=214575&productName=&status=&source=&isSensitive=&isGift=&isDistribution=&chooserId=&buyerId=")
print(r.status_code)
print(r.json())
rq = r.json()

print(55)
# r = requests.get(rq['location'])
print(r)

# url：接口地址
url = "https://www.baidu.com/"
# 请求的数据：以字典形式{key:value}
# data = {"name":"zhangsan","pwd":"a123456"}

# 发送get请求
# res = requests.get(url,data) # 发送带有请求参数的GET请求
res = requests.get(url)
# 输出响应数据
print(res) # 输出响应数据中最后的HTTP状态码
print(res.text) # 输出字符串格式
print(res.json()) # 输出json格式



