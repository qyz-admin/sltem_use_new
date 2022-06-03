# coding: utf-8
import requests, json


# 发送钉钉消息
def send_dingtalk_message(url, content, mobile_list, isAtAll):
    headers = {'Content-Type': 'application/json',"Charset": "UTF-8"}
    if isAtAll =='是':
        data = {"msgtype": "text",
                "text": {  # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                        "content": content
                    },
                "at": {
                    # "atMobiles": mobile_list, @所有人写True并且将上面atMobiles注释掉
                    # 是否@所有人
                    "isAtAll": True
                }
        }
    else:
        data = {
            "msgtype": "text",
            "text": {
                # 要发送的内容【支持markdown】【！注意：content内容要包含机器人自定义关键字，不然消息不会发送出去，这个案例中是test字段】
                "content": content
            },
            "at": {
                # 要@的人
                "atMobiles": mobile_list,
                # 是否@所有人
                "isAtAll": False  #@全体成员（在此可设置@特定某人）
            }
    }

    # 4、对请求的数据进行json封装
    sendData = json.dumps(data)  # 将字典类型数据转化为json格式
    sendData = sendData.encode("utf-8")  # python3的Request要求data为byte类型

    r = requests.post(url, headers=headers, data=json.dumps(data))
    print(r.text)
    return r.text


if __name__ == "__main__":
    # 获取dingtalk token url
    # access_token = "https://oapi.dingtalk.com/robot/send?access_token=8ca07327d17b01673cd9cd76bbfad0a90764d110871136aa1ee755da21d6057f"  # url为机器人的webhook
    # access_token = "https://oapi.dingtalk.com/robot/send?access_token=a86d784649c4e0f02a53a856afabd343b6c0ffb4a459bf536b3aa8499b074ba2"  # url为机器人的webhook

    access_token = "https://oapi.dingtalk.com/robot/send?access_token=bdad3de3c4f5e8cc690a122779a642401de99063967017d82f49663382546f30"  # url为机器人的webhook
    content = '1,测试消息'      # 钉钉消息内容，注意test是自定义的关键字，需要在钉钉机器人设置中添加，这样才能接收到消息
    mobile_list = ['18538110674']    # 要@的人的手机号，可以是多个，注意：钉钉机器人设置中需要添加这些人，否则不会接收到消息
    isAtAll = '是'            # 是否@所有人
    # 发送钉钉消息
    send_dingtalk_message(access_token, content, mobile_list, isAtAll)