class Settings():
    def __init__(self):
        self.excelPath = r'D:\Users\Administrator\Desktop\直发表'
        self.mysql1 = {'host': 'localhost',      #数据库地址
                      'user': 'root',           #数据库账户
                      'port': '3306',
                      'password': '123456',     #数据库密码   654321
                      'datebase': 'logistics_status',   #数据库名称
                      'charset': 'utf8'         #数据库编码
                       }
        self.mysql200 = {'host': 'tidb.giikin.com',  # 数据库地址
                       'user': 'shenlongkf',  # 数据库账户
                       'port': '4000',
                       'password': 'SIK87&67asd',  # 数据库密码
                       'datebase': 'gdqs_shenlong',  # 数据库名称
                       'charset': 'utf8'  # 数据库编码
                       }
        self.mysql20 = {'host': 'tidb.giikin.cn',  # 数据库地址
                       'user': 'jp',  # 数据库账户
                       'port': '39999',
                       'password': 'vVYkaElSON',  # 数据库密码
                       'datebase': 'gdqs_jp',  # 数据库名称
                       'charset': 'utf8'  # 数据库编码
                       }
        self.mysql2 = {'host': 'tidb.giikin.cn',  # 数据库地址--本地神龙客服
                       'user': 'tw',  # 数据库账户
                       'port': '39999',
                       'password': '1tSRyqQEfF',  # 数据库密码
                       'datebase': 'gdqs_hw',  # 数据库名称
                       'charset': 'utf8'  # 数据库编码
                       }
        self.mysql3 = {'host': 'localhost',      #数据库地址
                      'user': 'root',           #数据库账户
                      'port': '3306',
                      'password': '123456',     #数据库密码
                      'datebase': '订单数据',   #数据库名称
                      'charset': 'utf8'         #数据库编码
                       }
        self.mysql4 = {'host': 'tidb.giikin.cn',      #数据库地址
                      'user': 'jinpeng',           #数据库账户
                      'port': '39999',
                      'password': 'BpWzdlaG8',     #数据库密码
                      'datebase': 'gdqs_jinpeng',   #数据库名称
                      'charset': 'utf8'         #数据库编码
                       }
        self.logistics = {
            '时丰': {'apiToken': '6XgVJqpdy3lttVh7DniJtXdLdxzDhHk12RgMurLqM9aCFeimGrNB08cs0233',
                   'apiUrl': r'https://timesoms.com/api/orders/track/'},
            '博佳图': {'apiAccount': 'cc4ea61156676d5e51f85fd1c0588c56',
                    'apiPassword': 'cc4ea61156676d5e51f85fd1c0588c5630dfa67fb8269572f64ab3e1e2dae0bb',
                    'apiUrl': r'http://120.79.190.37/default/svc/wsdl'},
            '超时代': {'apiUrl': r'http://134.175.15.128:8082/selectTrack.htm?documentCode='}}
        self.system = {
            '后台': {'username': 'louweibin@giikin.com',
                   'password': 'wo.1683485'}}
        self.email = {'smtp': 'smtp.163.com',
                      'email': 'qyz1404039293@163.com',      # 密码： qyz04163510
                      'password': 'UIYCFUQSGJZMYNDY'}
        #  https://blog.csdn.net/weixin_36931308/article/details/103767758 （危机处理需要）
        #  LECOGDYYBJUJJBST

        #  台湾token, 日本token:  token_sl
        #  新马token, 泰国token:  token_hfh
        self.token = {'token_sl': 'fa1e6e884f18c4151be7b33cdec00f79',
                      'token_hfh': 'c96f9540a158db716dbdfb8e0a695b35'}

        self.logistics_name = '''
                            "台湾-森鸿-新竹-自发头程", "台湾-森鸿-新竹","台湾-大黄蜂普货头程-森鸿尾程","台湾-立邦普货头程-森鸿尾程", "台湾-易速配头程-铱熙无敌尾",
                            "台湾-立邦普货头程-易速配尾程","台湾-大黄蜂普货头程-易速配尾程",  
                            "台湾-易速配-新竹", "台湾-易速配-TW海快", "台湾-易速配-海快头程【易速配尾程】",
                            "台湾-铱熙无敌-711超商","台湾-铱熙无敌-新竹", "台湾-铱熙无敌-黑猫", 
                            "台湾-速派-711超商", "台湾-速派-新竹", "台湾-速派-黑猫", "台湾-速派宅配通",
                            "台湾-天马-新竹","台湾-天马-顺丰","台湾-天马-黑猫",
                            "香港-圆通", "香港-立邦-顺丰","香港-森鸿物流", "香港-森鸿-SH渠道","香港-森鸿-顺丰渠道","香港-易速配-顺丰",
                            "森鸿","龟山",
                            "速派新竹", "速派黑猫", "速派宅配通",
                            "台湾-铱熙无敌-新竹改派", "台湾-铱熙无敌-黑猫改派",
                            "天马顺丰","天马黑猫","天马新竹",
                            "香港-圆通-改派","香港-立邦-改派","香港-森鸿-改派","香港-易速配-改派"
                        '''
        self.team_name = '''IF(团队 in ('火凤凰-台湾','火凤凰-香港'),'火凤凰港台',
                                            IF(团队 in ('神龙家族-台湾','神龙-香港'),'神龙港台',
                                            IF(团队 = '客服中心-港台','客服中心港台',
                                            IF(团队 = '研发部-研发团队','研发部港台',
                                            IF(团队 = '神龙-主页运营','神龙主页运营',
                                            IF(团队 = '红杉家族-港澳台','红杉港台',
                                            IF(团队 = '郑州-北美','郑州北美',
                                            IF(团队 = '金狮-港澳台','金狮港台',
                                            IF(团队 = '金鹏家族-4组','金鹏港台',
                                            IF(团队 = '翼虎家族-mercadolibre','翼虎港台',团队)))))))))) AS 家族
                        '''
        self.team_name2 = '''
                        "神龙港台","火凤凰港台","雪豹港台","金蝉项目组","金蝉家族优化组","金蝉家族公共团队","客服中心港台","奥创队","神龙主页运营","APP运营","Line运营","红杉港台","郑州北美""研发部港台","金鹏港台","金狮港台","翼虎港台"
                    '''