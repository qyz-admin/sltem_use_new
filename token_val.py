import time
from selenium import webdriver
import win32com.client as win32

# option = webdriver.ChromeOptions()
# option.binary_location=r'F:\360\ChromePortable_x64_v770386590\GoogleChromePortable77.0.3865.90(x64)\GoogleChromePortable.exe'
# driver = webdriver.Chrome()
# driver.get('https://www.baidu.com')


# driver = webdriver.Chrome()
driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')
driver.get('https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
driver.implicitly_wait(5)
# 定位搜索按钮
button = driver.find_element_by_xpath("/html/body/section[2]/a[1]")
# 执行单击操作
button.click()

# driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
# driver.implicitly_wait(5)

# js = 'arguments[0].setAttribute("style", arguments[1]);'
# element = driver.find_element('id', 'kw')
# style = 'background: red; border: 2px solid yellow;'
# driver.execute_script(js, element, style)
#
# page_height = driver.execute_script('return document.documentElement.scrollHeight;')
# print(page_height)

time.sleep(3)

# js = 'arguments[0].value = arguments[1];'
js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
            data: { mobile: '+86-18538110674',
                    pwd: 'qyz04163510',
                    goto: 'https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=http://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode',
                    pdmToken: '',
                    araAppkey: '1917',
                    araToken: '0#19171646622570440595157649661652144581488416G6D6E584D74E37BE891FAC3A49235AAA00C9B53',
                    araScene: 'login',
                    captchaImgCode: '',
                    captchaSessionId: '',
                    type: 'h5'
                },
                type: 'POST',
                timeout: '10000',
                async:false, 
                beforeSend(xhr, settings) {
                    xhr.setRequestHeader = XMLHttpRequest.prototype.setRequestHeader;
                },
                success: function(data) {
                    if (data.success) {
                         console.log(data.data);
                         console.log("loginTmpCode值是：", data.data.split('loginTmpCode=')[1]);
                          document.documentElement.getElementsByClassName("noGoto")[0].textContent = data.data.split('loginTmpCode=')[1];
                         arguments[0].value=data.data.split('loginTmpCode=')[1];
                         alert(arguments[0].value)
                    } else {
                            console.log(data.code);
                    }
                },
                error: function(error) {
                    alert("请检查网络");
                }
            });
            '''
element = driver.find_element('id', 'mobile')
style = '99999;'
driver.execute_script(js, element)

driver.implicitly_wait(5)

page_height = driver.execute_script('return document.documentElement.getElementsByClassName("noGoto")[0].textContent;')
print(page_height)
# driver.quit()

print(page_height)


# def fun(self, val):
#     # 实际上测试excel时也可以直接这样做  提用wps使用
#     application = win32.Dispatch('Excel.Application')
#
#     # Path指的是本地表格文件的路径，比如:
#     Path = r"H:\桌面\需要用到的文件\slgat_签收计算(ver5.24)(20).xlsm"
#     # 通过Win32的方式并不限制xls和xlsx（因为操作是wps在做）
#     workbook = application.Workbooks.Open(Path)
#
#     # Sheet指的就是表名（注意不是文件名）
#     worksheet = workbook.Worksheets('Sheet1')
#     worksheet.Cells(1, 1).Value = 'Hello World'
#     worksheet.Name = 'MySheet'
#     # Path是指要将文件保存到哪一个位置
#     worksheet.SaveAs(Path)
#
#     # 注意，此处退出的是文件，而不是WPS程序
#     # 相当于在WPS中关闭了文件
#     # 如果操作完表格不关闭文件，极有可能会造成文件无法打开
#     workbook.Close()

