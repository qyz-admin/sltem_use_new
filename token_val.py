import time
from selenium import webdriver

# option = webdriver.ChromeOptions()
# option.binary_location=r'F:\360\ChromePortable_x64_v770386590\GoogleChromePortable77.0.3865.90(x64)\GoogleChromePortable.exe'
# driver = webdriver.Chrome()
# driver.get('https://www.baidu.com')


# driver = webdriver.Chrome()
driver = webdriver.Chrome(r'C:\Program Files\Google\Chrome\Application\chromedriver.exe')
driver.get('https://login.dingtalk.com/login/index.htm?goto=https://oapi.dingtalk.com/connect/oauth2/sns_authorize?appid=dingoajqpi5bp2kfhekcqm&response_type=code&scope=snsapi_login&state=STATE&redirect_uri=https://gsso.giikin.com/admin/dingtalk_service/getunionidbytempcode')
driver.implicitly_wait(5)

# js = 'arguments[0].setAttribute("style", arguments[1]);'
# element = driver.find_element('id', 'kw')
# style = 'background: red; border: 2px solid yellow;'
# driver.execute_script(js, element, style)
#
# page_height = driver.execute_script('return document.documentElement.scrollHeight;')
# print(page_height)

# time.sleep(3)

# js = 'arguments[0].value = arguments[1];'
js = '''$.ajax({url: "https://login.dingtalk.com/login/login_with_pwd",
            data: { mobile: '+86-18538110674',
                    pwd: 'qyz35100416',
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
driver.quit()

print(page_height)

