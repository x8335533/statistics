"""获得某一日的航班数据，始发地和日期都是由用户输入的"""
from selenium import webdriver
import requests
import re
def get_cookie_value():   #登录一次soc，返回cookie中的关键数值。
	chrome_options = webdriver.ChromeOptions()
	chrome_options.add_argument('--headless')
	browser = webdriver.Chrome(chrome_options=chrome_options)
	browser.get('https://soc.csair.com/opws-web/security-Remember-gotoLogin.action')#打开soc网站
	input_u = browser.find_element_by_id('j_username') #登录的用户名来自文件“user.txt”
	input_u.send_keys(get_username())
	input_p = browser.find_element_by_id('j_password')  #登录的密码来自文件“user.txt”
	input_p.send_keys(get_code())	
	jsid = browser.get_cookies()[0]['value']  #返回的cookie中的关键数值
	header = {
'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.78 Safari/537.36',
'Cookie':'languageValue=zh_cn; JSESSIONID='+jsid+'; JSESSIONID='+jsid
}
	input_c = browser.find_element_by_id('captcha')
	houzhui = browser.find_element_by_css_selector('[title="点击换一张"]')  #获取验证码图片，根据图片内容填验证码
	url_code = houzhui.get_attribute('src')
	r = requests.get(url_code,headers=header)
	with open('code.jpg','wb') as f:
		f.write(r.content)
	code_captcha = input('输入验证码：')
	input_c.send_keys(code_captcha)
	button_login = browser.find_element_by_id('loginbtn')             
	button_login.click()
	array_cookie = browser.get_cookies()
	browser.close()
	return jsid #获得cookie
	
def get_html_items(jsid,DATE,DEPARTURE): #获得日期为DATE，出发地为DEPARTURE的航班数据数组，该数组的每个成员都为一个数组，成员从索引0到3分别为航班号，始发地，目的地，机型
	print(jsid)
	header = {
'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.78 Safari/537.36',
'Cookie':'languageValue=zh_cn; JSESSIONID='+jsid+'; JSESSIONID='+jsid
}#获得某一日、某地始发历史航班数据的请求头
	
	url_history = 'https://soc.csair.com/opws-web/flightinfo-FlightHistory\
-findByFlightHistoryDataJSON.action?dataJson=%7B%0A++%22schDepDtFrom%22\
+%3A+%22'+DATE+'+%22%2C%0A++%22schDepDtTo%22+%3A+%22'+DATE+'%22%2C%0A+\
+%22fromHour%22+%3A+null%2C%0A++%22toHour%22+%3A+null%2C%0A++%22branch\
Code%22+%3A+%22ALL%22%2C%0A++%22scfArp%22+%3A+null%2C%0A++%22tailBranc\
h%22+%3A+null%2C%0A++%22depCd%22+%3A+%22'+DEPARTURE+'%22%2C%0A++%22arvCd\
%22+%3A+%22%22%2C%0A++%22fltNr%22+%3A+%22%22%2C%0A++%22latestTailNr%22+\
%3A+%22%22%2C%0A++%22latestEqpCd%22+%3A+%22190%40319%4031C%4031G%40320%\
40321%4032C%4032D%4032E%4032G%4032L%4032M%4032N%40330%40332%40333%4033B\
%4033G%4033W%40380%40737%40738%4073C%4073D%4073K%4073L%4073M%4073N%4073\
Q%40747%4074F%40757%40777%4077F%4077W%40787%40789%4078W%4073S%407M8%403\
3C%4033H%40773%40788%4032Q%4032H%407MA%4032Q%4078Z%4078C%407MD%4032Y%40\
32Z%40359%40350%22%2C%0A++%22reason%22+%3A+null%2C%0A++%22svcType%22+%3A\
+null%2C%0A++%22arpCd%22+%3A+null%2C%0A++%22dataType%22+%3A+null%2C%0A++\
%22isCancel%22+%3A+null%0A%7D'
	forms={
'schDepDtFrom':DATE,
'schDepDtTo':DATE,
"fromHour" : 'null',
"toHour" : 'null',
  "branchCode" : "ALL",
  "scfArp" : 'null',
  "tailBranch" : 'null',
  "depCd" : DEPARTURE,
  "arvCd" : "",
  "fltNr" : "",
  "latestTailNr" : "",
  "latestEqpCd" : "190@319@31C@31G@320@321@32C@32D@32E@32G@32L@32M@32N@330@332@333@33B@33G@33W@380@737@738@73C@73D@73K@73L@73M@73N@73Q@747@74F@757@777@77F@77W@787@789@78W@73S@7M8@33C@33H@773@788@32Q@32H@7MA@32Q@78Z@78C@7MD@32Y@32Z@359@350",
  "reason" : 'null',
  "svcType" : 'null',
  "arpCd" : 'null',
  "dataType" : 'null',
  "isCancel" : 'null',
}

	reponse_history = requests.post(url_history,headers=header,data=forms)  #获取了网页的html代码
	pattern = re.compile('\"chnDesc\":\"(.*?)\".*?\"fltNr\":\"(.*?)\".*?\"latestDepArpCd\":\"(\w{3})-(\w{3})\".*?\"latestEqpCd\":\"(.*?)\"')#\"chnDesc\":\"(.*?)\".*?
	items = re.findall(pattern,reponse_history.text)  #找到了每个航班的数据
	return items #items的每个成员都为一个数组，成员分别为航班性质，航班号，始发地，目的地，机型
if __name__ =="__main__":
	
	print(get_html_items())
