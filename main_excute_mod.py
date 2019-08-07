import tkinter as tk
import json
from tkinter import Label,ttk,filedialog,messagebox
from tkinter.constants import *
import selenium
from selenium import webdriver
import requests
import re
import datetime
import xlrd
import xlwt
from class_flight import *
from date_result import *
from get_flight_data_chrome import *
excel_plan = xlrd.open_workbook("航班数据.xlsx")
sheet_plan = excel_plan.sheet_by_name('计划航班数据')


array_plan=[]  #计划航班的数组，每个成员都为一个Doule_Unit对象
array_flt_dom =[]  #国内FlightClass对象的数组
array_flt_int =[]  #国际FlightClass对象的数组
array_Doule_Unit_make(sheet_plan,array_plan)
DATE = ''
array_departure = []

i = datetime.datetime.now()    #获取日期的准备数据和具体函数
array_month = [1,2,3,4,5,6,7,8,9,10,11,12]
array_day_sum = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]
def get_array_day(nian,yue):
	if yue in [1,3,5,7,8,10,12]:
		return array_day_sum
	elif yue in [4,6,9,11]:
		return array_day_sum[0:-1]
	elif nian%4 == 0:
		return array_day_sum[0:29]
	else:
		return array_day_sum[0:28]
nian = i.year
yue = i.month
ri = i.day


window = tk.Tk()   #主窗口
window.title('日数据上传')  #窗口标题
window.geometry('500x500')   #窗口尺寸

l_yuangonghao = Label(window,text='员工号')  #员工号输入窗口
l_yuangonghao.pack()
e_yuangonghao = tk.Entry(window,show=None)
e_yuangonghao.pack()

l_mima = Label(window,text='密码')  #密码输入窗口
l_mima.pack()
e_mima = tk.Entry(window,show='*')
e_mima.pack()

l_yanzhengma = Label(window,text='验证码')  #验证码输入窗口
l_yanzhengma.pack()
e_yanzhengma = tk.Entry(window,show=None)
e_yanzhengma.pack()



label_img = tk.Label(window)  #验证码图片窗口
label_img.pack() 

chrome_options = webdriver.ChromeOptions()  #登陆soc
chrome_options.add_argument('--headless')
browser = webdriver.Chrome(chrome_options=chrome_options)
browser.get('https://soc.csair.com/opws-web/security-Remember-gotoLogin.action')#打开soc网站
jsid = browser.get_cookies()[0]['value']  #返回的cookie中的关键数值
header = {
'User-Agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.78 Safari/537.36',
'Cookie':'languageValue=zh_cn; JSESSIONID='+jsid+'; JSESSIONID='+jsid
}

houzhui = browser.find_element_by_css_selector('[title="点击换一张"]')  #获取验证码图片，根据图片内容填验证码
url_code = houzhui.get_attribute('src')
r = requests.get(url_code,headers=header)
with open('code.jpg','wb') as f:
	f.write(r.content)
	f.close()
img_gif = tk.PhotoImage(file = 'code.jpg')
label_img.config(image=img_gif)




def login_test():     #登陆函数
	global label_img,img_gif
	
	input_u = browser.find_element_by_id('j_username') 
	input_u.send_keys(e_yuangonghao.get())
	input_p = browser.find_element_by_id('j_password')  
	input_p.send_keys(e_mima.get())
	input_c = browser.find_element_by_id('captcha')
	input_c.send_keys(e_yanzhengma.get())
	button_login = browser.find_element_by_id('loginbtn')
	button_login.click()
	try:
		if browser.find_element_by_css_selector('[id="checkout"]'):
			l_login_success.config(text = '登陆成功')
			return True
	except selenium.common.exceptions.NoSuchElementException:
		l_login_success.config(text = '验证码或密码错误，请重新输入')
		houzhui = browser.find_element_by_css_selector('[title="点击换一张"]')  #获取验证码图片，根据图片内容填验证码
		url_code = houzhui.get_attribute('src')
		r = requests.get(url_code,headers=header)
		with open('code.jpg','wb') as f:
			f.write(r.content)
			f.close()
		img_gif = tk.PhotoImage(file = 'code.jpg')
		label_img.config(image=img_gif)

 

b = tk.Button(window,    #登陆按钮
    text='登陆',      
    width=5, height=1, 
    command=login_test)    
b.pack()

l_login_success = Label(window,text='')  #登陆信息显示
l_login_success.pack()

def test_canlinder(*args):    #根据日期数据整理成类似'2019-01-25'的字符串格式
	nian_ = combox_year.get()
	yue_ = combox_month.get()
	ri_ = combox_ri.get()
	combox_ri["values"] = get_array_day(int(nian_),int(yue_))
	if len(yue_)==1:
		yue_ = '0'+yue_
	if len(ri_)==1:
		ri_ = '0'+ri_ 
	return(nian_+'-'+yue_+'-'+ri_)
	
frame_root = tk.Frame(window)   #日期下拉框
  

combox_year = ttk.Combobox(frame_root,width=12, height=1)
combox_year["values"] = [nian-2,nian-1,nian]
combox_year.current(2)
combox_year.bind("<<ComboboxSelected>>",test_canlinder)
combox_year.pack(side = LEFT)

combox_month = ttk.Combobox(frame_root,width=12, height=1)
combox_month["values"] = array_month
combox_month.current(yue-1)
combox_month.bind("<<ComboboxSelected>>",test_canlinder)
combox_month.pack(side = LEFT)

if ri==1:
	combox_month.current(yue-2)
else:
	combox_month.current(yue-1)
combox_month.bind("<<ComboboxSelected>>",test_canlinder)
combox_month.pack(side = LEFT)

combox_ri = ttk.Combobox(frame_root,width=12, height=1)
if ri!=1:
	combox_ri["values"] = get_array_day(nian,yue)
	combox_ri.current(ri-2)
else:
	ri = get_array_day(nian,yue-1)[-1]
	combox_ri["values"] = get_array_day(nian,yue-1)
	combox_ri.current(ri-1)

combox_ri.bind("<<ComboboxSelected>>",test_canlinder)
combox_ri.pack(side = LEFT)

frame_root.pack()

l_blank1 = Label(window,text='')
l_blank1.pack()


array_branch = ['吉林分公司','河南分公司','深圳分公司','北京分公司','新疆分公司','广州营业部','北方分公司','黑龙江分公司','大连分公司','海南分公司','湖南分公司','湖北分公司']

frame_shifa = tk.Frame(window)
e_shifa = tk.Entry(frame_shifa,show=None)
e_shifa.insert(END,string = 'CGQ/NBS/YNJ')     #始发地输入框
e_shifa.pack(side = LEFT)
l_shifa = Label(frame_shifa,text='输入始发地，以’/’隔开，如图所示')
l_shifa.pack(side = LEFT)
frame_shifa.pack()

l_blank2 = Label(window,text='')
l_blank2.pack()

frame_branch = tk.Frame(window)
combox_branch = ttk.Combobox(frame_branch,width=12, height=1)
combox_branch["values"] = array_branch
combox_branch.current(0)
combox_branch.pack(side = LEFT)
l_brh = Label(frame_branch,text='选择公司名称')
l_brh.pack(side = LEFT)
frame_branch.pack()

l_blank3 = Label(window,text='')
l_blank3.pack()


file_data = ''
def flightfileopen():
	global file_data
	file_data = filedialog.askopenfilename()
	l_data_file.config(text = "打开的数据文件："+file_data)
bt_open_data = tk.Button(window,text='打开数据文件',command=flightfileopen)
bt_open_data.pack()	
	

l_data_file =tk.Label(window,text='')
l_data_file.pack()	




def make_data():   #主函数，开始生产数据了
	l_data_file =tk.Label(window,text='')
	l_data_file.pack()
	print(file_data)
	excel_data = xlrd.open_workbook(file_data)
	sheet_data = excel_data.sheet_by_name('Sheet1')
	if l_login_success.cget("text")!='登陆成功':
		tk.messagebox.showinfo(title='', message='请先登录soc')
		return False
	DATE = test_canlinder()
	str_shifa = e_shifa.get()
	num_shifa=1
	for c in str_shifa:
		if c=='/':
			num_shifa+=1
	i=0
	while i<num_shifa:
		j = i*4
		array_departure.append(str_shifa[j:j+3])
		i+=1
	print(DATE,array_departure,file_data)
	for DEPARTURE in array_departure:
		items = get_html_items(jsid,DATE,DEPARTURE)
		for i in items:
			print(i)
		create_array_FlightClass(items,array_flt_dom,array_flt_int,array_plan)
	data_make(sheet_data,array_flt_int,array_flt_dom)
	data_complete(array_flt_int,array_flt_dom)
	data_write(array_flt_int,array_flt_dom,DATE,combox_branch.get())
	tk.messagebox.showinfo(title='', message='已完成')
	window.destroy()  #窗口退出函数

	
	
	
b_make_data = tk.Button(window, 
    text='开始生成数据',      # 显示在按钮上的文字
    width=12, height=1, 
    command=make_data)     # 点击按钮式执行的命令
b_make_data.pack()  









window.mainloop() 
