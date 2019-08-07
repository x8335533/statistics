"""生成报表格中每行数据类对象的数组"""
import xlrd
import xlwt
from class_flight import *
import re
excel_fentan = xlrd.open_workbook('分摊.xls')
sheet_fentan = excel_fentan.sheet_by_name('国内')


class Doule_Unit:    #计划航班的类，其对象成员分别为flt_num航班号，flt_dep出发地，des1第一目的地，des2第二目的地，如第二目的地为空，则des2的值为'nothing'
	def __init__(self,fn,fdep,fdes1,fdes2):
		self.flt_num = fn
		self.flt_dep = fdep
		self.des1=fdes1
		self.des2=fdes2


def judge_flight_type(s):
	if re.search("货机",s):
		return "货机"
	else:
		return "客机"	
	

			


def array_Doule_Unit_make(sht,array_plan):  #根据航班计划表填充计划航班的数组
	i=0
	while i<sht.nrows:
		if sht.cell(i,3).ctype!=0:
			array_plan.append(Doule_Unit(sht.cell_value(i,0),\
			sht.cell_value(i,1),sht.cell_value(i,2),sht.cell_value(i,3)))
		else:
			array_plan.append(Doule_Unit(sht.cell_value(i,0),\
			sht.cell_value(i,1),sht.cell_value(i,2),'nothing'))	
		i+=1


def create_array_FlightClass(items,array_dom,array_int,array_plan):  #取得上报表格中每行数据类对象的数组
	lenth = len(array_plan)                                          #items数组成员分别为航班性质，航班号，始发地，目的地，机型
	for item in items:
		ftp = judge_flight_type(item[0])
		
		dep = item[1]
		dep = re.sub('CZ0','CZ',dep)   #把类似CZ0623的航班号变为CZ623
		if item[3] in int_place:       #如果条目中的航班号和始发站能与航班计划中的数据一致,则把航班计划的目的站填入最终数组,目的是为了纠正SOC数据的错误
			i = 0
			for a in array_plan:   
				if dep == a.flt_num and item[2] == a.flt_dep:
					array_int.append(FlightClass(dep,item[4],item[2],a.des1,cq = 'INT',fq = ftp))
					if a.des2!='nothing':
						array_int.append(FlightClass(dep,item[4],item[2],a.des2,cq = 'INT',fq = ftp))
					break		
				i+=1
				
			if i==lenth:               #如果出现SOC的航班号在航班计划中找不到的情况,则在最终数组添加一个新的对象,航班号是SOC的航班号
				array_int.append(FlightClass(dep,item[4],item[2],item[3],cq = 'INT',fq = ftp))	
		else:
			i = 0
			for a in array_plan:
				if dep == a.flt_num and item[2] == a.flt_dep:
					array_dom.append(FlightClass(dep,item[4],item[2],a.des1,cq = 'DOM',fq = ftp))
					if a.des2!='nothing':
						array_dom.append(FlightClass(dep,item[4],item[2],a.des2,cq = 'DOM',fq = ftp))
					break
				i+=1	
			if i==lenth:
				array_dom.append(FlightClass(dep,item[4],item[2],item[3],cq = 'DOM',fq = ftp))			


def filldata(obj,sheet,line):
	if int(float(sheet.cell_value(line,32))) == 0:  #公务货物
		w = sheet.cell_value(line,13)
		w = int(float(w))
		obj.set_weight_gongwu(w)
	elif re.match('PST',sheet.cell_value(line,0)):   #邮件
		w = sheet.cell_value(line,13)
		w = round(float(w),1)
		obj.set_weight_youjian(w)
		m =  sheet.cell_value(line,32)
		m = round(float(m),2)
		obj.set_money_youjian(m)
	elif sheet.cell_value(line,8) != sheet.cell_value(line,9): #中转货物
		w = sheet.cell_value(line,13)
		w = int(float(w))
		obj.set_weight_zhongzhuan(w)
		m =  sheet.cell_value(line,32)
		m = int(float(m))
		obj.set_money_zhongzhuan(m)
		m = earn_income(sheet.cell_value(line,7),sheet.cell_value(line,8),sheet.cell_value(line,9),m,sheet_fentan)
		obj.set_earn_zhongzhuan(m)
	elif sheet.cell_value(line,6) != sheet.cell_value(line,7): #联程货物
		w = sheet.cell_value(line,13)
		w = int(float(w))
		obj.set_weight_liancheng(w)
		m =  sheet.cell_value(line,32)
		m = int(float(m))
		obj.set_money_liancheng(m)
	elif sheet.cell_value(line,8) == sheet.cell_value(line,9):  #直达货物
		w = sheet.cell_value(line,13)
		w = int(float(w))
		obj.set_weight_zhida(w)
		m =  sheet.cell_value(line,32)
		m = int(float(m))
		obj.set_money_zhida(m)

def data_make(sht,arrayint,arraydom):  #填报上报表格数据数组对象的数据
	i=5
	j=0
	num_arrayint = len(arrayint)  #国际上报数据数组的长度
	num_arraydom = len(arraydom)  #国内上报数据数组的长度
	czpt = re.compile('CZ')
	mailpt = re.compile('PST')
	while i<sht.nrows:
		if sht.cell_value(i,8) not in int_place and sht.cell_value(i,9) in int_place:  #如果是国内中转至国际
			j=0  #计数器，表示找打了第几个数据行
			for a in arrayint:
				if a.flight_num == sht.cell_value(i,4).strip() and a.depature == sht.cell_value(i,7) and a.destination == sht.cell_value(i,8): #在国际上报数组中找到了数据行
					filldata(a,sht,i)
					break
				j+=1
			if j == num_arrayint:  #没有找到数据行，计数器的值为国际上报数据数组的长度
				a_add = FlightClass(sht.cell_value(i,4).strip(),'unknown',sht.cell_value(i,7),sht.cell_value(i,8),cq = 'INT')   #根据唐翼数据生成一个新的对象,但是机型未知
				filldata(a_add,sht,i)
				num_arrayint+=1
				arrayint.append(a_add)
		elif sht.cell_value(i,8) in int_place:  #如果是国际航班
			for a in arrayint:
				if a.flight_num == sht.cell_value(i,4).strip() and a.depature == sht.cell_value(i,7) and a.destination == sht.cell_value(i,8):
					filldata(a,sht,i)
					break
		elif sht.cell_value(i,8) not in int_place:  #如果是国内航班
			for a in arraydom:
				if a.flight_num == sht.cell_value(i,4).strip() and a.depature == sht.cell_value(i,7) and a.destination == sht.cell_value(i,8):
					filldata(a,sht,i)
					break
		i+=1

def data_complete(arrayint,arraydom):   #对于国内转国际的航班机型未知的情况,查找国内数组对应的对象,填入机型数据
	for a in arrayint:
		if a.flight_type == 'unknown':
			for b in arraydom:
				if a.flight_num == b.flight_num:
					a.flight_type = b.flight_type
		a.set_rate()
	for a in arraydom:
		a.set_rate()				
					
def data_write(arrayint,arraydom,date,branch):
	date = re.sub('-','/',date)
	date = re.sub('/0','/',date)
	excel_write = xlwt.Workbook()
	sheet_write = excel_write.add_sheet("Sheet1",True)
	arr_write = ['序号','航班号','机型','航班出发站','航班到达站','直达运量','中转运量','联程运量','邮件运量	','公务运量','直达收入','中转收入','联程收入','邮件收入','货物性质','航班日期','航班性质','所属单位','平均费率','中转收入']				
	i=0
	for c in arr_write:
		sheet_write.write(0,i,c)
		i+=1
	i=1		
	for a in arraydom:
		sheet_write.write(i,0,i)
		sheet_write.write(i,1,a.flight_num)
		sheet_write.write(i,2,a.flight_type)
		sheet_write.write(i,3,a.depature)
		sheet_write.write(i,4,a.destination)
		sheet_write.write(i,5,a.weight_zhida)
		sheet_write.write(i,6,a.weight_zhongzhuan)
		sheet_write.write(i,7,a.weight_liancheng)
		sheet_write.write(i,8,a.weight_youjian)
		sheet_write.write(i,9,a.weight_gongwu)
		sheet_write.write(i,10,a.money_zhida)
		sheet_write.write(i,11,a.money_zhongzhuan)
		sheet_write.write(i,12,a.money_liancheng)
		sheet_write.write(i,13,a.money_youjian)
		sheet_write.write(i,14,a.cargo_quality)
		sheet_write.write(i,15,date)
		sheet_write.write(i,16,a.flight_quality)
		sheet_write.write(i,17,branch)
		sheet_write.write(i,18,a.rate)
		sheet_write.write(i,19,a.earn_zhongzhuan)
		i+=1
	for a in arrayint:
		sheet_write.write(i,0,i)
		sheet_write.write(i,1,a.flight_num)
		sheet_write.write(i,2,a.flight_type)
		sheet_write.write(i,3,a.depature)
		sheet_write.write(i,4,a.destination)
		sheet_write.write(i,5,a.weight_zhida)
		sheet_write.write(i,6,a.weight_zhongzhuan)
		sheet_write.write(i,7,a.weight_liancheng)
		sheet_write.write(i,8,a.weight_youjian)
		sheet_write.write(i,9,a.weight_gongwu)
		sheet_write.write(i,10,a.money_zhida)
		sheet_write.write(i,11,a.money_zhongzhuan)
		sheet_write.write(i,12,a.money_liancheng)
		sheet_write.write(i,13,a.money_youjian)
		sheet_write.write(i,14,a.cargo_quality)
		sheet_write.write(i,15,date)
		sheet_write.write(i,16,a.flight_quality)
		sheet_write.write(i,17,branch)
		sheet_write.write(i,18,a.rate)
		sheet_write.write(i,19,a.earn_zhongzhuan)
		i+=1
	excel_write.save('数据导入1.xls')
		
if __name__ =="__main__":
	pass 					

