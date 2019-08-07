""" 最终上报表格的数每行据都是一个类FlightClass的对象"""
import tkinter as tk
int_place =('ICN','NRT','NGO','TPE','CJU','KIX','FRA','CDG','CJJ','BKK','PUS','CEB')
import xlrd
from xlutils.copy import copy
import xlwt


class FlightClass():
	def __init__(self,fn,ft,dpt,des,wzd=0,wzz=0,wlc=0,wyj=0,wgw=0,mzd=0,mzz=0,mlc=0,myj=0,cq = 'DOM',fq = "客机",rt=0,ezz=0):
		self.flight_num = fn #航班号
		self.flight_type = ft #机型
		self.depature = dpt #始发地
		self.destination = des #目的地
		self.weight_zhida = wzd
		self.weight_zhongzhuan = wzz
		self.weight_liancheng = wlc
		self.weight_youjian = wyj
		self.weight_gongwu = wgw
		self.money_zhida = mzd
		self.money_zhongzhuan = mzz
		self.money_liancheng = mlc
		self.money_youjian = myj
		self.cargo_quality = cq
		self.flight_quality = fq
		self.rate = rt
		self.earn_zhongzhuan = ezz
	def set_weight_zhida(self,w):
		self.weight_zhida+=w
	def set_weight_zhongzhuan(self,w):
		self.weight_zhongzhuan+=w		
	def set_weight_liancheng(self,w):
		self.weight_liancheng+=w
	def set_weight_youjian(self,w):
		self.weight_youjian+=w
	def set_weight_gongwu(self,w):
		self.weight_gongwu+=w
		
	def set_money_zhida(self,w):
		self.money_zhida+=w
	def set_money_zhongzhuan(self,w):
		self.money_zhongzhuan+=w		
	def set_money_liancheng(self,w):
		self.money_liancheng+=w
	def set_money_youjian(self,w):
		self.money_youjian+=w
	
	def set_cargo_quality(self):
		if self.destination in int_place:
			self.cargo_quality = '国际'

	def set_rate(self):
		if self.weight_zhida+self.weight_zhongzhuan+\
		self.weight_liancheng+self.weight_youjian+self.weight_gongwu!=0:
			self.rate = (self.money_zhida+self.money_zhongzhuan+\
			self.money_liancheng+self.money_youjian)/(self.weight_zhida\
			+self.weight_zhongzhuan+self.weight_liancheng\
			+self.weight_youjian+self.weight_gongwu)
			self.rate = round(self.rate,2)
	def set_earn_zhongzhuan(self,m):
		self.earn_zhongzhuan += m
		
	def show_flight_data(self):
		print('航班号：',self.flight_num)
		print('机型：',self.flight_type)
		print('始发地：',self.depature)
		print('目的地：',self.destination)
		print('货物性质：',self.cargo_quality)
		print('直达运量:',self.weight_zhida)
		print('费率:',self.rate)

def earn_income(dep,arr,des,mon,sheet):  #计算分摊收入，参数为始发地，中转地，目的地，收入，分摊.xls的数据页
	
	i=1
	first_dis=0
	second_dis=0
	dic_place={
	'pek':'bjs',
	#'pvg':'sha',
	'icn':'sel',
	'ord':'chi',
	'bka':'mow',
	'PEK':'bjs',
	#'PVG':'sha',
	'ICN':'sel',
	'ORD':'chi',
	'BKA':'mow',
	'HHA':'csx',
	'hha':'csx',	
	}
	if dep in dic_place.keys():
		dep = dic_place[dep]
	if arr in dic_place.keys():
		arr = dic_place[arr]
	if des in dic_place.keys():
		des = dic_place[des]
	while i<sheet.nrows:
		if sheet.cell_value(i,0)==dep.upper() and sheet.cell_value(i,1)==arr.upper():
			first_dis = sheet.cell_value(i,2)
			break
		i+=1
	i=1
	while i<sheet.nrows:
		if sheet.cell_value(i,0)==arr.upper() and sheet.cell_value(i,1)==des.upper():
			second_dis=sheet.cell_value(i,2)
			break
		i+=1	
	sum_dis = first_dis + second_dis
	if sum_dis==0:
		print('error',dep,arr,des)
		tk.messagebox.showinfo(title='', message='从'+dep+'经'+arr+'至'+des+'中转的货物无法计算')
		return 0
	else:
		return round(mon*first_dis/sum_dis,1)
			
	
if __name__ == "__main__":
	
	result = earn_income('cgq','can','SWA',10000)
	print(result)

	
