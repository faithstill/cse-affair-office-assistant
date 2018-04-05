# -*- coding:UTF-8 -*-
import sys  
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import urllib
import requests
import string
import time
from bs4 import BeautifulSoup as bs4
from xlwt import Workbook, Formula
import xlrd
reload(sys)  
sys.setdefaultencoding('utf8')
total_num= input("How many pieces of information need to be filed today?")
the_page = input("What page does it start?")
the_Bplace = input("where the begin")
the_Bplace = the_Bplace -1
work_head = [u'发布时间',u'公司名称',u'招聘时间',u'招聘地点',u'信息来源']
workbook = Workbook(encoding = 'ascii')
worksheet = workbook.add_sheet("test")
worksheet.col(0).width = 7000
worksheet.col(1).width = 7000
worksheet.col(2).width = 7000
worksheet.col(3).width = 7000
worksheet.col(4).width = 10000
for i in range(0,5):
	worksheet.write(0,i,work_head[i])

driver = webdriver.Chrome() # Get local session of firefox
driver.get('http://job.neu.edu.cn/Pages/Frt/FrontMoreNewsListPage.aspx?PlateId=2')
h = driver.current_window_handle
id_first='ctl00_ContentPlaceHolder1_FrontMoreNewList1_DataList1_ctl'
id_num2=0
id_num=-1
id_last='_LinkButton1'
id_car=[]
tag = []
campus1=u'南湖校区'
campus2=u'浑南校区'
S1=Select(driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_DDL1'))
S1.select_by_index(1)
#print S1.all_selected_options
time.sleep(3)
driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_confirm').click()
time.sleep(2)
driver.close()
h_all = driver.window_handles
driver.switch_to_window(h_all[-1])
while(the_page > 1):
	driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_NextPage').click()
	time.sleep(7)
	the_page = the_page-1
	h_all = driver.window_handles
	driver.switch_to_window(h_all[-2])
	driver.close()
	driver.switch_to_window(h_all[-1])

while(id_num<19):
	id_num=id_num+1
	if id_num<10:
		f=id_first+str(id_num2)+str(id_num)+id_last
	else :
		f=id_first+str(id_num)+id_last
	id_car.append(f)

if (20-the_Bplace < total_num) :
	x=total_num-(20-the_Bplace)
	the_Lpalce=x%20
	the_addpage=x/20
	the_addpage=the_addpage+1
else :
	the_Lpalce = the_Bplace+total_num
	the_addpage = 0
information_num=1

if the_addpage == 0:
	for i in range(the_Bplace,the_Lpalce):
		tag =[]
		driver.find_element_by_id(id_car[i]).click()
		time.sleep(5)
		h_all= driver.window_handles
		driver.switch_to_window(h_all[1])
		title = driver.title#标题
		bs=bs4(driver.page_source,"html.parser")
		pubdate = bs.select('#ctl00_ContentPlaceHolder1_newsTimeLabel')[0].getText()#发布时间
		p=bs.p
		for item in bs.findAll('p'):
			temp = item.get_text()
			tag.append(item.get_text())
		hiringdata = tag[0]  #招聘时间
		hiringsite = tag[1]	 #招聘地点
		url= driver.current_url # 信息来源
		worksheet.write(information_num+1,0,pubdate[5:15])
		worksheet.write(information_num+1,1,title[12:])
		worksheet.write(information_num+1,2,hiringdata[6:])
		worksheet.write(information_num+1,3,hiringsite[6:])
		worksheet.write(information_num+1,4,url)
		information_num=information_num+1
		driver.close()
		time.sleep(5)
		del tag
		driver.switch_to_window(h_all[0])

elif (the_addpage > 0):
	for i in range(the_Bplace,20):
		tag =[]
		driver.find_element_by_id(id_car[i]).click()
		time.sleep(5)
		h_all= driver.window_handles
		driver.switch_to_window(h_all[1])
		title = driver.title#标题
		bs=bs4(driver.page_source,"html.parser")
		pubdate = bs.select('#ctl00_ContentPlaceHolder1_newsTimeLabel')[0].getText()#发布时间
		p=bs.p
		for item in bs.findAll('p'):
			temp = item.get_text()
			tag.append(item.get_text())
		hiringdata = tag[0]  #招聘时间
		hiringsite = tag[1]	 #招聘地点
		url= driver.current_url # 信息来源
		worksheet.write(information_num+1,0,pubdate[5:15])
		worksheet.write(information_num+1,1,title[12:])
		worksheet.write(information_num+1,2,hiringdata[6:])
		worksheet.write(information_num+1,3,hiringsite[6:])
		worksheet.write(information_num+1,4,url)
		information_num=information_num+1
		driver.close()
		time.sleep(5)
		del tag
		driver.switch_to_window(h_all[0])
	if(the_addpage == 1):
		driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_NextPage').click()
		time.sleep(3)
		h_all= driver.window_handles
		driver.switch_to_window(h_all[-1])
		for t in range(0,the_Lpalce):
			tag =[]
			driver.find_element_by_id(id_car[t]).click()
			time.sleep(3)
			h_all= driver.window_handles
			driver.switch_to_window(h_all[-1])
			title = driver.title#标题
			bs=bs4(driver.page_source,"html.parser")
			pubdate = bs.select('#ctl00_ContentPlaceHolder1_newsTimeLabel')[0].getText()#发布时间
			p=bs.p
			for item in bs.findAll('p'):
				temp = item.get_text()
				tag.append(item.get_text())
			hiringdata = tag[0]  #招聘时间
			hiringsite = tag[1]	 #招聘地点
			url= driver.current_url # 信息来源
			worksheet.write(information_num+1,0,pubdate[5:15])
			worksheet.write(information_num+1,1,title[12:])
			worksheet.write(information_num+1,2,hiringdata[6:])
			worksheet.write(information_num+1,3,hiringsite[6:])
			worksheet.write(information_num+1,4,url)
			information_num=information_num+1
			driver.close()
			time.sleep(3)
			del tag
			driver.switch_to_window(h_all[-2])
	elif(the_addpage > 1):
		for i in range(0,the_addpage-1):
			driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_NextPage').click()
			time.sleep(3)
			h_all= driver.window_handles
			driver.switch_to_window(h_all[-1])
			for t in range(0,20):
				tag =[]
				driver.find_element_by_id(id_car[t]).click()
				time.sleep(3)
				h_all= driver.window_handles
				driver.switch_to_window(h_all[-1])
				title = driver.title#标题
				bs=bs4(driver.page_source,"html.parser")
				pubdate = bs.select('#ctl00_ContentPlaceHolder1_newsTimeLabel')[0].getText()#发布时间
				p=bs.p
				for item in bs.findAll('p'):
					temp = item.get_text()
					tag.append(item.get_text())
				hiringdata = tag[0]  #招聘时间
				hiringsite = tag[1]	 #招聘地点
				url= driver.current_url # 信息来源
				worksheet.write(information_num+1,0,pubdate[5:15])
				worksheet.write(information_num+1,1,title[12:])
				worksheet.write(information_num+1,2,hiringdata[6:])
				worksheet.write(information_num+1,3,hiringsite[6:])
				worksheet.write(information_num+1,4,url)
				information_num=information_num+1
				driver.close()
				time.sleep(3)
				del tag
				driver.switch_to_window(h_all[-2])
		driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_NextPage').click()
		time.sleep(3)
		h_all= driver.window_handles
		driver.switch_to_window(h_all[-1])
		for i in range(0,the_Lpalce):
			tag =[]
			driver.find_element_by_id(id_car[i]).click()
			time.sleep(3)
			h_all= driver.window_handles
			driver.switch_to_window(h_all[-1])
			title = driver.title#标题
			bs=bs4(driver.page_source,"html.parser")
			pubdate = bs.select('#ctl00_ContentPlaceHolder1_newsTimeLabel')[0].getText()#发布时间
			p=bs.p
			for item in bs.findAll('p'):
				temp = item.get_text()
				tag.append(item.get_text())
			hiringdata = tag[0]  #招聘时间
			hiringsite = tag[1]	 #招聘地点
			url= driver.current_url # 信息来源
			worksheet.write(information_num+1,0,pubdate[5:15])
			worksheet.write(information_num+1,1,title[12:])
			worksheet.write(information_num+1,2,hiringdata[6:])
			worksheet.write(information_num+1,3,hiringsite[6:])
			worksheet.write(information_num+1,4,url)
			information_num=information_num+1
			driver.close()
			time.sleep(3)
			del tag
			driver.switch_to_window(h_all[-2])
#name_xls = pubdate[5:15] + '.xls' 
workbook.save('xiaoyuanzhaoping.xls')
print "ok"

#elem= driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_DataList1_ctl00_LinkButton1')
#elem = driver.find_element_by_id('kw')  
#elem.send_keys(u'php')  
# elem.send_keys(Keys.ENTER)  #点击键盘上的Enter按钮  
#driver.find_element_by_id('su').click()  # 点击了百度页面上的‘百度一下’按钮
#driver.forward()
#driver.refresh()
#driver.get_screenshot_as_file('123.txt')
#print('页面标题：', driver.title)
#print elem  # 页面标题
#print elem  
#print(driver.current_url)  # 当前页面url  
#print('搜索后的页面源码：\n', driver.page_source)  # 页面源码  

#print (pubdate+'\n')
#newcampus=campus.encode('gb2312')
	#print ("校区"+ newcampus+'\n')
#	print (title+'\n')
#	print (hiringdata+'\n')
#	print (hiringsite+'\n')
#	print (url+'\n')
#	'''