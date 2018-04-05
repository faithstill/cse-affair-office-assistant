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
work_head = [u'发布时间',u'项目',u'信息来源']
workbook = Workbook(encoding = 'ascii')
worksheet = workbook.add_sheet("test")
worksheet.col(0).width = 7000
worksheet.col(1).width = 7000
worksheet.col(2).width = 10000
for i in range(0,3):
	worksheet.write(0,i,work_head[i])
driver = webdriver.Chrome()
driver.get('http://job.neu.edu.cn/pages/Frt/FrontMoreNewsListPage.aspx?PlateId=24')
driver.maximize_window() 
h = driver.current_window_handle
id_first='ctl00_ContentPlaceHolder1_FrontMoreNewList1_DataList1_ctl'
id_num2=0
id_num=-1
id_last='_LinkButton1'
id_car=[]
tag = []

while(the_page > 1):
	driver.find_element_by_id('ctl00_ContentPlaceHolder1_FrontMoreNewList1_NextPage').click()
	time.sleep(3)
	the_page = the_page-1
	h_all = driver.window_handles
	driver.switch_to_window(h_all[-2])
	driver.close()
	driver.switch_to_window(h_all[-1])
#driver.switch_to_window(h_all[-1])
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

#print x
#print the_Lpalce
#print the_addpage
if the_addpage == 0:
	for i in range(the_Bplace,the_Lpalce):
		tag =[]
		driver.find_element_by_id(id_car[i]).click()
		time.sleep(3)
		h_all= driver.window_handles
		print h_all
		driver.switch_to_window(h_all[1])
		title = driver.title#标题
		print title
		bs=bs4(driver.page_source,"html.parser")
		pubdate = bs.select('#ctl00_ContentPlaceHolder1_newsTimeLabel')[0].getText()#发布时间
		url= driver.current_url # 信息来源
		worksheet.write(information_num+1,0,pubdate[5:15])
		worksheet.write(information_num+1,1,title)
		worksheet.write(information_num+1,2,url)
		information_num=information_num+1
		driver.close()
		time.sleep(3)
		h_all= driver.window_handles
		print h_all
		del tag
		driver.switch_to_window(h_all[0])
		print driver.title

elif (the_addpage > 0):
	for i in range(the_Bplace,20):
		tag =[]
		driver.find_element_by_id(id_car[i]).click()
		time.sleep(3)
		h_all= driver.window_handles
		driver.switch_to_window(h_all[1])
		title = driver.title#标题
		bs=bs4(driver.page_source,"html.parser")
		pubdate = bs.select('#ctl00_ContentPlaceHolder1_newsTimeLabel')[0].getText()#发布时间
		url= driver.current_url # 信息来源
		worksheet.write(information_num+1,0,pubdate[5:15])
		worksheet.write(information_num+1,1,title)
		worksheet.write(information_num+1,2,url)
		information_num=information_num+1
		driver.close()
		time.sleep(3)
		del tag
		driver.switch_to_window(h_all[0])
	if (the_addpage == 1):
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
			url= driver.current_url # 信息来源
			worksheet.write(information_num+1,0,pubdate[5:15])
			worksheet.write(information_num+1,1,title)
			worksheet.write(information_num+1,2,url)
			information_num=information_num+1
			driver.close()
			time.sleep(3)
			del tag
			driver.switch_to_window(h_all[-2])
	elif (the_addpage > 1):
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
				url= driver.current_url # 信息来源
				worksheet.write(information_num+1,0,pubdate[5:15])
				worksheet.write(information_num+1,1,title)
				worksheet.write(information_num+1,2,url)
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
			url= driver.current_url # 信息来源
			worksheet.write(information_num+1,0,pubdate[5:15])
			worksheet.write(information_num+1,1,title)
			worksheet.write(information_num+1,2,url)
			information_num=information_num+1
			driver.close()
			time.sleep(3)
			del tag
			driver.switch_to_window(h_all[-2])
#name_xls=pubdate[5:15]+'.xls'
workbook.save('zuixinzhaoping.xls')
print "ok"