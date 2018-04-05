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
driver = webdriver.Chrome()
driver.get('http://job.neu.edu.cn/pages/Frt/FrontMoreNewsListPage.aspx?PlateId=24') #就业网最新招聘首页
h = driver.current_window_handle
bs=bs4(driver.page_source,"html.parser")
fp = open ('123456.txt','w')
fp.writer(bs)
fp.close