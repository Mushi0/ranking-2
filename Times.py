from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import time
from selenium.webdriver.support.select import Select
import re
import xlwt, xlrd
import xlutils.copy
import MyRanking as mr

ROW = ['Rank', 'Name', 'NO. of FTE Students', 'NO. of Students Per Staff', 'International Students', 'Female:Male', 'Score:Overall', 'Score:Teaching', 'Score:Research', 'Score:Citations', 'Score:Industry Income', 'Score:International Outlook']

URL1 = 'https://www.timeshighereducation.com/world-university-rankings/201'
URL2 = '/world-ranking#!/page/0/length/25/locations/CN/sort_by/rank/sort_order/asc/cols/stats'
URL3 = '/world-ranking#!/page/0/length/25/locations/CN/sort_by/rank/sort_order/asc/cols/scores'

PATTERN1 = 
PATTERN2 = 

def Times(x):
	name = 'Results/Times201' + str(x) + '.xls'
	print('Writing to ' + name + ' ... ...')
	mr.initExcel(ROW, name)
	url = URL1 + str(x) + URL2
	html = mr.getOnePage(url, 'utf-8')
	pattern = re.compile(PATTERN1, re.S)
	result = re.findall(pattern, html)
	t = 1
	l = len(result[0])
	for item in result:
		for i in range(l):
			mr.writeToExcel(name, t, i, item[i])
		print(t)
		t += 1
	url = URL1 + str(x) + URL3
	html = mr.getOnePage(url, 'utf-8')
	pattern = re.compile(PATTERN2, re.S)
	result = re.findall(pattern, html)
	t = 1
	l = len(result[0])
	for item in result:
		for i in range(l):
			mr.writeToExcel(name, t, i + 6, item[i])
		print(t)
		t += 1

'''def Times(x):
	name = 'Results/Times201' + str(x) + '.xls'
	print('Writing to ' + name + ' ... ...')
	mr.initExcel(ROW, name)
	url = URL1 + str(x) + URL2
	browser = webdriver.Chrome()
	browser.get(url)
	Selections = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'location')))
	Select(Selections).select_by_value('CN')
	tables = WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.ID, 'datatable-1')))
	print(tables.get_attribute('textContent'))'''

Times(9)
