import re
import xlwt, xlrd
import xlutils.copy
import MyRanking as mr

PATTERN = "container right p10 w90p.*?location.href='(.*?)'"

def QS0():
	for i in range(11):
		url = 'https://www.universityrankings.ch/results?ranking=QS&region=Asia&year=20' + str(10 + i) + '&q=China+'
		html = mr.getOnePage(url, 'utf-8')
		pattern = re.compile(PATTERN, re.S)
		result = re.findall(pattern, html)
		url = 'https://www.universityrankings.ch' + result[0]
		file_name = 'QS20' + s + '.csv'
		mr.saveOneFile(file_name, url)

def QS():
	for i in range(11):
		s = str(10 + i)
		url = 'https://www.universityrankings.ch/results?ranking=QS&region=Asia&year=20' + s + '&mode=csv'
		file_name = 'QS20' + s + '.csv'
		mr.saveOneFile(file_name, url)

# QS()
