import xlwt, xlrd
import xlutils.copy
import requests
from requests.exceptions import RequestException

def initExcel(row0, name):
	f = xlwt.Workbook()
	sheet1 = f.add_sheet('sheet1', cell_overwrite_ok=True)
	for i in range(0,len(row0)):
		sheet1.write(0, i, row0[i])
	f.save(name)

def writeToExcel(name, x, y, res):
	f = xlrd.open_workbook(name)
	ws = xlutils.copy.copy(f)
	table = ws.get_sheet(0)
	table.write(x, y, res)
	ws.save(name)

def getOnePage(url, deco):
	try:
		headers = {
			'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'
		}
		response = requests.get(url, headers = headers)
		if response.status_code == 200:
			return response.content.decode(deco)
		else:
			print(response)
		return None
	except RequestException:
		return None

def saveOneFile(file_name, url):
	headers = {
		'User-Agent': 'Mozilla/4.0 (compatible; MSIE 5.5; Windows NT)'
	}
	resp = requests.get(url, headers = headers)
	with open('Results/' + file_name, 'wb') as f:
		f.write(resp.content)
		print('Already Downloaded', file_name)
