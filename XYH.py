import re
import xlwt, xlrd
import xlutils.copy
import MyRanking as mr

ROW2019 = ['排名', '学校名称', '综合得分', '星级排名', '办学层次']
ROW2017 = ['排名', '学校名称', '星级排名', '办学层次', '总分']
ROW2016 = ['排名', '学校名称', '所在地区', '总分']
ROW2014 = ['排名', '学校名称', '所在地区', '总分', '科学研究', '人才培养']
ROW2013 = ['排名', '学校名称', '总分', '科学研究', '人才培养']

PATTERN2019 = '<tr height="19"><td width="69">(.*?)</td><td width="160">(.*?)</td><td width="85">(.*?)</td><td width="84">(.*?)</td><td width="230">(.*?)</td></tr>'
PATTERN2016 = '<tr> <td>(.*?)</td> <td>(.*?)</td> <td>(.*?)</td> <td>(.*?)</td> </tr>'
PATTERN2015 = '<tr height="15"><td height="15" align="right" style="text-align:center;">(.*?)</td><td style="text-align:center;">(.*?)</td><td style="text-align:center;">(.*?)</td><td align="right" style="text-align:center;">(.*?)</td></tr>'
PATTERN2014 = '<tr height="20"> <td height="20">(.*?)</td> <td>(.*?)</td> <td>(.*?)</td> <td>(.*?)</td> <td>(.*?)</td> <td>(.*?)</td> </tr>'
PATTERN2013 = '<tr height="15" style="height:15.0pt;"> <td height="15" class="xl7\d" width="54" style="height:15.0pt;border-top:none;width:54pt;">(.*?)</td> <td class="xl7\d" width="100" style="border-top:none;border-left:none;width:100pt;">(.*?)</td> <td class="xl7\d" width="54" style="border-top:none;border-left:none;width:54pt;">(.*?)</td> <td class="xl7\d" width="54" style="border-top:none;border-left:none;width:54pt;">(.*?)</td> <td class="xl7\d" width="54" style="border-top:none;border-left:none;width:54pt;">(.*?)</td> </tr> '

def XYH1(row, x, url, pt):
	name = 'Results/XYH201' + str(x) + '.xls'
	print('Writing to ' + name + ' ... ...')
	mr.initExcel(row, name)
	url = 'https://www.dxsbb.com/news/' + url
	html = mr.getOnePage(url, 'gbk')
	pattern = re.compile(pt, re.S)
	result = re.findall(pattern, html)
	t = 1
	l = len(result[0])
	pattern1 = re.compile("[\u4e00-\u9fa5]+", re.S)
	for item in result:
		for i in range(l):
			if i == 1:
				r = re.findall(pattern1, item[i])
				mr.writeToExcel(name, t, i, "".join(r))
			else:
				mr.writeToExcel(name, t, i, item[i])
		print(t)
		t += 1

def XYH2017():
	name = 'Results/XYH2017-2018.xls'
	print('Writing to ' + name + ' ... ...')
	mr.initExcel(ROW2017, name)
	url = 'https://www.dxsbb.com/news/1383.html'
	html = mr.getOnePage(url, 'gbk')
	pattern = re.compile('<tr height="19"><td x:num="(.*?)</td><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td><td x:num="(.*?)</td></tr>', re.S)
	result = re.findall(pattern, html)
	t = 1
	pattern1 = re.compile("[\u4e00-\u9fa5]+", re.S)
	pattern2 = re.compile("\d+\.?\d*", re.S)
	for item in result:
		for i in range(5):
			if i == 1 :
				r = re.findall(pattern1, item[i])
				mr.writeToExcel(name, t, i, "".join(r))
			elif i == 0 or i == 4:
				r = re.findall(pattern2, item[i])
				mr.writeToExcel(name, t, i, r[0])
			else:
				mr.writeToExcel(name, t, i, item[i])
		print(t)
		t += 1

def XYH2(row, x, url, pt):
	name = 'Results/XYH201' + str(x) + '.xls'
	print('Writing to ' + name + ' ... ...')
	mr.initExcel(row, name)
	url = 'https://www.dxsbb.com/news/' + url
	html = mr.getOnePage(url, 'gbk')
	pattern = re.compile(pt, re.S)
	result = re.findall(pattern, html)
	t = 1
	l = len(result[0])
	pattern1 = re.compile("[\u4e00-\u9fa5]+", re.S)
	for item in result:
		for i in range(l):
			if i == 1 or i == 2 :
				r = re.findall(pattern1, item[i])
				mr.writeToExcel(name, t, i, "".join(r))
			else:
				mr.writeToExcel(name, t, i, item[i])
		print(t)
		t += 1

# XYH1(ROW2019, 9, '5463.html', PATTERN2019)
# XYH2017()
# XYH2(ROW2016, 6, '27207.html', PATTERN2016)
# XYH2(ROW2016, 5, '5808.html', PATTERN2015)
# XYH2(ROW2014, 4, '1382.html', PATTERN2014)
# XYH1(ROW2013, 3, '1386.html', PATTERN2013)
