import re
import xlwt, xlrd
import xlutils.copy
import MyRanking as mr

FILENAME2019 = 'Results/XYH2019.xls'
FILENAME2017 = 'Results/XYH2017-2018.xls'
FILENAME2016 = 'Results/XYH2016.xls'
FILENAME2015 = 'Results/XYH2015.xls'
FILENAME2014 = 'Results/XYH2014.xls'
FILENAME2013 = 'Results/XYH2013.xls'

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

def XYH1(row, name, url, pt):
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
		t += 1

def XYH2017():
	mr.initExcel(ROW2017, FILENAME2017)
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
				mr.writeToExcel(FILENAME2017, t, i, "".join(r))
			elif i == 0 or i == 4:
				r = re.findall(pattern2, item[i])
				mr.writeToExcel(FILENAME2017, t, i, r[0])
			else:
				mr.writeToExcel(FILENAME2017, t, i, item[i])
		t += 1

def XYH2(row, name, url, pt):
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
		t += 1

# XYH1(ROW2019, FILENAME2019, '5463.html', PATTERN2019)
# XYH2017()
# XYH2(ROW2016, FILENAME2016, '27207.html', PATTERN2016)
# XYH2(ROW2016, FILENAME2015, '5808.html', PATTERN2015)
# XYH2(ROW2014, FILENAME2014, '1382.html', PATTERN2014)
# XYH1(ROW2013, FILENAME2013, '1386.html', PATTERN2013)
