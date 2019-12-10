import re
import xlwt, xlrd
import xlutils.copy
import MyRanking as mr

FILENAME2018 = 'Results/WSL2018-2019.xls'
FILENAME2015 = 'Results/WSL2015.xls'
FILENAME2014 = 'Results/WSL2014.xls'
FILENAME2013 = 'Results/WSL2013.xls'

ROW2018 = ['排名', '学校名称', '类型', '所在地', '综合得分']
ROW2015 = ['2015排名', '学校名称', '总得分', '研究生培养', '本科生培养', '自然科学研究', '社会科学研究']
ROW2014 = ['排名', '学校名称', '总得分', '人才培养', '科学研究']

PATTERN2015 = '<tr><td valign="bottom" width="39" ""><p align="center">(.*?)</p></td><td valign="bottom" width="127" ""><p style="text-align:center;">(.*?)</p></td><td valign="bottom" width="45" ""><p align="right" style="text-align:center;">(.*?)</p></td><td valign="bottom" width="41" ""><p align="right" style="text-align:center;">(.*?)</p></td><td valign="bottom" width="41" ""><p align="right" style="text-align:center;">(.*?)</p></td><td valign="bottom" width="41" ""><p align="right" style="text-align:center;">(.*?)</p></td><td valign="bottom" width="41" ""><p align="right" style="text-align:center;">(.*?)</p></td></tr>'
PATTERN2014 = '</tr> <tr height="14"> <td height="14" width="33">(.*?)</td> <td width="128">(.*?)</td> <td width="97">(.*?)</td> <td width="70">(.*?)</td> <td width="78">(.*?)</td> </tr>'
PATTERN2013 = '<tr height="15"> <td height="15" width="47">(.*?)</td> <td width="119">(.*?)</td> <td width="59">(.*?)</td> <td width="58">(.*?)</td> <td width="62">(.*?)</td> </tr>'

def WSL2018():
	mr.initExcel(ROW2018, FILENAME2018)
	url = 'https://www.dxsbb.com/news/46702.html'
	html = mr.getOnePage(url, 'gbk')
	pattern = re.compile('<tr><td style="text-align:center;">(.*?)</td><td style="text-align:center;">(.*?)</td><td style="text-align:center;">(.*?)</td><td style="text-align:center;">(.*?)</td><td style="text-align:center;">(.*?)</td></tr>', re.S)
	result = re.findall(pattern, html)
	t = 1
	pattern1 = re.compile("[\u4e00-\u9fa5]+", re.S)
	for item in result:
		for i in range(5):
			if i == 1 or i == 3:
				r = re.findall(pattern1, item[i])
				mr.writeToExcel(FILENAME2018, t, i, "".join(r))
			else:
				mr.writeToExcel(FILENAME2018, t, i, item[i])
		t += 1

def WSL2(row, name, url, pt):
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

# WSL2018()
# WSL2(ROW2015, FILENAME2015, '6119.html', PATTERN2015)
# WSL2(ROW2014, FILENAME2014, '1387.html', PATTERN2014)
# WSL2(ROW2014, FILENAME2013, '1389.html', PATTERN2013)
