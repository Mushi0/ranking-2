import re
import xlwt, xlrd
import xlutils.copy
import MyRanking as mr

URL = 'http://www.zuihaodaxue.com/Greater_China_Ranking201'

ROW1 = ['排名', '学校名称', '地区', '总分']
ROW2 = ['人才培养', '科学研究', '师资质量', '学校质量']
RC2 = [3, 10, 18, 23, 25]
ROW_RCPY = ['研究生比例(5%)', '留学生比例(5%)', '师生比(5%)', '博士学位授予数(10%)', '校友获奖(10%)']
ROW3 = ['总量', '师均', '生均']
ROW_KXYJ = ['科研经费(5%)', '顶尖论文(10%)', '国际论文(10%)', '国际专利(10%)']
ROW_SZZL = ['博士学位教师比例(5%)', '教师获奖(10%)', '高被引科学家(10%)']
ROW_XXZY = '办学经费(5%)'

PATTERN = ['<tr><td>(.*?)</td><td class="align-left">.*?<div align="left">(.*?)</div></a></td><td>(.*?)</td><td>(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td></tr>', '<td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td></tr>', '<td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td></tr>', '<td class="hidden-xs">(.*?)</td><td class="hidden-xs">(.*?)</td></tr>']
PATTERN2 = ['<tr.*?>[\s]*?<td>(.*?)</td>[\s]*?<td class="align-left">[\s]*?<a h.*?>(.*?)</a>[\s]*?</td>[\s]*?<td>(.*?)</td>[\s]*?<td>(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>', '<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>', '<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>', '<td class="hidden-xs">(.*?)</td>[\s]*?<td class="hidden-xs">(.*?)</td>']

def initExcel_rk2(name):
	f = xlwt.Workbook()
	sheet1 = f.add_sheet('sheet1', cell_overwrite_ok=True)
	f.save(name)
	f = xlrd.open_workbook(name)
	ws = xlutils.copy.copy(f)
	table = ws.get_sheet(0)
	for i in range(4):
		table.write_merge(0, 2, i, i, ROW1[i])
	for i in range(4):
		t = RC2[i] + 1
		t2 = RC2[i + 1]
		table.write_merge(0, 0, t, t2, ROW2[i])
	for i in range(3):
		table.write_merge(1, 2, i + 4, i + 4, ROW_RCPY[i])
	for i in range(4, 6):
		table.write_merge(1, 1, i*2 - 1, i*2, ROW_RCPY[i - 1])
		table.write(2, i*2 - 1, ROW3[0])
		table.write(2, i*2, ROW3[1])
	table.write(2, 10, ROW3[2])
	for i in range(6, 10):
		table.write_merge(1, 1, i*2 - 1, i*2, ROW_KXYJ[i - 6])
		table.write(2, i*2 - 1, ROW3[0])
		table.write(2, i*2, ROW3[1])
	table.write_merge(1, 2, 19, 19, ROW_SZZL[0])
	for i in range(10, 12):
		table.write_merge(1, 1, i*2, i*2 + 1, ROW_SZZL[i - 9])
		table.write(2, i*2, ROW3[0])
		table.write(2, i*2 + 1, ROW3[1])
	table.write_merge(1, 1, 24, 25, ROW_XXZY)
	table.write(2, 24, ROW3[0])
	table.write(2, 25, ROW3[2])
	ws.save(name)

def RK_LASD(x, pt):
	name = 'Results/RK2_201' + str(x) + '.xls'
	print('Writing to ' + name + ' ... ...')
	initExcel_rk2(name)
	url = URL + str(x) + '_'
	for i in range(4):
		url1 = url + str(i) + '.html'
		html = mr.getOnePage(url1, 'utf-8')
		pattern = re.compile(pt[i], re.S)
		result = re.findall(pattern, html)
		if i == 0:
			t = 3
			l = len(result[0])
			for item in result:
				for j in range(l):
					mr.writeToExcel(name, t, j, item[j])
				print(t - 2)
				t += 1
		else:
			t = 3
			l = len(result[0])
			k = RC2[i] + 1
			for item in result:
				for j in range(l):
					mr.writeToExcel(name, t, j + k, item[j])
				print(t - 2)
				t += 1

# for i in range(6, 10):
# 	RK_LASD(i, PATTERN)
# for i in range(3, 6):
# 	RK_LASD(i, PATTERN2)
