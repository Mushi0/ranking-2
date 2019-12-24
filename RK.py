import re
import xlwt, xlrd
import xlutils.copy
import MyRanking as mr

URL = 'http://www.zuihaodaxue.com/zuihaodaxuepaiming201'

ROW = ['排名', '学校名称', '省份', '总分', '生源质量(新生高考成绩得分)', '培养结果(毕业生就业率)', '社会声誉(社会捐赠收入-千元)', '科研规模(论文数量-篇)', '科研质量(论文质量-FWCI)', '顶尖成果(高被引论文-篇)', '顶尖人才(高被引学者-人)', '科技服务(企业科研经费-千元)', '成果转化(技术转让收入-千元)', '学生国际化(留学生比例)']
ROW2017 = ['排名', '学校名称', '省份', '总分', '生源质量(新生高考成绩得分)', '培养结果(毕业生就业率)', '科研规模(论文数量-篇)', '科研质量(论文质量-FWCI)', '顶尖成果(高被引论文-篇)', '顶尖人才(高被引学者-人)', '科技服务(企业科研经费-千元)', '成果转化(技术转让收入-千元)', '学生国际化(留学生比例)']
ROW2016 = ['排名', '学校名称', '省份', '总分', '生源质量(新生高考成绩得分)', '培养结果(毕业生就业率)', '科研规模(论文数量-篇)', '科研质量(论文质量-FWCI)', '顶尖成果(高被引论文-篇)', '顶尖人才(高被引学者-人)', '科技服务(企业科研经费-千元)', '产学研合作(校企合作论文-篇)', '成果转化(技术转让收入-千元)']

PATTERN = '<tr class="alt"><td>(.*?)</td><td><div align="left">(.*?)</div></td><td>(.*?)</td><td>(.*?)</td><td class="hidden-xs need-hidden indicator5">(.*?)</td><td class="hidden-xs need-hidden indicator6"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator7"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator8"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator9"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator10"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator11"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator12"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator13"style="display: none;">(.*?)</td><td class="hidden-xs need-hidden indicator14"style="display: none;">(.*?)</td></tr>'
PATTERN2017 = '<tr.*?><td>(.*?)<td><div align="left">(.*?)</div></td><td>(.*?)</td><td>(.*?)</td><td class="hidden-xs need-hidden indicator5">(.*?)</td><td class="hidden-xs need-hidden indicator6"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator7"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator8"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator9"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator10"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator11"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator12"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator13"  style="display:none;">(.*?)</td></tr>'
PATTERN2016 = '<tr.*?><td>(.*?)</td>[\s]*?<td><div align="left">(.*?)</div></td>[\s]*?<td>(.*?)</td><td>(.*?)</td><td class="hidden-xs need-hidden indicator5">(.*?)</td><td class="hidden-xs need-hidden indicator6"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator7"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator8"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator9"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator10"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator11"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator12"  style="display:none;">(.*?)</td><td class="hidden-xs need-hidden indicator13"  style="display:none;">(.*?)</td></tr>'

def RK_ZHDX(x, row, pt):
	name = 'Results/RK201' + str(x) + '.xls'
	print('Writing to ' + name + ' ... ...')
	mr.initExcel(row, name)
	url = URL + str(x) + '.html'
	html = mr.getOnePage(url, 'utf-8')
	pattern = re.compile(pt, re.S)
	result = re.findall(pattern, html)
	t = 1
	l = len(result[0])
	for item in result:
		for i in range(l):
			mr.writeToExcel(name, t, i, item[i])
		print(t)
		t += 1

# RK_ZHDX(9, ROW, PATTERN)
# RK_ZHDX(8, ROW, PATTERN)
# RK_ZHDX(7, ROW2017, PATTERN2017)
# RK_ZHDX(6, ROW2016, PATTERN2016)
