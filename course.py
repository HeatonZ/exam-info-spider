import requests
import sys
import io
import json
import time
import xlwt
import re

# 浏览器登录后得到的cookie，也就是刚才复制的字符串
cookie_str = r''
# 把cookie字符串处理成字典，以便接下来使用
cookies = {}
for line in cookie_str.split(';'):
    key, value = line.split('=', 1)
    cookies[key] = value
# 设置请求头
headers = {
    'User-agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36'}
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('test', cell_overwrite_ok=True)
style = xlwt.easyxf('align: wrap on')

sheet.write(0, 0, '节次')
sheet.write(0, 1, '星期一')
sheet.write(0, 2, '星期二')
sheet.write(0, 3, '星期三')
sheet.write(0, 4, '星期四')
sheet.write(0, 5, '星期五')
sheet.write(1, 0, '1')
sheet.write(2, 0, '2')
sheet.write(3, 0, '3')
sheet.write(4, 0, '4')
sheet.write(5, 0, '5')
sheet.write(6, 0, '6')
sheet.write(7, 0, '7')
sheet.write(8, 0, '8')
sheet.write(9, 0, '9')
sheet.write(10, 0, '10')
sheet.write(11, 0, '11')
sheet.write(12, 0, '12')
sheet.write(13, 0, '13')
# 在发送get请求时带上请求头和cookies
# xx = ['信息技术学院', '管理学院', '不动产学院', '文学院', '教育学院', '艺术与传播学院', '法律与行政学院', '物流学院', '特许经营学院', '设计学院', '外国语学院', '应用数学学院', '工程技术学院',
#       '国际商学部', '运动休闲学院', '中加合作办学项目', '宋庆龄公益慈善教育中心', '通识中心', '公共体育教研部', '政治理论教研部', '教务处', '科研处', '图书馆', '国内合作办公室', '资产处', '保卫处']
xx = ['国际商学部']
url1 = 'http://es.bnuz.edu.cn/eam/WebService.asmx/getTask_teacher'
url2 = 'http://es.bnuz.edu.cn/eam/WebService.asmx/getTask_info_teachers'
xn = '2018-2019'
xq = 2
JXBMC = '16会计01'
d = {"0":"","1":"单","2":"双"}
for item in xx:
    resp = requests.get(url1, headers=headers, cookies=cookies, params={
                        'bm': item, 'kkxy': item, 'type': 1, 'xn': xn, 'xq': 1})
    teachers = json.loads(resp.content.decode('utf-8'))
    zgh = ""
    for teacher in teachers:
        d[teacher['ZGH']] = teacher['XM']
        zgh = zgh + "\"" + teacher['ZGH'] + "\","
    zgh = zgh[:-1]
    res = requests.get(url2, headers=headers, cookies=cookies,
                       params={'zgh': zgh, 'xn': xn, 'xq': 2})
    courses = json.loads(res.content.decode('utf-8'))
    for course in courses:
        if re.match(JXBMC,course['JXBMC']) != None:
            sheet.write_merge(course['QSSJD'], course['QSSJD'] + course['SKCD'] - 1, course['XQJ'], course['XQJ'], course['KCMC'] +
                              '  ' + d[course['ZGH']] + '   ' + '【第' + str(course['QSZ'])+'-' + str(course['JSZ'])+d[str(course['DSZVALUE'])]+'周'+course['JSMC']+'】',style)

book.save('d:\\'+JXBMC+'.xls')  # 在字符串前加r，声明为raw字符串，这样就不会处理其中的转义了。否则，可能会报错
