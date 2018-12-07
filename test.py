import requests
import sys
import io
import json
import time
import xlwt

# sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='utf8') #改变标准输出的默认编码
#登录后才能访问的网页
url = 'http://tm.bnuz.edu.cn/api/place/buildings/'
#浏览器登录后得到的cookie，也就是刚才复制的字符串
cookie_str = r''
#把cookie字符串处理成字典，以便接下来使用
cookies = {}
for line in cookie_str.split(';'):
    key, value = line.split('=', 1)
    cookies[key] = value
#设置请求头
headers = {'User-agent':'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36'}
#在发送get请求时带上请求头和cookies
resp = requests.get(url, headers = headers, cookies = cookies)
jsondata = json.loads(resp.content.decode('utf-8'))
print(jsondata['buildings'])
arr = [];

book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('test', cell_overwrite_ok=True)

for name in jsondata['buildings']:
    url2 = url + name + '/places/';
    print(url2);
    resp2 = requests.get(url2, headers = headers, cookies = cookies);
    jsondata2 = json.loads(resp2.content.decode('utf-8'))
    for classname in jsondata2:
        url3 = url2 + classname['id'] + '/usages'
        print(url3)
        resp3 = requests.get(url3, headers = headers, cookies = cookies);
        jsondata3 = json.loads(resp3.content.decode('utf-8'))
        for item in jsondata3:
            if item['type'] == 'ks':
                arr.append(item);
for index,item in enumerate(arr):
    for index2,err in enumerate(item):
        sheet.write(index,index2,item[err])
book.save(r'd:\test1.xls')  # 在字符串前加r，声明为raw字符串，这样就不会处理其中的转义了。否则，可能会报错