import requests
import json
import random
import requests.utils
import math
from urllib import request
from bs4 import BeautifulSoup
import xlwt
from xlwt import Workbook

excel_name = 'Taobao_Comments.xls'
sheet_name = 'Taobao_Comments'

Taobao_Comments_excel = Workbook(excel_name)
Taobao_Comments_excel = Workbook(encoding = 'utf-8')
Taobao_Comments_sheet = Taobao_Comments_excel.add_sheet(sheet_name,cell_overwrite_ok=True)
Taobao_Comments_sheet.write(0,0,u'序号')
Taobao_Comments_sheet.write(0,1,u'USER')
Taobao_Comments_sheet.write(0,2,u'COMMENTS')
Taobao_Comments_excel.save(excel_name)





#爬虫头
Cookie=" "#real cookie
header={"Connection": "keep-alive",
"Upgrade-Insecure-Requests": "1",
"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134",
"Cookie":Cookie,
"Accept":"text/html, application/xhtml+xml, application/xml; q=0.9, */*; q=0.8",
"Accept-Encoding":"gzip, deflate, br",
"Accept-Language":"zh-Hans-CN, zh-Hans; q=0.5",
'Connection': 'keep-alive'}
#代理列表（快代理可以更新）
proxy_list = [{"http" : "http://119.101.113.178:9999"},
    {"http" : "http://119.101.116.166:9999"},
    {"http" : "http://111.177.168.155:9999"},
    {"http" : "http://121.232.148.60:9000"},
    {"http" : "http://59.37.33.62:50686"},
    {"http":"http://119.101.113.240:9999"},
    {"http":"http://223.223.187.195:80"},
    {"http":"http://116.192.175.93:41944"},
    {"http":"http://193.112.57.222:8118"},
    {"http":"http://223.223.187.195:80"}] #just for test 
#登陆信息
post_data = {
    "commit": "Sign in",
    #"utf8": "✓",
    "authenticity_token": " ", #search in chrome for unique authenticity_token
    "login": " ", #real account and key
    "password": " "
}
proxy=random.choice(proxy_list)

def getmax(url):
    if url[url.find('id=')+14]!='&':
        id = url[url.find('id=')+3:url.find('id=')+15]
    else:
        id = url[url.find('id=')+3:url.find('id=')+14]
    url='https://rate.taobao.com/feedRateList.htm?auctionNumId='+id+'&currentPageNum=1'
    login('https://login.taobao.com/member/login.jhtml?style=mini&from=sm&full_redirect=false&redirect')
    s=requests.Session()
    html=s.get(url,headers=header,proxies=proxy).text
    jc=json.loads(html.strip().strip('()'))
    max=jc['total']   
    return max

def getUserComment(url,n,row):
    if url[url.find('id=')+14]!='&':
        id = url[url.find('id=')+3:url.find('id=')+15]
    else:
        id = url[url.find('id=')+3:url.find('id=')+14]
    url='https://rate.taobao.com/feedRateList.htm?auctionNumId='+id+'&currentPageNum=%d'%n
    login('https://login.taobao.com/member/login.jhtml?style=mini&from=sm&full_redirect=false&redirect')
    s=requests.Session()
    html=s.get(url,headers=header,proxies=proxy).text
    jc=json.loads(html.strip().strip('()'))
    max=jc['total']
    a = max
    users=[]
    comment=[]  
    count=0
    jc=jc['comments']
    #row = 1 + ( n - 1 ) * 20
    str1 = '评价方未及时做出评价,系统默认好评!'
    str2 = '此用户没有填写评价。'
    for j in jc:
        #if(j['content']!= str1 and j['content'] != str2):
        users.append(j['user']['nick'])
        comment.append( j['content'])
        #print(count+1,'>>',users[count],'\n        ',comment[count])
        if(comment[count] != str1 and comment[count] != str2):
            Taobao_Comments_sheet.write(row,0,row)
            Taobao_Comments_sheet.write(row,1,users[count])
            Taobao_Comments_sheet.write(row,2,comment[count])
            row += 1
        count=count+1
    Taobao_Comments_excel.save(excel_name)
    return row

    

#获取登陆的session，填入登陆信息
def login(url):   
    session=requests.session()
    response=session.post(url,data=post_data)
#主程序
a=getmax("https://item.taobao.com/item.htm?id=521341279470")
a=math.ceil(a/20)
row = 1
for i in range(1,a):
    row = getUserComment("https://item.taobao.com/item.htm?id=521341279470",i,row)
    

