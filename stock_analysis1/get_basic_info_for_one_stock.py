# -*- coding: cp936 -*-
import urllib2

url = 'http://hq.sinajs.cn/list=sh601003'
headers = {"User-Agent":"Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6"}
req = urllib2.Request( url, headers = headers)
content = urllib2.urlopen(req).read()
feedback = content.decode('gbk')
data = feedback.split('"')[1].split(',')
name = '%-6s' % data[0]
price_current = '%-6s' % float(data[3])
change_percent = (float(data[3])-float(data[2]))*100/float(data[2])
change_percent = '%-6s' % round(change_percent,2)
#print("股票名称:{0} 涨跌幅:{1} 最新价:{2}".format(name,change_percent,price_current))
print name,change_percent,price_current
