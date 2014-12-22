# -*- coding: gbk -*-
import urllib2
import xlwt

wb = xlwt.Workbook(optimized_write = True)
ws1 = wb.add_sheet(u'sh_stock_data')
ws1.write(0, 0, u'股票编号')
ws1.write(0, 1, u'股票名称')
ws1.write(0, 2, u'涨跌幅')
ws1.write(0, 3, u'当前价格')

ws2 = wb.add_sheet(u'sz_stock_data')
ws2.write(0, 0, u'股票编号')
ws2.write(0, 1, u'股票名称')
ws2.write(0, 2, u'涨跌幅')
ws2.write(0, 3, u'当前价格')

def get_one_stock_info(stock_list,row_offset,sheet):
    #count = len(stock_list)
    stock_num = ','.join(stock_list)
    url = 'http://hq.sinajs.cn/list='+stock_num
    headers = {"User-Agent":"Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6"}
    req = urllib2.Request( url, headers = headers)
    content = urllib2.urlopen(req).read()
    feedback = content.decode('gbk')
    stocks = feedback.split(';')
    stocks = stocks[:-1]
    useless_stock_num = []
    #for stock in stocks:
    for i in range(len(stocks)):
        #print stock
        #print "row_offset=%s,i=%s" % (row_offset,i)
        #print str(row_offset+i)
        is_useful = process_per_stock(stocks[i],row_offset+i,sheet)
        if is_useful != None:
            useless_stock_num.append(str(is_useful)[-9:-3])
    return useless_stock_num



def process_per_stock(stock,row_number,sheet):
    data = stock.split('"')[1].split(',')
    if len(data)>1:
        name = '%-6s' % data[0]
        price_current = '%-6s' % float(data[3])
        change_percent = (float(data[3])-float(data[2]))*100/float(data[2])
        change_percent = '%-6s' % round(change_percent,2)
        #print("股票名称:{0} 涨跌幅:{1} 最新价:{2}".format(name,change_percent,price_current))
        #print "股票名称:%s 涨跌幅:%s 最新价:%s" % (name,change_percent,price_current
        #print str(row_number+1)
        sheet.write(row_number+1, 0, stock[12:20])
        sheet.write(row_number+1, 1, name)
        sheet.write(row_number+1, 2, change_percent)
        sheet.write(row_number+1, 3, price_current)
        print name,change_percent,price_current
    else:
        return stock
        
        

def generate_sotck_list(count_start,count_end):
    sh_stock = []
    sz_stock = []
    count = count_start
    while count < count_end:
        sh_stock.append('sh'+str(count).zfill(6))
        sz_stock.append('sz'+str(count).zfill(6))
        count += 1
    return sh_stock,sz_stock

each_count = 500
count_max = 620000
for i in range(count_max/each_count):
    row_offset = i*each_count
    count_start = i*each_count
    count_end = (i+1)*each_count
    sh_stock,sz_stock = generate_sotck_list(count_start,count_end)
    print "欢迎来到上海股市"
    sh_useless = get_one_stock_info(sh_stock,row_offset,ws1)

    print "欢迎来到深圳股市"
    sz_useless = get_one_stock_info(sz_stock,row_offset,ws2)

wb.save('stock_basic_info.xls')
