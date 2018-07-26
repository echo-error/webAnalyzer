#coding:utf-8
import os
import pycurl
import xlsxwriter

URL = "http://www.126.com:80"

c = pycurl.Curl()
c.setopt(pycurl.URL,URL)
c.setopt(pycurl.CONNECTTIMEOUT,10)
c.setopt(pycurl.TIMEOUT,10)
c.setopt(pycurl.NOPROGRESS,1)
c.setopt(pycurl.FORBID_REUSE,1)
c.setopt(pycurl.MAXREDIRS,1)
c.setopt(pycurl.MAXREDIRS,1)
c.setopt(pycurl.DNS_CACHE_TIMEOUT,30)
indexfile = open(os.path.dirname(os.path.realpath(__file__))+"\content.txt","wb")
c.setopt(pycurl.WRITEHEADER,indexfile)
c.setopt(pycurl.WRITEDATA,indexfile)
c.perform()

NAMELOOKUP_TIME = c.getinfo(pycurl.NAMELOOKUP_TIME)
CONNECT_TIME = c.getinfo(pycurl.CONNECT_TIME)
TOTAL_TIME = c.getinfo(pycurl.TOTAL_TIME)
HTTP_CODE = c.getinfo(pycurl.HTTP_CODE)
SIZE_DOWNLOAD = c.getinfo(pycurl.SIZE_DOWNLOAD)
HEADER_SIZE = c.getinfo(pycurl.HEADER_SIZE)
SPEED_DOWNLOAD = c.getinfo(pycurl.SPEED_DOWNLOAD)
print u'HTTP状态码: %s ' %HTTP_CODE
print u'DNS解析时间: %.2f ms' %(NAMELOOKUP_TIME * 1000)
print u'建立连接时间：%.2f ms' %(CONNECT_TIME * 1000)
print u'传输结束总用时：%.2f ms' %(TOTAL_TIME * 1000)
print u'下载数据包大小: %d bytes' %(SIZE_DOWNLOAD)
print u'HTTP头部大小: %d byte' %(HEADER_SIZE)
print(u'下载速度：%d byte' %SPEED_DOWNLOAD)
indexfile.close()
c.close()

f = open(os.getcwd()+"/chart.txt",'a')
f.write(str(HTTP_CODE)+','+str(NAMELOOKUP_TIME * 1000)+','
        +str(CONNECT_TIME * 1000)+','+str(TOTAL_TIME * 1000)+','
        +str(SIZE_DOWNLOAD)+','+str(HEADER_SIZE)+','
        +str(SPEED_DOWNLOAD/1024)+'\n')
f.close()

workbook = xlsxwriter.Workbook(os.getcwd()+"\chart.xlsx")
worksheet = workbook.add_worksheet()
chart = workbook.add_chart({'type':'column'})
title =  [ URL,
          u'HTTP状态码',
          u'DNS解析时间',
          u'建立连接时间',
          u'传输结束总用时',
          u'下载数据包大小',
          u'HTTP头部大小',
          u'下载速度'
           ]

format = workbook.add_format()
format.set_border(1)
format.set_align('center')

format_title = workbook.add_format()
format.set_border(1)

format_title.set_bg_color('#00F00')
format_title.set_align('center')
format_title.set_bold()

worksheet.write_row(0,0,title,format_title)



f = open('chart.txt','r')
line = 1
for i in f:
    head = [line]
    linelist = i.split(',') ####chart.txt中的一行转为list型
    linelist = map(lambda mapi:int(float(mapi.replace("\n",''))),linelist) ####去掉每行中的换行符并整数化
    linelist = head + linelist
    worksheet.write_row(line,0,linelist,format)
    line = line + 1

average = [u'平均值',
           '=AVERAGE(B2:B' + str(line - 1) + ')',
           '=AVERAGE(C2:C' + str(line - 1) + ')',
           '=AVERAGE(D2:D' + str(line - 1) + ')',
           '=AVERAGE(E2:E' + str(line - 1) + ')',
           '=AVERAGE(F2:F' + str(line - 1) + ')',
           '=AVERAGE(G2:G' + str(line - 1) + ')',
           '=AVERAGE(H2:H' + str(line - 1) + ')',
           ]
worksheet.write_row(line,0,average,format)
f.close()

def chart_series(cur_row):
    chart.add_series(
        {
            'categories': '=Sheet1!$B$1:$H$1',
            'values'    : '=Sheet1!$B$'+cur_row+':$H$'+cur_row,
            'line'    : {'color':'black'}, ####图表柱边线颜色
            'name'      : '=Sheet1!$A$'+cur_row,

        }
    )
for row in range(2,line+1):
    chart_series(str(row))

chart.set_size({'width':876,'height':287})

worksheet.insert_chart(line + 7,0,chart)

workbook.close()


