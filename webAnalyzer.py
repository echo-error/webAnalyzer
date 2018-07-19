
#coding:utf-8
import os
import sys
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
indexfile = open(os.path.dirname(os.path.realpath(__file__))+"/content.txt","wb")
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
        +str(SIZE_DOWNLOAD/1024)+','+str(HEADER_SIZE)+','
        +str(SPEED_DOWNLOAD/1024)+'\n')
f.close()

workbook = xlsxwriter.Workbook(os.getcwd()+"/chart.xlsx")
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
format.set_bold(1)








