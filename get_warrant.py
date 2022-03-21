# -*- coding: utf-8 -*-
import pymysql
import dbconfig
import openpyxl
import datetime

DISK='D:/'
d=-1
day = (datetime.datetime.now()+datetime.timedelta(days=d)).strftime("%Y%m%d")
path = DISK+day

dbconnect = dbconfig.setting
conn = pymysql.connect(**dbconnect)
cursor = conn.cursor()

for t in ['trading_detail','trading_otc']: 
    #權證
    file=openpyxl.load_workbook(DISK+day+'/權證-'+day+'.xlsx')
    sheet1=file['權證']
    cursor.execute("select security_code,security_name,fd_totalbuy,fd_totalsell,fd_difference,d_difference,dh_totalbuy,dh_totalsell,dh_difference,total_difference,trade_date from {} where length(security_code)=6 and trade_date='{}'".format(t,day))
    all_data=cursor.fetchall()
    for data in all_data:
        temp=list(data)
        sheet1.append(temp)
    file.save(DISK+day+'/權證-'+day+'.xlsx')