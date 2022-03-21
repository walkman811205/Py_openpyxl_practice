# -*- coding: utf-8 -*-
import openpyxl
import datetime
import pymysql
import dbconfig

DISK='D:/'
d=-1
day = (datetime.datetime.now()+datetime.timedelta(days=d)).strftime("%Y%m%d")
path = DISK+day

dbconnect = dbconfig.setting
conn = pymysql.connect(**dbconnect)
cursor = conn.cursor()

cursor.execute("select trade_date from trading_detail where security_code='2330'")
all_day=[]
for i in cursor.fetchall():
    all_day.append(*i)

def change(data):
    y=''
    for i in data:
        if i!=',':
            y=y+i
    return float(y)

def change_int(data):
    y=''
    for i in data:
        if i!=',':
            y=y+i
    return int(float(y))

def stock_name(code):
        cursor.execute("select security_name from trading_detail where security_code='{}'".format(code))
        name=cursor.fetchone()
        if name==None:
            cursor.execute("select security_name from trading_otc where security_code='{}'".format(code))
            name=cursor.fetchone()
            if name==None:
                return ''
        return name[0]

def increase(code,n):
    s=0
    for i in range(n):
        incre=0
        cursor.execute("select dir,change_ from daily_quotes where security_code='{}' and quotes_date='{}'".format(code, all_day[-1-i]))
        temp=cursor.fetchone()
        if temp==None:
            cursor.execute("select dir from daily_otc where security_code='{}' and quotes_date='{}'".format(code, all_day[-1-i]))
            temp=cursor.fetchone()
            if temp==None:
                return ''
            else:
#                    incre=float(temp[0])
                incre=(temp[0])
        else:
            if temp[0]=='+':
                incre=float(temp[1])
            elif temp[0]=='-':
                incre=-float(temp[1])
            else:
                incre=float(temp[1])
            s=s+incre
    return round(s/n,2)

def k(code):
        num=0
        cursor.execute("select opening_price,closing_price from daily_quotes where security_code='{}' and quotes_date='{}'".format(code, day))
        temp=cursor.fetchone()
        if temp!=None:
            t='daily_quotes'
        else:
            t='daily_otc'
        for d in all_day[::-1]:
            cursor.execute("select opening_price,closing_price from {} where security_code='{}' and quotes_date='{}'".format(t,code, d))
            temp=cursor.fetchone()
            if change(temp[0])<=change(temp[1]):
                num+=1
            else:
                if num==0:
                    num-=1
                    continue
                else:
                    break
        return num

file=openpyxl.load_workbook(path+'/標的物(購)-'+day+'.xlsx')
file1=openpyxl.load_workbook(path+'/標的物-Top20(購)-'+day+'.xlsx')

sheet=file['標的物(購)']
sheet20=file1['標的物-Top20(購)']

buy_data=sheet.values   
buy=[]
for i in buy_data:
    if list(i)[2]=='總計' and len(list(i)[-1])>3:
        buy.append(tuple(i))
buy_buy20=sorted(buy,key=lambda x:change_int(x[-5]),reverse=True)[:20]
buy_sell20=sorted(buy,key=lambda x:change_int(x[-4]),reverse=True)[:20]

final_data1=[['買賣超股數(購)-TOP20','','','(自營商買進股數)']]
for i,x in zip([buy_buy20,buy_sell20],range(2)):
    for j in i:
        d1=increase(j[-1], 1)
        d5=increase(j[-1], 5)
        k_day=k(j[-1])
        cursor.execute("select trade_volume from daily_quotes where security_code='{}' and quotes_date='{}'".format(j[-1], day))
        volume_sum=cursor.fetchone()
        if volume_sum==None:
            cursor.execute("select trade_volume from daily_otc where security_code='{}' and quotes_date='{}'".format(j[-1], day))
            volume_sum=cursor.fetchone()
        #print(j[-3],volume_sum[0])
        diff_rate=change(j[-3])/change(volume_sum[0])*100
        if x==0:
            final_data1.append([j[-1],stock_name(j[-1]),j[-5],j[-2],k_day,d1,d5,round(diff_rate,2)])
        elif x==1:
            final_data1.append([j[-1],stock_name(j[-1]),j[-4],j[-2],k_day,d1,d5,round(diff_rate,2)])
    final_data1.append(['買賣超股數(購)-TOP20','','','(自營商賣出股數)'])
final_data2=[['買賣超股數(售)-TOP20','','','(自營商買進股數)']]

for i in final_data1[:-1]:
    sheet20.append(i)

file1.save(path+'/標的物-Top20(購)-'+day+'.xlsx')  
cursor.close()
conn.close()