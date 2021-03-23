# -*- coding: utf-8 -*-
"""
Created on Tue May  1 22:58:58 2018

@author: Max
"""

import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
import re
import time
from io import StringIO

def getMITable(year, month): #月營收
    print('start '+str(year)+'  '+str(month))
    if year > 1990:
        year -= 1911
    
    url = 'http://mops.twse.com.tw/nas/t21/sii/t21sc03_'+str(year)+'_'+str(month)+'_0.html'
    if year <= 98:
        url = 'http://mops.twse.com.tw/nas/t21/sii/t21sc03_'+str(year)+'_'+str(month)+'.html'
    
   
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
    
    # 下載該年月的網站，並用pandas轉換成 dataframe
    r = requests.get(url, headers)
    r.encoding = 'big5'
    html_df = pd.read_html(StringIO(r.text))
    
    # 處理資料
    if html_df[0].shape[0] > 500:
        df = html_df[0].copy()
    else:
        df = pd.concat([df for df in html_df if df.shape[1] <= 11])
    df = df[list(range(0,10))]
    column_index = df.index[(df[0] == '公司代號')][0]
    df.columns = df.iloc[column_index] 
    df['當月營收'] = pd.to_numeric(df['當月營收'], 'coerce')
    df['去年同月增減(%)'] = pd.to_numeric(df['去年同月增減(%)'], 'coerce')
    df = df[~df['當月營收'].isnull()]
    df = df[df['公司代號'] != '合計']
    ##df=df.reset_index()
    ##df = df.drop(['index'], axis=1)
    df.index=df['公司代號']
    time.sleep(1)
    df=df.loc[:,['去年同月增減(%)']]
    
    return df

#newestYear='106'#年報已出最新年度
#newestSeason='106 4Q'#季報已出最新一季
sy=2018
ss=3  #尚未公佈的季度
def season():
    global sy
    global ss
    ss=ss-1
    if ss==0:
        ss=ss+4
        sy=sy-1
    return str(sy)+'_'+str(ss)+'Q'
seasonColumns=[season(),season(),season(),season(),season(),season(),season(),season()]#季報欄位
y=2018 
m=10  #尚未公佈營收的月份
def month():  
    return m

def year():
    global m
    global y
    m=m-1
    if m==0:
        m=m+12
        y=y-1
    return y


df=getMITable(year(),month())


df2=getMITable(year(),month())
df3=getMITable(year(),month())
df4=getMITable(year(),month())
df5=getMITable(year(),month())
df6=getMITable(year(),month())
df7=getMITable(year(),month())
df8=getMITable(year(),month())
df9=getMITable(year(),month())
df10=getMITable(year(),month())
df11=getMITable(year(),month())
df12=getMITable(year(),month())
df13=getMITable(year(),month())
df14=getMITable(year(),month())
month=pd.concat([df,df2,df3,df4,df5,df6,df7,df8,df9,df10,df11,df12,df13,df14],axis=1,join_axes=[df.index])
month.to_csv('月營收表.csv',encoding='utf8') #輸出月營收表


def getISYTable(ids): #損益年報===============================================
    print('start '+str(ids))
    url='http://jsjustweb.jihsun.com.tw/z/zc/zcq/zcqa/zcqa_'+str(ids)+'.djhtm'
    headers = {'user-agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36',
               'Accept-Language':'en-US,en;q=0.9,zh-TW;q=0.8,zh;q=0.7',
              'Accept-Encoding':'gzip, deflate',
              'Cache-Control':'max-age=0',
              'Connection':'keep-alive',
              }
    r = requests.get(url, headers=headers)
    soup=BeautifulSoup(r.text)
    
    yearTable=[]
    yearTable.append(str(ids))
    try:
        if(soup.select('.t01 tr')[4].select('td')[0].text=='營業收入淨額'and
          soup.select('.t01 tr')[99].select('td')[0].text=='每股盈餘'and
          soup.select('.t01 tr')[100].select('td')[0].text=='加權平均股數'):  #判斷表格是否正確
            if(soup.select('.t01 tr')[1].select('td')[1].text=='2017'):#最新2017年
                try:
                    yearTable.append(int(re.sub(r',',"",soup.select('.t01 tr')[4].select('td')[1].text)))
                except:
                    yearTable.append(None)                           
                try:
                    yearTable.append(float(re.sub(r',',"",soup.select('.t01 tr')[99].select('td')[1].text)))
                except:
                    yearTable.append(None)              
                try:
                    yearTable.append(int(re.sub(r',',"",soup.select('.t01 tr')[100].select('td')[1].text)))
                except:
                    yearTable.append(None)                     
            else:
                print("損益年報-無最新一季:"+str(ids))
                for i in range(1,4):
                    yearTable.append(None)
        else:#表格錯誤
            print('損益年報-表格錯誤:'+str(ids))
            for i in range(1,4):
                yearTable.append(None)
    except:#表格遺失
        print('損益年報-表格遺失:'+str(ids))
        for i in range(1,4):
            yearTable.append(None)
    
    df=pd.DataFrame(yearTable).T
    df.columns=['公司代號','營業收入淨額','每股盈餘','加權平均股數']
    
    
    #抓歷年EPS
    global oldEps
    oldEps=[]
    data=[]
    for i in range(1,5):
        try:
            data.append(float(re.sub(r',',"",soup.select('.t01 tr')[99].select('td')[i].text)))
        except:
            print('oldEps error')
            data.append(None)
    oldEps.append([str(ids),str(ids),str(ids),str(ids)])
    oldEps.append(['2017','2016','2015','2014'])
    oldEps.append(data)
    
    return df


def getISTable(ids): #損益表==========================================================
    print('start '+ids)
    url='http://jsjustweb.jihsun.com.tw/z/zc/zcq/zcq_'+str(ids)+'.djhtm'
    headers = {'user-agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36',
               'Accept-Language':'en-US,en;q=0.9,zh-TW;q=0.8,zh;q=0.7',
              'Accept-Encoding':'gzip, deflate',
              'Cache-Control':'max-age=0',
              'Connection':'keep-alive',
              }
    r = requests.get(url, headers=headers)
    soup=BeautifulSoup(r.text)

    netIncome=[] #營業收入淨額
    income=[]#營業利益
    outIncome=[]#營業外收入及支出
    afterTax=[]#常續性稅後淨利
    eps=[] #每股盈餘
    try:
        if(soup.select('.t01 tr')[3].select('td')[0].text=='\u3000營業收入毛額'and
          soup.select('.t01 tr')[17].select('td')[0].text=='營業利益'and
          soup.select('.t01 tr')[64].select('td')[0].text=='營業外收入及支出'and
          soup.select('.t01 tr')[102].select('td')[0].text=='常續性稅後淨利'and
          soup.select('.t01 tr')[99].select('td')[0].text=='每股盈餘'):#判斷表格是否正確
            if(soup.select('.t01 tr')[1].select('td')[1].text=='2018.2Q'):#如果是最新一季
                 for i in range(1,9):
                    try:
                        netIncome.append(int(re.sub(r',',"",soup.select('.t01 tr')[3].select('td')[i].text)))
                    except :
                        netIncome.append(None)
                    try:    
                        income.append(int(re.sub(r',',"",soup.select('.t01 tr')[17].select('td')[i].text)))
                    except :
                        income.append(None)
                    try:    
                        outIncome.append(int(re.sub(r',',"",soup.select('.t01 tr')[64].select('td')[i].text)))
                    except :
                        outIncome.append(None)
                    try:   
                        afterTax.append(int(re.sub(r',',"",soup.select('.t01 tr')[102].select('td')[i].text)))
                    except : 
                        afterTax.append(None)
                    try:    
                        eps.append(float(soup.select('.t01 tr')[99].select('td')[i].text))
                    except :
                        eps.append(None)     
            else:        #不是最新一季
                print("損益表-無最新一季:"+str(ids))
                netIncome.append(None)
                income.append(None)
                outIncome.append(None)
                afterTax.append(None)
                eps.append(None)
                for i in range(1,8):
                    try:
                        netIncome.append(int(re.sub(r',',"",soup.select('.t01 tr')[3].select('td')[i].text)))
                    except :
                        netIncome.append(None)
                    try:    
                        income.append(int(re.sub(r',',"",soup.select('.t01 tr')[17].select('td')[i].text)))
                    except :
                        income.append(None)
                    try:    
                        outIncome.append(int(re.sub(r',',"",soup.select('.t01 tr')[64].select('td')[i].text)))
                    except :
                        outIncome.append(None)
                    try:   
                        afterTax.append(int(re.sub(r',',"",soup.select('.t01 tr')[102].select('td')[i].text)))
                    except : 
                        afterTax.append(None)
                    try:    
                        eps.append(float(soup.select('.t01 tr')[99].select('td')[i].text))
                    except :
                        eps.append(None)
        else:   #表格錯誤處理
            print("損益表-表格錯誤:"+str(ids))
            for i in range(1,9):
                netIncome.append(None)
                income.append(None)
                outIncome.append(None)
                afterTax.append(None)
                eps.append(None)
    except: #表格遺失
        print("損益表-表格遺失:"+str(ids))
        for i in range(1,9):
            netIncome.append(None)
            income.append(None)
            outIncome.append(None)
            afterTax.append(None)
            eps.append(None)

    df2=pd.DataFrame(netIncome)
    df3=pd.DataFrame(income)
    df4=pd.DataFrame(outIncome)
    df5=pd.DataFrame(afterTax)
    df6=pd.DataFrame(eps)
    df2=pd.concat([df2,df3,df4,df5,df6],axis=1)
    df2.columns=['營業收入毛額','營業利益','營業外收入及支出','常續性稅後淨利','每股盈餘']
    df2=df2.transpose()
#    time.sleep(0.5)
    return df2


def getCFTable(ids): #現金流量表==========================================================
    print('start '+ids)
    url='http://jsjustweb.jihsun.com.tw/z/zc/zc3/zc3_'+str(ids)+'.djhtm'
    headers = {'user-agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36',
               'Accept-Language':'en-US,en;q=0.9,zh-TW;q=0.8,zh;q=0.7',
              'Accept-Encoding':'gzip, deflate',
              'Cache-Control':'max-age=0',
              'Connection':'keep-alive',
              }
    r = requests.get(url, headers=headers)
    soup=BeautifulSoup(r.text)

    useCash=[]#來自營運之現金流量
    investCash=[]#投資活動之現金流量
    eventCash=[]#籌資活動之現金流量
    cash=[]#本期產生現金流量
    
    try:
        if(soup.select('.t01 tr')[55].select('td')[0].text==' 來自營運之現金流量'and
          soup.select('.t01 tr')[71].select('td')[0].text==' 投資活動之現金流量'and
          soup.select('.t01 tr')[93].select('td')[0].text==' 籌資活動之現金流量'and
          soup.select('.t01 tr')[95].select('td')[0].text==' 本期產生現金流量'):#判斷表格是否正確
            if(soup.select('.t01 tr')[1].select('td')[1].text=='2018.2Q'):#如果是最新一季
                 for i in range(1,9):
                    try:
                        useCash.append(int(re.sub(r',',"",soup.select('.t01 tr')[55].select('td')[i].text)))
                    except :
                        useCash.append(None)
                    try:    
                        investCash.append(int(re.sub(r',',"",soup.select('.t01 tr')[71].select('td')[i].text)))
                    except :
                        investCash.append(None)
                    try:    
                        eventCash.append(int(re.sub(r',',"",soup.select('.t01 tr')[93].select('td')[i].text)))
                    except :
                        eventCash.append(None)
                    try:   
                        cash.append(int(re.sub(r',',"",soup.select('.t01 tr')[95].select('td')[i].text)))
                    except : 
                        cash.append(None)
            else:
                print("現金流量表-無最新一季:"+str(ids))
                useCash.append(None)
                investCash.append(None)
                eventCash.append(None)
                cash.append(None)
                for i in range(1,8):
                    try:
                        useCash.append(int(re.sub(r',',"",soup.select('.t01 tr')[55].select('td')[i].text)))
                    except :
                        useCash.append(None)
                    try:    
                        investCash.append(int(re.sub(r',',"",soup.select('.t01 tr')[71].select('td')[i].text)))
                    except :
                        investCash.append(None)
                    try:    
                        eventCash.append(int(re.sub(r',',"",soup.select('.t01 tr')[93].select('td')[i].text)))
                    except :
                        eventCash.append(None)
                    try:   
                        cash.append(int(re.sub(r',',"",soup.select('.t01 tr')[95].select('td')[i].text)))
                    except : 
                        cash.append(None)   
        else:#表格錯誤處理
            print('現金流量表-表格錯誤:'+str(ids))
            for i in range(1,9):
                useCash.append(None)
                investCash.append(None)
                eventCash.append(None)
                cash.append(None)
    except:
        print('現金流量表-表格遺失'+str(ids))
        for i in range(1,9):#如果連表都沒有==
            useCash.append(None)
            investCash.append(None)
            eventCash.append(None)
            cash.append(None)

    df2=pd.DataFrame(useCash)
    df3=pd.DataFrame(investCash)
    df4=pd.DataFrame(eventCash)
    df5=pd.DataFrame(cash)
    df2=pd.concat([df2,df3,df4,df5],axis=1)
    df2.columns=['來自營運之現金流量','投資活動之現金流量','籌資活動之現金流量','本期產生現金流量']
    df2=df2.transpose()
    return df2


def getPrice(ids):#收盤價==============================================================
    print('start '+str(ids))
    url='https://goodinfo.tw/StockInfo/StockBzPerformance.asp?STOCK_ID='+str(ids)+'&SHEET=PER/PBR'
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
      
    r = requests.get(url, headers=headers)
    r.encoding = 'unicode'
    soup=BeautifulSoup(r.text)
    
    global closePrice #個別處理收盤價
    closePrice.append([str(ids),soup.select('#row0 td')[5].text])
    
    data=[]#公司代號  , 年度  ,  最高  , 最低
    for i in range(0,5):
        try:
            data.append([str(ids),soup.select('#row'+str(i)+' td')[0].text,soup.select('#row'+str(i)+' td')[3].text,soup.select('#row'+str(i)+' td')[4].text])
        except:
            print('查無資料')
            data.append([str(ids),None,None,None])
    df2=pd.DataFrame(data)
    return df2
    

def getPriceTable(date):#收盤價==============================================================
    
    url='http://www.tse.com.tw/exchangeReport/MI_INDEX?response=json&date='+str(date)+'&type=ALLBUT0999&_=1525186585429'
    headers = {'user-agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/52.0.2743.116 Safari/537.36',
               'Accept-Language':'en-US,en;q=0.9,zh-TW;q=0.8,zh;q=0.7',
              'Accept-Encoding':'gzip, deflate',
              'Cache-Control':'max-age=0',
              'Connection':'keep-alive',
              }
    r = requests.get(url, headers=headers)
    js=r.json()
    df=pd.DataFrame.from_dict(js['data5'])
    df.columns=['id','名稱','成交股數','成交股數','成交金額','開盤價','最高價','最低價','收盤價','漲跌','漲跌價差','最後揭示買價','最後揭示買量','最後揭示賣價','最後揭示賣量','本益比']   
    return df
#----------------------------------------收盤價-----------------------------------------#
today=time.strftime("%Y%m%d", time.localtime())#yyyymmdd
price_df=getPriceTable(today).iloc[:,[0,8]] #收盤日期
IDdf.columns=['shit','id']
IDdf.index=IDdf['id']
IDdf['id']=IDdf['id'].astype(str)#將ID轉STR，才能合併
price_df['id']=price_df['id'].astype(str)

result=pd.merge(IDdf,price_df,how='left',on=['id'])#合併
result.to_csv('股價.csv',encoding='utf_8_sig')

#----------------------------------------損益年報----------------------------------------#
oldEps=[]#歷年EPS
idList=df.index.tolist()#所有公司代號
first=idList.pop(0) #取出第一個創建dataframe
ISY=getISYTable(first)
OEdf=pd.DataFrame(oldEps).T
for i in idList:
    temp_ISY=getISYTable(i)
    ISY=pd.concat([ISY,temp_ISY],axis=0)
    temp_oldEps_df=pd.DataFrame(oldEps).T
    OEdf=pd.concat([OEdf,temp_oldEps_df],axis=0)
ISY.to_csv('損益年報.csv',encoding='utf8',index=False)
OEdf.columns=['公司代號','年度','EPS']
OEdf.to_csv('歷年EPS.csv',encoding='utf8')


#----------------------------------------損益季表----------------------------------------#
idList=df.index.tolist()#所有公司代號
first=idList.pop(0) #取出第一個創建dataframe
IS=getISTable(first)
IS=pd.concat([IS],keys=[first])
for i in idList:
    temp_IS=getISTable(i)
    temp_IS=pd.concat([temp_IS],keys=[i])
    IS=pd.concat([IS,temp_IS],axis=0)
IS.columns=seasonColumns
IS.to_csv('損益季表.csv',encoding='utf8')

#----------------------------------------現金流量----------------------------------------#
idList=df.index.tolist()#所有公司代號
first=idList.pop(0) #取出第一個創建dataframe
CF=getCFTable(first)
CF=pd.concat([CF],keys=[first])
for i in idList:
    temp_CF=getCFTable(i)
    temp_CF=pd.concat([temp_CF],keys=[i])
    CF=pd.concat([CF,temp_CF],axis=0)
CF.columns=seasonColumns
CF.to_csv('現金流量.csv',encoding='utf8')


#------------------------------------------IP更換-----------------------------------------

import os
g_adsl_account = {"name": "寬頻連線 2",
                "username": "73318395@hinet.net",
                "password": "pauli246"}  
class Adsl(object):
    # =============================
    # __init__ : name: adsl名
    def __init__(self):
        self.name = g_adsl_account["name"]
        self.username = g_adsl_account["username"]
        self.password = g_adsl_account["password"]
    # =============================
    # set_adsl : 修改adsl設置
    def set_adsl(self, account):
        self.name = account["name"]
        self.username = account["username"]
        self.password = account["password"]     
    # =============================
    # connect : 寬頻連線
    def connect(self):
        cmd_str = 'rasdial "寬頻連線 2" 73318395@hinet.net pauli246'
        os.system(cmd_str)
        time.sleep(5)
    # =============================
    # disconnect : 斷線
    def disconnect(self):
        cmd_str ="rasdial /DISCONNECT"
        os.system(cmd_str)
        time.sleep(5)
    #=============================
    # reconnect : 重撥
    def reconnect(self):
        self.disconnect()
        self.connect()
        
#----------------------------------------收盤價-----------------------------------------#
closePrice=[]
adsl=Adsl()
idList=df.index.tolist()#所有公司代號
first=idList.pop(0) #取出第一個創建dataframe
price=getPrice(first)
count=0
for i in idList:
    count=count+1
    try:
        temp_price=getPrice(i)
    except:
        print('error reconnect!')
        adsl.reconnect()
        time.sleep(5)
        temp_price=getPrice(i)
    price=pd.concat([price,temp_price],axis=0)
price.columns=['公司代號','年度','最高價','最低價']
price.to_csv('最高最低價.csv',encoding='utf8')
CPdf=pd.DataFrame(closePrice)
CPdf.to_csv('收盤價.csv',encoding='utf8')

#---------------------------------------idList備份--------------------------------------
idList=df.index.tolist()
IDdf=pd.DataFrame(idList)
IDdf.to_csv('ids.csv',encoding='utf8')

IDdf=pd.read_csv('ids.csv')#恢復idList
idList=IDdf.iloc[:,1].tolist()

#---------------------------------------新增單月營收-------------------------------------


IDdf.columns=['shit','id']
IDdf.index=IDdf['id']
IDdf['id']=IDdf['id'].astype(str)
df['id'] = df.index
newMonth=pd.merge(IDdf,df,how='left',on=['id'])
newMonth.to_csv('新單月營收.csv',encoding='utf8')

#TEST ZONE===================================================


url='https://goodinfo.tw/StockInfo/StockBzPerformance.asp?STOCK_ID=2330&SHEET=PER/PBR'
headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}

r = requests.get(url, headers=headers)
r.encoding = 'unicode'
soup=BeautifulSoup(r.text)

data=[]
df2=pd.DataFrame(data)
for i in range(0,5):
    data.append([soup.select('#row'+str(i)+' td')[0].text,soup.select('#row'+str(i)+' td')[3].text],soup.select('#row'+str(i)+' td')[4].text)


def printAll():    
    count=2
    for tr in soup.select('.t01 tr')[3:]:
        td=tr.select('td')
        count=count+1
        print(count,td[0].text,td[1].text,td[2].text,td[3].text,td[4].text,td[5].text,td[6].text,td[7].text,td[8].text)
        #print(count,td[0])
printAll()

