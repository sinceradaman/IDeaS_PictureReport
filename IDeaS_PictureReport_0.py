# -*- coding: utf-8 -*-
"""
Created on Thu Jul 20 21:41:13 2017

@author: Administrator
"""
# PictureReport.py
import pandas as pd
import xlwings as xw
import gc
#import datetime as dt
from sqlalchemy import create_engine
#Day及以下的函数
#from pandas.tseries.offsets import Day

# ===================Refresh Calculation=======================================
def Refresh_A_Glance():

    wb = xw.Book.caller()
#将Excel中的参数传入Python   
    sht = wb.sheets['A_Glance']
    Now_Date=sht.range('A1').value
#    Now_Date=pd.to_datetime(['2017-8-30'],infer_datetime_format=True)
    Hotel_Name=sht.range('E1').value
    Filter_Month=sht.range('N1').value    
#测试数据是否传入Python    
#    sht.range('A9').value=Now_Date
#    sht.range('A10').value = Hotel_Name
#    sht.range('A11').value  = Filter_Month

#以Filter_Month所选择的月份起，并以60天为一个Period(或者以当月为一个Period) 
    start_time=str(Now_Date.year)+'-'+'0'+str(int(Filter_Month))+'-'+'01'
    
#    end_time=pd.to_datetime(start_time)+dt.timedelta(60)
    end_time=pd.to_datetime(start_time)+pd.tseries.offsets.MonthEnd()
    end_time=end_time.strftime('%Y-%m-%d')
#通过传入的函数值截取数据，'BUSINESS_DATE'，'OCCUPANCY_DATE'，'HOTEL_NAME_ENG'   
#日期截取需要转化成DatatimeIndex会比较方便 
    global data_used
    
    data_used=data_used.set_index('OCCUPANCY_DATE')
    data_used=data_used[start_time : end_time]
    data_used=data_used[(data_used['BUSINESS_DATE'] == Now_Date)
                        &(data_used['HOTEL_NAME_ENG'] == Hotel_Name)]
#求出所需要字段的Occ%   
    data_used['Actual']=data_used['ROOMS_SOLD'].div(data_used['HOTEL_CAPACITY'])
    data_used['Transient']=data_used['ROOMS_SOLD_TRAN'].div(data_used['HOTEL_CAPACITY'])
    data_used['Group']=data_used['ROOMS_SOLD_GROUP'].div(data_used['HOTEL_CAPACITY'])
    data_used['Forecast']=data_used['OCC_FCST_TOTAL'].div(data_used['HOTEL_CAPACITY'])
#将星期作为字符串附在日期之后，所以首先求出weekday，并进行缩写转化 
#将日期截取需要转化成DatatimeIndex会比较方便重新转化成Column才可以做字符串转化
    data_used=data_used.reset_index()
#求出weekday并取出缩写（前三个字符）    
    data_used['weekday']=data_used['OCCUPANCY_DATE'].dt.weekday_name.str[0:3]
#转化成字符串并与‘weekday’粘贴在一起    
    data_used['OCCUPANCY_DATE']=data_used['OCCUPANCY_DATE'].dt.strftime('%Y-%m-%d')
    data_used['Datetime_weekday']=data_used['OCCUPANCY_DATE'].str.cat(' '+data_used['weekday'])
    data_used=data_used.set_index('Datetime_weekday')
#选择项目所需要的字段进行Slice，并且转置，清空之前表中内容并传入数据      
    data_used=data_used.loc[:,['Actual','Transient','Group','Forecast']].T
    sht.range('B3:BK7').clear_contents()
    sht.range('B3').options(index=False).value = data_used

    return    

# ===================Load_Database=============================================
def Load_Database():
#数据库类型+数据库驱动名称://用户名:口令@机器地址:端口号/数据库名
    engine = create_engine('mssql+pymssql://sa:Ppguiandzp543@4BQGENW6GCHWI6V/mytest')
    with engine.connect() as conn, conn.begin():
        data = pd.read_sql('IDeaS_Hotel', conn)
        HotelList= pd.read_sql('Hotel_List', conn) 

#这句尼玛至关重要，一定要让KEY一致,所以要用strip()去除空格    
    data['HOTEL_CODE']=data['HOTEL_CODE'].str.strip()
#映射HotelList中的数据，等同于vlookup
    data_merge=pd.merge(data
                    ,HotelList.loc[:,['HOTEL_CODE','HOTEL_NAME_ENG','OWNER']]
                    ,on='HOTEL_CODE'
                    ,how='left')
#截取PictureReport所需要的数据，后期放到select语句中，减少内存 
    global data_used
    data_used=data_merge.loc[:,['HOTEL_CODE'
                                ,'HOTEL_NAME_ENG'
                                ,'BUSINESS_DATE'
                                ,'OCCUPANCY_DATE'
                                ,'HOTEL_CAPACITY'
                                ,'ROOMS_SOLD'
                                ,'ROOMS_SOLD_TRAN'
                                ,'ROOMS_SOLD_GROUP'
                                ,'OCC_FCST_TOTAL']]
 
    return data_used
# ===================Reset_Database=============================================
def Reset_Database():
    
    global data_used
    del data_used
    gc.collect()
    
    return
   
#Debug调试程序所需函数    
if __name__ == '__main__':
    xw.Book('IDeaS_PictureReport.xlsm').set_mock_caller()
    Load_Database()
    Refresh_A_Glance()
    Reset_Database()    




