# -*- coding: utf-8 -*-
"""
Created on Thu Jul 20 21:41:13 2017

@author: Administrator
"""
import numpy as np
import pandas as pd
import xlwings as xw
import datetime as dt
import shelve
from sqlalchemy import create_engine
# ===================Load_Database=============================================
def Load_Database():

    wb = xw.Book.caller()     
#将今天的日期传入A_Glance  
    Now_Date=dt.datetime.today().strftime('%Y-%m-%d')
    Now_Date_LY=str(int(dt.datetime.today().year)-1)
    sht = wb.sheets['A_Glance']
    sht.range('A1').clear_contents()
    sht.range('A1').value = Now_Date
#建立SQL的SELECT的内容      
    IDeaS_Hotel_Data=("SELECT HOTEL_CODE"
                      ",AO"
                      ",BUSINESS_DATE"
                      ",OCCUPANCY_DATE"
                      ",HOTEL_CAPACITY"
                      ",ROOMS_SOLD"
                      ",ROOMS_SOLD_TRAN"
                      ",ROOMS_SOLD_GROUP"
                      ",ROOM_REVENUE"
                      ",OCC_FCST_TOTAL"
                      ",OCC_FCST_TRAN"
                      ",OCC_FCST_GROUP"
                      ",LRV"
                      ",SYS_UNCON_DEMAND_TOTAL"
                      ",SE_IMPACT_Y"
                      ",TRAN_ADR"
                      ",GROUP_ADR"
                      " FROM IDeaS_Hotel" 
                      " WHERE (AO='A' AND BUSINESS_DATE>=('"+ Now_Date_LY +"')) OR (BUSINESS_DATE=('"+ Now_Date +"') AND AO='O')")
    Property_Data=("SELECT code"
                      ",Eng_Name" 
                      " FROM Property")   
#数据库类型+数据库驱动名称://用户名:口令@机器地址:端口号/数据库名
    engine = create_engine('mssql+pymssql://sa:Abcd12345@10.1.93.200/BF_1')
    with engine.connect() as conn, conn.begin():
        data = pd.read_sql(IDeaS_Hotel_Data, conn)
    
    engine = create_engine('mssql+pymssql://sa:Abcd12345@10.1.93.200/wanda')
    with engine.connect() as conn, conn.begin():
        Property= pd.read_sql(Property_Data, conn) 
#这句尼玛至关重要，一定要让KEY一致,所以要用strip()去除空格    
    data['HOTEL_CODE']=data['HOTEL_CODE'].str.strip()
#映射HotelList中的数据，等同于vlookup
    data_merge=pd.merge(data
                    ,Property
                    ,left_on='HOTEL_CODE'
                    ,right_on='code'
                    ,how='left')
    data_used=data_merge.drop('code', axis=1)
    
    s = shelve.open('PictureReport_Data.db', writeback=True) 
    s['key'] = data_used
    s.close
    return
#===================Refresh_A_Glance=======================================
def Refresh_A_Glance():
    wb = xw.Book.caller()
#将Excel中的参数传入Python   
    sht = wb.sheets['A_Glance']
    Now_Date_Year=sht.range('H1').value    
    Hotel_Name=sht.range('E1').value
    Filter_Month=sht.range('I1').value    
#以Filter_Month所选择的月份起，并以60天为一个Period(或者以当月为一个Period) 
    start_time=str(int(Now_Date_Year))+'-'+str(int(Filter_Month))+'-'+'01'
    end_time=pd.to_datetime(start_time)+pd.tseries.offsets.MonthEnd()
    end_time=end_time.strftime('%Y-%m-%d')
#从Cache中取数据
    s = shelve.open('PictureReport_Data.db') 
    data_used = s['key'] 
    s.close
#序列化后的数据date数据会转化为Object数据,所以需要重新转化为date数据
    data_used['OCCUPANCY_DATE']=pd.to_datetime(data_used['OCCUPANCY_DATE'])
#通过传入的函数值截取数据，'BUSINESS_DATE'，'OCCUPANCY_DATE'，'HOTEL_NAME_ENG'   
#日期截取需要转化成DatatimeIndex会比较方便      
    data_used=data_used.set_index('OCCUPANCY_DATE')   
    data_used=data_used[data_used['Eng_Name'] == Hotel_Name]
    data_used=data_used[start_time : end_time] 
#求出所需要字段的Occ%   
    data_used['Actual']=data_used['ROOMS_SOLD'].div(data_used['HOTEL_CAPACITY'])
    data_used['Transient']=data_used['ROOMS_SOLD_TRAN'].div(data_used['HOTEL_CAPACITY'])
    data_used['Group']=data_used['ROOMS_SOLD_GROUP'].div(data_used['HOTEL_CAPACITY'])
    data_used['Forecast']=data_used['OCC_FCST_TOTAL'].div(data_used['HOTEL_CAPACITY'])
#将星期作为字符串附在日期之后，所以首先求出weekday，并进行缩写转化 
#将日期截取需要转化成DatatimeIndex会比较方便重新转化成Column才可以做字符串转化
    data_used=data_used.reset_index()
    data_used.sort_values(by='OCCUPANCY_DATE',inplace=True)
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
#===================Refresh_Date_Report=======================================
def Date_Report():
    wb = xw.Book.caller()  
    sht = wb.sheets['Date_Report']
#将Excel中的参数传入Python     
    Now_Date_Year=sht.range('H1').value  
    Hotel_Name=sht.range('E1').value
    Filter_Month=sht.range('I1').value 
#以Filter_Month所选择的月份计算同比月区间    
    start_time_TY=str(int(Now_Date_Year))+'-'+str(int(Filter_Month))+'-'+'01'
    end_time_TY=pd.to_datetime(start_time_TY)+pd.tseries.offsets.MonthEnd()
    end_time_TY=end_time_TY.strftime('%Y-%m-%d')    
    
    start_time_LY=str(int(Now_Date_Year)-1)+'-'+str(int(Filter_Month))+'-'+'01'
    end_time_LY=pd.to_datetime(start_time_LY)+pd.tseries.offsets.MonthEnd()
    end_time_LY=end_time_LY.strftime('%Y-%m-%d') 
#从Cache中取数据
    s = shelve.open('PictureReport_Data.db') 
    data_used = s['key'] 
    s.close
#序列化后的数据date数据会转化为Object数据,所以需要重新转化为date数据
    data_used['OCCUPANCY_DATE']=pd.to_datetime(data_used['OCCUPANCY_DATE'])
#通过传入的函数值截取数据，'BUSINESS_DATE'，'OCCUPANCY_DATE'，'HOTEL_NAME_ENG'   
#日期截取需要转化成DatatimeIndex会比较方便      
    data_used=data_used.set_index('OCCUPANCY_DATE')   
    data_used=data_used[data_used['Eng_Name'] == Hotel_Name]
##将无序的时序进行排序inplace=ture很重要
#    data_used.sort_index(inplace=True)
#    data_used_TY=data_used[start_time_TY : end_time_TY]
#    data_used_TY['YEAR']='TY'
#    data_used_LY=data_used[start_time_LY : end_time_LY]
#    data_used_LY['YEAR']='LY'
#由于部分酒店缺失数据（Occupancy Date缺失），造成错误故采用重塑Index的方法    
    TY_DATE_Index=pd.DatetimeIndex(start=start_time_TY, end=end_time_TY, freq='D')
    data_used_TY=data_used.reindex(TY_DATE_Index, fill_value=0)
    data_used_TY['YEAR']='TY'    
    LY_DATE_Index=pd.DatetimeIndex(start=start_time_LY, end=end_time_LY, freq='D')
    data_used_LY=data_used.reindex(LY_DATE_Index, fill_value=0)
    data_used_LY['YEAR']='LY'
#合并TY和LY的数据，目的是为了成为DataFrame整体好切分，之后再分开  
    data_used=data_used_TY.append(data_used_LY)
#计算RoomSold（已有），ADR和RevPAR    
    data_used['ADR']=data_used['ROOM_REVENUE'].div(data_used['ROOMS_SOLD'])
    data_used['RevPAR']=data_used['ROOM_REVENUE'].div(data_used['HOTEL_CAPACITY'])
#为datatime添加weekday
    data_used=data_used.reset_index()
    data_used.rename(columns={'index':'OCCUPANCY_DATE'}, inplace=True)
    data_used['weekday']=data_used['OCCUPANCY_DATE'].dt.weekday_name.str[0:3]  
    data_used['OCCUPANCY_DATE']=data_used['OCCUPANCY_DATE'].dt.strftime('%Y-%m-%d')
    data_used['Datetime_weekday']=data_used['OCCUPANCY_DATE'].str.cat(' '+data_used['weekday'])
    data_used=data_used.set_index('Datetime_weekday')
#拆分TY和LY的同比数据   
    data_used_TY=data_used[data_used['YEAR']=='TY'].loc[:,['ROOMS_SOLD','ADR','RevPAR']].T
    data_used_LY=data_used[data_used['YEAR']=='LY'].loc[:,['ROOMS_SOLD','ADR','RevPAR']].T

    sht.range('B3:AF6').clear_contents()
    sht.range('B3').options(index=False).value = data_used_TY
    
    sht.range('B8:AF11').clear_contents()
    sht.range('B8').options(index=False).value = data_used_LY

    return
#===================Refresh_Date_Report=======================================
def Business_Type_Report():
    wb = xw.Book.caller()  
    sht = wb.sheets['Business_Type_Report']
#将Excel中的参数传入Python     
    Now_Date_Year=sht.range('H1').value  
    Hotel_Name=sht.range('E1').value
    Filter_Month=sht.range('I1').value 
#以Filter_Month所选择的月份计算同比月区间    
    start_time_TY=str(int(Now_Date_Year))+'-'+str(int(Filter_Month))+'-'+'01'
    end_time_TY=pd.to_datetime(start_time_TY)+pd.tseries.offsets.MonthEnd()
    end_time_TY=end_time_TY.strftime('%Y-%m-%d')    
    
    start_time_LY=str(int(Now_Date_Year)-1)+'-'+str(int(Filter_Month))+'-'+'01'
    end_time_LY=pd.to_datetime(start_time_LY)+pd.tseries.offsets.MonthEnd()
    end_time_LY=end_time_LY.strftime('%Y-%m-%d') 
#从Cache中取数据
    s = shelve.open('PictureReport_Data.db') 
    data_used = s['key'] 
    s.close
#序列化后的数据date数据会转化为Object数据,所以需要重新转化为date数据
    data_used['OCCUPANCY_DATE']=pd.to_datetime(data_used['OCCUPANCY_DATE'])  
#日期截取需要转化成DatatimeIndex会比较方便      
    data_used=data_used.set_index('OCCUPANCY_DATE')   
    data_used=data_used[data_used['Eng_Name'] == Hotel_Name]
    TY_DATE_Index=pd.DatetimeIndex(start=start_time_TY, end=end_time_TY, freq='D')
    data_used_TY=data_used.reindex(TY_DATE_Index, fill_value=0)
    data_used_TY['YEAR']='TY'    
    LY_DATE_Index=pd.DatetimeIndex(start=start_time_LY, end=end_time_LY, freq='D')
    data_used_LY=data_used.reindex(LY_DATE_Index, fill_value=0)
    data_used_LY['YEAR']='LY'  
    data_used=data_used_TY.append(data_used_LY)
#为datatime添加weekday
    data_used=data_used.reset_index()
    data_used.rename(columns={'index':'OCCUPANCY_DATE'}, inplace=True)
    data_used['weekday']=data_used['OCCUPANCY_DATE'].dt.weekday_name.str[0:3]  
    data_used['OCCUPANCY_DATE']=data_used['OCCUPANCY_DATE'].dt.strftime('%Y-%m-%d')
    data_used['Datetime_weekday']=data_used['OCCUPANCY_DATE'].str.cat(' '+data_used['weekday'])
    data_used=data_used.set_index('Datetime_weekday')    
#拆分TY和LY的同比数据   
    data_used_TY=data_used[data_used['YEAR']=='TY'].loc[:,['HOTEL_CAPACITY','ROOMS_SOLD_TRAN','ROOMS_SOLD_GROUP','OCC_FCST_TRAN','OCC_FCST_GROUP']].T
    data_used_LY=data_used[data_used['YEAR']=='LY'].loc[:,['HOTEL_CAPACITY','ROOMS_SOLD_TRAN','ROOMS_SOLD_GROUP','OCC_FCST_TRAN','OCC_FCST_GROUP']].T    
    
    sht.range('B3:AF8').clear_contents()
    sht.range('B3').options(index=False).value = data_used_TY
    
    sht.range('B10:AF15').clear_contents()
    sht.range('B10').options(index=False).value = data_used_LY

    return
#===================Refresh_Forecast_Validation=======================================
def Forecast_Validation():
    wb = xw.Book.caller()  
    sht = wb.sheets['Forecast_Validation']
#将Excel中的参数传入Python     
    Now_Date_Year=sht.range('H1').value  
    Hotel_Name=sht.range('E1').value
    Filter_Month=sht.range('I1').value 
#以Filter_Month所选择的月份计算同比月区间    
    start_time_TY=str(int(Now_Date_Year))+'-'+str(int(Filter_Month))+'-'+'01'
    end_time_TY=pd.to_datetime(start_time_TY)+pd.tseries.offsets.MonthEnd()
    end_time_TY=end_time_TY.strftime('%Y-%m-%d')    
    
    start_time_LY=str(int(Now_Date_Year)-1)+'-'+str(int(Filter_Month))+'-'+'01'
    end_time_LY=pd.to_datetime(start_time_LY)+pd.tseries.offsets.MonthEnd()
    end_time_LY=end_time_LY.strftime('%Y-%m-%d') 
#从Cache中取数据
    s = shelve.open('PictureReport_Data.db') 
    data_used = s['key'] 
    s.close
#序列化后的数据date数据会转化为Object数据,所以需要重新转化为date数据
    data_used['OCCUPANCY_DATE']=pd.to_datetime(data_used['OCCUPANCY_DATE'])   
#日期截取需要转化成DatatimeIndex会比较方便      
    data_used=data_used.set_index('OCCUPANCY_DATE')   
    data_used=data_used[data_used['Eng_Name'] == Hotel_Name]
    TY_DATE_Index=pd.DatetimeIndex(start=start_time_TY, end=end_time_TY, freq='D')
    data_used_TY=data_used.reindex(TY_DATE_Index, fill_value=0)
    data_used_TY['YEAR']='TY'    
    LY_DATE_Index=pd.DatetimeIndex(start=start_time_LY, end=end_time_LY, freq='D')
    data_used_LY=data_used.reindex(LY_DATE_Index, fill_value=0)
    data_used_LY['YEAR']='LY'  
    data_used=data_used_TY.append(data_used_LY)
#为datatime添加weekday
    data_used=data_used.reset_index()
    data_used.rename(columns={'index':'OCCUPANCY_DATE'}, inplace=True)
    data_used['weekday']=data_used['OCCUPANCY_DATE'].dt.weekday_name.str[0:3]  
    data_used['OCCUPANCY_DATE']=data_used['OCCUPANCY_DATE'].dt.strftime('%Y-%m-%d')
    data_used['Datetime_weekday']=data_used['OCCUPANCY_DATE'].str.cat(' '+data_used['weekday'])
    data_used=data_used.set_index('Datetime_weekday')    
#计算ADR
    data_used['ADR']=data_used['ROOM_REVENUE'].div(data_used['ROOMS_SOLD'])
#拆分TY和LY的同比数据   
    data_used_TY=data_used[data_used['YEAR']=='TY'].loc[:,['HOTEL_CAPACITY','ADR','LRV','ROOMS_SOLD','OCC_FCST_TOTAL','SYS_UNCON_DEMAND_TOTAL','SE_IMPACT_Y']]
    data_used_TY=data_used_TY.reset_index()
    data_used_LY=data_used[data_used['YEAR']=='LY'].loc[:,['ROOMS_SOLD','SE_IMPACT_Y']]    
    data_used_LY.rename(columns={'ROOMS_SOLD':'ROOMS_SOLD_LY','SE_IMPACT_Y':'SE_IMPACT_Y_LY'}, inplace=True)
    data_used_LY=data_used_LY.reset_index().drop('Datetime_weekday', axis=1)
#拼接上述数据,通过Index进行拼接
    data_used=data_used_TY.join(data_used_LY)
    data_used=data_used.set_index('Datetime_weekday') 
#如果Special Event中有值则其等于Capacity
    data_used['SE_IMPACT_Y']=np.where(data_used['SE_IMPACT_Y']=='0','',data_used['HOTEL_CAPACITY'])
    data_used['SE_IMPACT_Y_LY']=np.where(data_used['SE_IMPACT_Y_LY']=='0','',data_used['HOTEL_CAPACITY'])
    data_used=data_used[['ADR'
                        ,'LRV'
                        ,'ROOMS_SOLD'
                        ,'ROOMS_SOLD_LY'
                        ,'OCC_FCST_TOTAL'
                        ,'SYS_UNCON_DEMAND_TOTAL'
                        ,'SE_IMPACT_Y'
                        ,'SE_IMPACT_Y_LY']].T
    sht.range('B3:AF11').clear_contents()
    sht.range('B3').options(index=False).value = data_used
    
    return   
#===================Refresh_DOW_Distribution()=======================================
def DOW_Distribution():
    wb = xw.Book.caller()  
    sht = wb.sheets['DOW_Distribution']
#将Excel中的参数传入Python，酒店名称，Period值     
    Hotel_Name=sht.range('E1').value 
    start_time_P1=sht.range('I1').value 
    end_time_P1=sht.range('K1').value    
    start_time_P2=sht.range('M1').value 
    end_time_P2=sht.range('O1').value 
#从Cache中取数据
    s = shelve.open('PictureReport_Data.db') 
    data_used = s['key'] 
    s.close
#序列化后的数据date数据会转化为Object数据,所以需要重新转化为date数据
    data_used['OCCUPANCY_DATE']=pd.to_datetime(data_used['OCCUPANCY_DATE'])  
#日期截取需要转化成DatatimeIndex会比较方便      
    data_used=data_used[data_used['Eng_Name'] == Hotel_Name]
    data_used=data_used.set_index('OCCUPANCY_DATE')
    
    P1_DATE_Index=pd.DatetimeIndex(start=start_time_P1, end=end_time_P1, freq='D')
    data_used_P1=data_used.reindex(P1_DATE_Index, fill_value=0)
    data_used_P1['Period']='P1'    
    P2_DATE_Index=pd.DatetimeIndex(start=start_time_P2, end=end_time_P2, freq='D')
    data_used_P2=data_used.reindex(P2_DATE_Index, fill_value=0)
    data_used_P2['Period']='P2'  
    
    data_used=data_used_P1.append(data_used_P2)    
#原始数据中只有Trans和Group的ADR，需转化为Revenue才能再之后的groupby.sum(ADR为平均值)
    data_used['TRAN_ROOM_REVENUE']=data_used['TRAN_ADR'].mul(data_used['ROOMS_SOLD_TRAN'])
    data_used['GROUP_ROOM_REVENUE']=data_used['GROUP_ADR'].mul(data_used['ROOMS_SOLD_GROUP'])
#为datatime添加weekday,并按照Period和weekday进行groupby
    data_used=data_used.reset_index()
    data_used.rename(columns={'index':'OCCUPANCY_DATE'}, inplace=True)
    data_used['weekday']=data_used['OCCUPANCY_DATE'].dt.weekday.astype(str)
    data_used=data_used.groupby([data_used['Period'],data_used['weekday']]).sum()
    data_used=data_used.reset_index()
    data_used.sort_values(by='weekday',inplace=True)
    
    weekday_name = {
    '0':'Mon',
    '1':'Tue',
    '2':'Wed',
    '3':'Thu',
    '4':'Fri',
    '5':'Sat',
    '6':'Sun'    }
    
    data_used['weekday']=data_used['weekday'].map(weekday_name)
    data_used['weekday_Period']=data_used['weekday'].str.cat('_'+data_used['Period'])
    data_used=data_used.set_index('weekday_Period') 
#计算Occ,ADR,RevPAR  
    data_used['TOTAL_Occ']=data_used['ROOMS_SOLD'].div(data_used['HOTEL_CAPACITY'])
    data_used['TOTAL_ADR']=data_used['ROOM_REVENUE'].div(data_used['ROOMS_SOLD'])
    data_used['TOTAL_RevPAR']=data_used['ROOM_REVENUE'].div(data_used['HOTEL_CAPACITY'])
    
    data_used['TRAN_Occ']=data_used['ROOMS_SOLD_TRAN'].div(data_used['HOTEL_CAPACITY'])
    data_used['TRAN_ADR']=data_used['TRAN_ROOM_REVENUE'].div(data_used['ROOMS_SOLD_TRAN'])
    data_used['TRAN_RevPAR']=data_used['TRAN_ROOM_REVENUE'].div(data_used['HOTEL_CAPACITY'])
    
    data_used['GROUP_Occ']=data_used['ROOMS_SOLD_GROUP'].div(data_used['HOTEL_CAPACITY'])
    data_used['GROUP_ADR']=data_used['GROUP_ROOM_REVENUE'].div(data_used['ROOMS_SOLD_GROUP'])
    data_used['GROUP_RevPAR']=data_used['GROUP_ROOM_REVENUE'].div(data_used['HOTEL_CAPACITY'])
    
    data_used=data_used.loc[:,['TOTAL_Occ'
                               ,'TOTAL_ADR'
                               ,'TOTAL_RevPAR'
                               ,'TRAN_Occ'
                               ,'TRAN_ADR'
                               ,'TRAN_RevPAR'
                               ,'GROUP_Occ'
                               ,'GROUP_ADR'
                               ,'GROUP_RevPAR']].T
                               
    sht.range('B4:O12').clear_contents()
    sht.range('B4').options(index=False).value = data_used.values
    return           
# ===================Debug调试程序所需函数======================================
if __name__ == '__main__':
    xw.Book('IDeaS_PictureReport_Test.xlsm').set_mock_caller()
#    Load_Database()
#    Refresh_A_Glance()
#    Date_Report()
#    Business_Type_Report()
#    Forecast_Validation()
    DOW_Distribution()