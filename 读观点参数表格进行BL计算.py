# -*- coding: utf-8 -*-
"""
Created on Mon Sep  4 16:04:48 2017

@author: skofield
"""

import os
from bl_funcs import getpqc
from bl_funcs import bl
import xlrd
import numpy as np
import pandas as pd
from WindPy import w
import datetime
from xlrd import xldate_as_tuple
import xlwt

#读取当前目录作为工作目录
working_dir=os.getcwd()

#定义BL初始化参数文件路径
bl_ini_filepath=working_dir+"\\"+"bl_ini.xls"

#定义资产日收益率表格路径
daily_r_path=working_dir+"\\"+"资产日收益率.xls"

#定义BL观点参数表格路径
bl_view_filepath=working_dir+"\\"+"bl_view.xls"

#打开BL初始化参数文件、资产日收益率文件、BL观点参数文件
bl_ini_data=xlrd.open_workbook(bl_ini_filepath)
daily_r_data=xlrd.open_workbook(daily_r_path)
bl_view_data=xlrd.open_workbook(bl_view_filepath)

#读取BL初始化参数文件中的交易开始日期和结束日期并格式化处理
bl_ini_table=bl_ini_data.sheets()[0]
trade_start_date=bl_ini_table.cell(11,1).value
trade_end_date=bl_ini_table.cell(13,1).value
trade_start_date=datetime.datetime(*xldate_as_tuple(trade_start_date,0)).date()
trade_end_date=datetime.datetime(*xldate_as_tuple(trade_end_date,0)).date()
trade_start_date=str(trade_start_date)
trade_end_date=str(trade_end_date)

#取BL初始化参数文件中的初始市场权重
w_mkt=bl_ini_table.row_values(3)[1:]
w_mkt=np.mat(w_mkt).T

#取BL初始化参数文件中的股票代码列表并计数
stock_list=bl_ini_table.row_values(0)[1:]
stock_count=len(stock_list)

#读取资产名称列表stock_name
stock_name=bl_ini_table.row_values(1)[1:]
            
#读取风险厌恶系数delta
delta=bl_ini_table.cell(5,1).value

#读取回溯天数
recall_days=bl_ini_table.cell(15,1).value                       
                       
#读取资产日收益率表格数据构造dataframe
daily_r_df=pd.read_excel(daily_r_path,sheetname=0,header=0,index_col=0)
daily_r_df.index=pd.to_datetime(daily_r_df.index)

#取交易区间内的交易日，构造交易时间序列
trade_series=daily_r_df[trade_start_date:trade_end_date].index


###############################################################################
#                                                                             #
#                             交易开始                                          #
#                                                                             #
###############################################################################

#在交易区间内，按月调仓进行操作。每进入下个月，取最近一年历史统计数据，按需读取新观点
view_sheet=-1
cur_month=''

#定义BL组合净值list
port_netval=[]

#新建一个excel，用于记录每月BL计算出的最新权重配置
monthly_bl_w=xlwt.Workbook()
#定义上面的excel存放路径
bl_result_filepath=working_dir+'\\output\\bl_result.xls'

#对交易日期进行遍历循环
for i in range(0,len(trade_series)):
  #判断是否进入下个月
  if(str(trade_series[i])[0:7]!=cur_month):
    #进入新的月份后，按照回溯天数重新回溯数据
    recall_start_date=trade_series[i]-datetime.timedelta(recall_days)
    recall_end_date=trade_series[i]-datetime.timedelta(1)
    #当前月份重新赋值
    cur_month=str(trade_series[i])[0:7]
    
    #取下一个sheet
    view_sheet+=1    

    #判断本月对应的sheet是否有新观点，如有，则读取新观点
    if bl_view_data.sheets()[view_sheet].cell(1,0).value==1:
      P,Q,LC,view_count=getpqc(bl_view_filepath,view_sheet,stock_count)

    #计算本月BL权重
    w_bl=bl(daily_r_path,recall_start_date,recall_end_date,delta,w_mkt,P,Q,LC,view_count)
    
    
    #在excel中新建一个sheet，记录本月计算出的权重结果
    new_sheet=monthly_bl_w.add_sheet(cur_month,cell_overwrite_ok=True)
    
    #定义BL结果写入excel时显示保留四位小数
    style1=xlwt.XFStyle()
    fmt='##0.0000'
    style1.num_format_str=fmt
    new_sheet.write(0,0,'资产列表: ')
    new_sheet.write(3,0,'当月BL权重: ')
    new_sheet.col(0).width = (len('当月BL权重')*460)
    for j in range(0,len(stock_list)):
      new_sheet.write(0,j+1,stock_list[j])
      new_sheet.write(1,j+1,stock_name[j])
      new_sheet.write(3,j+1,float(w_bl[j]),style1)
      new_sheet.col(j+1).width = (len('沪深300工业')*460) #设置excel列宽
    

    
  
  #计算每日组合净值
  net_val=float(np.mat(daily_r_df[str(trade_series[i])]*0.01+1)*w_bl)
  port_netval.append(net_val)

#将每月的BL权重配置保存到文件
monthly_bl_w.save(bl_result_filepath)

#对BL组合日净值list加上时间序列索引，构造为Series
port_netval=pd.Series(port_netval,index=trade_series,name='BL组合日净值')
  
#定义每日组合净值保存到csv文件的路径，将每日净值存入
port_netval_filepath=working_dir+'\\output\\port_netval.csv'
port_netval.to_csv(port_netval_filepath)







