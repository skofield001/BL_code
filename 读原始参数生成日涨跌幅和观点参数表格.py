# -*- coding: utf-8 -*-
"""
Created on Fri Sep  1 16:01:08 2017

@author: skofield
"""
import os,sys
import xlrd
import xlwt
from xlutils.copy import copy
from WindPy import w
import datetime
import pandas as pd
import time
from xlrd import xldate_as_tuple

#取当前目录作为工作目录,该目录用于程序运行所有文件的生成
working_dir=os.getcwd()

print("当前工作目录为："+working_dir)  

#定义BL初始化参数文件路径
ini_file_path=working_dir+"\\" + "bl_ini.xls"


#打开BL初始化参数文件，读取第一个sheet
ini_data=xlrd.open_workbook(ini_file_path)
ini_table=ini_data.sheets()[0]

#读取资产代码列表stock_list和资产名称stock_name
stock_list=ini_table.row_values(0)[1:]
stock_name=ini_table.row_values(1)[1:]


#读取历史数据开始日期和结束日期并格式化处理，用于生成历史收益率表格
his_start_date=ini_table.cell(7,1).value
his_end_date=ini_table.cell(9,1).value
his_start_date=datetime.datetime(*xldate_as_tuple(his_start_date,0)).date()
his_end_date=datetime.datetime(*xldate_as_tuple(his_end_date,0)).date()

#读取交易开始日期和结束日期并格式化处理，用于生成BL观点表格
trade_start_date=ini_table.cell(11,1).value
trade_end_date=ini_table.cell(13,1).value
trade_start_date=datetime.datetime(*xldate_as_tuple(trade_start_date,0)).date()
trade_end_date=datetime.datetime(*xldate_as_tuple(trade_end_date,0)).date()



#*****************************************************************************
#                                                                            *
#                           生成资产日收益率表格                                 *
#                                                                            * 
#*****************************************************************************

#定义资产日收益率表格生成路径
daily_r_path=working_dir+"\\"+"资产日收益率.xls"

#从wind读取数据，生成日收益率数据表格
w.start()

stock_name=[]
for i in range(0,len(stock_list)):
  stock_name.append(w.wsd(stock_list[i],'sec_name').Data[0][0])



his_date=w.tdays(his_start_date,his_end_date).Data[0]
for i in range(0,len(his_date)):
  his_date[i]=his_date[i].date()
his_date=pd.to_datetime(his_date)


             
stock_r=list(range(len(stock_list)))
for i in range(0,len(stock_list)):
  stock_r[i]=w.wsd(stock_list[i],'pct_chg',his_start_date,his_end_date,\
  'PriceAdj=F').Data[0]
  stock_r[i]=pd.Series(stock_r[i],index=his_date)

#创建资产日收益率工作簿
assets=xlwt.Workbook() 

#创建第一个sheet：
sheet1=assets.add_sheet(u'资产日收益率',cell_overwrite_ok=True)
#生成第一行
for i in range(0,len(stock_list)):
  sheet1.write(0,i+1,stock_list[i])
  sheet1.write(1,i+1,stock_name[i])
  sheet1.col(i+1).width = (len('沪深300工业')*460)

for i in range(0,len(stock_list)):
  sheet1.write(2,i+1,'涨跌幅(%)')


for i in range(0,len(his_date)):
  sheet1.write(i+3,0,his_date[i].strftime("%Y-%m-%d"))
sheet1.col(0).width = (len('yyyy-mm-dd')*300)

style1=xlwt.XFStyle()
fmt='##0.0000'
style1.num_format_str=fmt
for i in range(0,len(stock_list)):
  for j in range(0,len(his_date)):
    sheet1.write(j+3,i+1,stock_r[i][j],style1)

assets.save(daily_r_path)

print("资产日收益率表格生成完毕！表格路径为："+daily_r_path)




#*****************************************************************************
#                                                                            *
#                           生成观点参数表格                                    *
#                                                                            * 
#*****************************************************************************

#定义BL观点参数表格生成路径
bl_view_file_path=working_dir+"\\"+"bl_view.xls"

#创建BL观点参数表格
bl_view=xlwt.Workbook()

#取模拟交易区间的所有月份
begin_date=trade_start_date
end_date=trade_end_date
trade_month_list=[]
while begin_date <= end_date:
  date_str=begin_date.strftime("%Y%m")
  trade_month_list.append(date_str)
  begin_date+=datetime.timedelta(days=1)
trade_month_list=list(set(trade_month_list))
trade_month_list.sort()


#对模板文件进行操作，每个交易月份增加一个sheet，用于输入相关信息
for i in trade_month_list:
  ws=bl_view.add_sheet(i+"观点",cell_overwrite_ok=True)
  temp_str=i+'是否有新观点:(如有则在第二行第一列单元格填1,并在下方输入新的P、Q)'
  ws.write(0,0,temp_str)
  ws.write(3,1,'新的观点矩阵P：')
  ws.write(7,0,'观点1:')
  ws.write(8,0,'观点2:')
  ws.write(9,0,'...')
  ws.write(19,1,'新的观点超额收益矩阵Q：')
  ws.write(20,0,'观点1超额收益：')
  ws.write(21,0,'观点2超额收益：')
  ws.write(22,0,'...')
  ws.col(0).width = (len('观点2超额收益：')*460)
  for j in range(0,len(stock_list)):
    ws.write(5,j+1,stock_list[j])
    ws.col(j+1).width = (len('沪深300工业')*460)
    ws.write(6,j+1,stock_name[j])
  ws.write(6,len(stock_list)+1,'信心水平')
bl_view.save(bl_view_file_path)

print("BL参数表格生成完毕，表格路径为："+bl_view_file_path+",请打开表格填写相关参数！")

