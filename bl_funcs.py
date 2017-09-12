# -*- coding: utf-8 -*-
"""
Created on Fri Jun 23 15:26:52 2017

@author: skofield
"""

import pandas as pd
import numpy as np
import xlrd
from cvxopt import solvers,matrix

#定义是否为debug模式，该模式下函数会输出一些中间变量供调试
debug_mode=0


#***********************************************************************
#                                                                      *
# 定义根据取sheet中新观点P,Q及信心水平LC的函数get_pqc,参数data为               *
# 读取的excel文件数据，view_sheet为sheet号,stock_count为股票个数             *
#                                                                      *
#***********************************************************************
def getpqc(bl_view_filepath,view_sheet,stock_count):
  bl_view_data=xlrd.open_workbook(bl_view_filepath)
  cftable=bl_view_data.sheets()[view_sheet]
  view_count=0
  for j in range(7,cftable.nrows):
    if(cftable.cell(j,1).value!=''):
      view_count+=1
    else:
      break
   
    #读取观点矩阵P
  P=list(range(view_count))
  for j in range(0,view_count):
    P[j]=cftable.row_values(7+j)[1:stock_count+1]
  P=np.mat(P)

  #读取观点信心水平LC
  LC=[]
  for i in range(0,view_count):
    LC.append(cftable.cell(7+i,stock_count+1).value)

   #读取观点超额收益矩阵Q
  Q=[]
  for j in range(0,view_count):
    Q.append(cftable.cell(20+j,1).value)
  Q=np.mat(Q).T
          
  #函数返回矩阵P,Q,观点信心列表LC,观点个数view_count
  return P,Q,LC,view_count
#******************************getpqc函数定义结束*******************************




#***********************************************************************
#                                                                      *
#                     定义BL计算函数                                    *
#                                                                      *
#***********************************************************************

def bl(daily_r_path,his_start_date,his_end_date,delta,w_mkt,P,Q,LC,view_count):
  
  #打开日收益率表格的第一个sheet,即日收益率表格
  daily_r_data=xlrd.open_workbook(daily_r_path)
  daily_r_table=daily_r_data.sheets()[0]
  
  #取日收益率表格第一列的时间序列并化为datetime格式
  date=daily_r_table.col_values(0)[3:]
  date=pd.to_datetime(date)
  
  #取股票日收益率数据，转换成Series
  stock_count=daily_r_table.ncols-1
  stock_r=list(range(stock_count))
  for i in range(0,stock_count):
    stock_r[i]=daily_r_table.col_values(i+1)[3:]
    stock_r[i]=pd.Series(stock_r[i],index=date)
    stock_r[i]=stock_r[i][his_start_date:his_end_date]
    for j in range(0,len(stock_r[0])):
      stock_r[i][j]*=0.01
  #定义股票日收益率的年化协方差矩阵epsi
  epsi=[]
  for i in range(0,stock_count):
    epsi.append(stock_r[i])
  epsi=np.cov(epsi,ddof=1)*250
  epsi=np.mat(epsi)
  
  if debug_mode==1:
    print("epsi:")
    print(epsi)
    print("LC:")
    print(LC)
  



######################BL模型计算过程开始######################################
  



#&&&&&&&&&取历史数据终止日期的市值权重来计算隐含均衡收益率矩阵pai&&&&&

  #计算隐含均衡收益率矩阵pai
  pai=delta*epsi*w_mkt    
  

  if(debug_mode==1):
    print("pai:")
    print(pai)
  #pai=np.mat([0.22707,0.21833,0.19397,0.2034,0.15009,0.17566,0.16312,0.21116,0.17238])
  #pai=pai.T  
  #&&&&&&&&&&&&&&&&计算BL模型下的预期收益率矩阵E_bl&&&&&&&&&&&&&&&&&&
  #计算P_star矩阵,P_star矩阵为P按列求和生成的1*k矩阵，k表示观点数
  P_star=sum(P)
  
  if debug_mode==1:
    print('P_star:')
    print(P_star)
  
  #中间变量pep，用于计算标准刻度因子
  pep=P_star*epsi*P_star.T
  
  if debug_mode==1:
    print("pep:")
    print(pep)
  
  
  #标准刻度因子CF
  CF=float(0.5*pep)

  if debug_mode==1:
    print("CF:")
    print(CF)  
  
  

  #计算看法置信度矩阵omega
  cfli=0
  omega=[]
  for i in range(0,view_count):
    cfli+=CF/LC[i]
    omega.append(CF/LC[i])

  #计算刻度因子tao
  tao=float(pep*view_count/cfli)

  omega=np.diag(omega)
  omega=np.mat(omega)
  #omega=np.diag(P*(tao*epsi)*P.T)
  #omega=np.mat(omega)
  
  if debug_mode==1:
    
    print("omega:")
    print(omega)
  
  


  
  #计算新的加权后的收益向量E_bl
  E_bl=((tao*epsi).I+P.T*omega.I*P).I*((tao*epsi).I*pai+P.T*omega.I*Q)
  #E_bl=pai+tao*epsi*P.T*(omega+tao*P*epsi*P.T).I*(Q-P*pai)
    
  #E_bl=np.linalg.inv(te_ni+P.T*omega_ni*P)*(te_ni*pai+P.T*omega_ni*Q)
  
  #E_bl=np.mat([0.25784,0.25183,0.21006,0.22032,0.16402,0.19481,0.15941,0.23495,0.18543])
  #E_bl=E_bl.T
  
  #&&&&&&&&&&&&&&&&&&&&&&计算新的最优组合权重&&&&&&&&&&&&&&&&&&&&&&&&
  #构造cvxopt公式中的参数矩阵p,q
  p=matrix(delta*epsi)
  q=matrix(-1*E_bl)

  #构造参数矩阵A,b
  A=[]
  for i in range(0,stock_count):
    A.append(1.0)
  A=np.mat(A)
  A=matrix(A)

  b=np.mat([1.0])
  b=matrix(b)
    
  #构造参数矩阵G
  G=[]
  for i in range(0,stock_count):
    G.append(-1.0)
  G=np.diag(G)
  G=matrix(G)


  #构造参数矩阵h
  h=[]
  for i in range(0,stock_count):
    h.append(0.0)
  h=matrix(h)
  
  #计算BL模型下最优权重矩阵w_bl
  sol=solvers.qp(p,q,G,h,A,b)
  w_bl=sol['x']
  w_bl=np.mat(w_bl)
  
  if debug_mode==1:
    print("delta="+str(delta))
    print("tao="+str(tao))
    #print("pai="+str(pai))
    print("E_bl:")
    print(E_bl)
  
  #返回市值权重矩阵w_mkt和BL模型下最优权重矩阵w_bl
  return w_bl    
  
  #*****************************bl函数定义结束******************************


