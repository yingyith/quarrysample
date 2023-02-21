#coding:utf-8
import pandas
import datetime
import pdb
import win32com.client
from datetime import timedelta
today=datetime.date.today()
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from functools import partial
wb1 = pandas.read_excel(u"C:\\Users\\admin\\Desktop\\客户分类.xls",sheet_name=None)
wb2=load_workbook('C:\\Users\\admin\\Desktop\\总销量.xlsx')
date_aimday_start_str='2021-5-31'
date_aimday_start=datetime.datetime.strptime(date_aimday_start_str,'%Y-%m-%d') #0点0分0秒
date_aimday_stop_str='2021-6-1'
date_aimday_stop=datetime.datetime.strptime(date_aimday_stop_str,'%Y-%m-%d')
tonws=wb2['总销量']
sumres=pandas.DataFrame(tonws.values)[1:]
alist=[i for i in wb1]
for j in range(27):
    
    date_aimday_start+=timedelta(days=1)
    date_aimday_stop+=timedelta(days=1)
    res=sumres[sumres[0]>=date_aimday_start]
    res=res[res[0]<date_aimday_stop]
    dlist=res.groupby(4)[5].sum()
    inch_sum=0
    hinch_sum=0
    tquarter_sum=0
    teight_sum=0
    hc_sum=0
    sb_sum=0
    dust_sum=0
    dust_amount_sum=0
    hc_amount_sum=0
    sb_amount_sum=0
    teight_amount_sum=0
    hinch_amount_sum=0
    tquarter_amount_sum=0
    inch_amount_sum=0
   # DUST	HC	SB	3/8''	1/2''	3/4''	1''  合计
    for t in range(len(res)):
         sitem = res.iloc[t]
         if sitem[4] in ('DUST',' DUST'):
              dust_sum+=sitem[5]
              dust_amount_sum+=sitem[5]*sitem[6]
         elif sitem[4]=='HC':
              hc_sum+=sitem[5]
              hc_amount_sum+=sitem[5]*sitem[6]
         elif sitem[4]=='SB':
              sb_sum+=sitem[5]
              sb_amount_sum+=sitem[5]*sitem[6]
         elif sitem[4]=='3/8':
              teight_sum+=sitem[5]
              teight_amount_sum+=sitem[5]*sitem[6]
         elif sitem[4]=='1/2':
              hinch_sum+=sitem[5]
              hinch_amount_sum+=sitem[5]*sitem[6]
         elif sitem[4]=='3/4':
              tquarter_sum+=sitem[5]
              tquarter_amount_sum+=sitem[5]*sitem[6]
         elif sitem[4] in ['1',1]:
              inch_sum+=sitem[5]
              inch_amount_sum+=sitem[5]*sitem[6]
         #pdb.set_trace()
    display_ton_list=(dust_sum,hc_sum,sb_sum,teight_sum,hinch_sum,tquarter_sum,inch_sum)
    display_amount_list=(dust_amount_sum,hc_amount_sum,sb_amount_sum,teight_amount_sum,hinch_amount_sum,tquarter_amount_sum,inch_amount_sum)
    f1=lambda f,x:f(x,2)
    af = partial(f1,round)
    adisplay_ton_list=list(map(af,display_ton_list))
    sum_adisplay_ton_list=sum(adisplay_ton_list)
    adisplay_ton_list=[date_aimday_start.date()]+adisplay_ton_list+[sum_adisplay_ton_list]
    adisplay_ton_tuple=tuple(adisplay_ton_list)
    display_amount_list=list(display_amount_list)+[sum(display_amount_list)]
    adisplay_amount_list=tuple(map(af,display_amount_list))
    #print(*adisplay_ton_list)
    print(*adisplay_amount_list)
    #print(dlist[0],dlist[1],dlist[2],dlist[3])
    
        
    #找到客户分类在天数内的所有拉的总吨数
    
#找到总销量在天数内的该客户的总吨数
    
