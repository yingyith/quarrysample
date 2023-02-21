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
wb1 = pandas.read_excel(u"C:\\Users\\admin\\Desktop\\客户分类.xls",sheet_name=None)
wb2=load_workbook('C:\\Users\\admin\\Desktop\\总销量.xlsx')
date_aimday_start_str='2021-5-1'
date_aimday_start=datetime.datetime.strptime(date_aimday_start_str,'%Y-%m-%d') #0点0分0秒
date_aimday_stop_str='2021-5-2'
date_aimday_stop=datetime.datetime.strptime(date_aimday_stop_str,'%Y-%m-%d')
tonws=wb2['总销量']
sumres=pandas.DataFrame(tonws.values)[1:]
alist=[i for i in wb1]
print(alist)
for j in range(60):
    date_aimday_start+=timedelta(days=1)
    date_aimday_stop+=timedelta(days=1)
    res=sumres[sumres[0]>=date_aimday_start]
    res=res[res[0]<date_aimday_stop]
    tons_time_sum=res.groupby(3)[5].sum()
    price_time_sum=res.groupby(3)[6].sum()
    for i in alist:
        clientws=wb1[i]
        date_index=clientws.columns[clientws.isin(['DATE']).any().tolist().index(True)]
        weight_index=clientws.columns[clientws.isin(['WEIGHT']).any().tolist().index(True)]
        price_index=clientws.columns[clientws.isin(['PRICE']).any().tolist().index(True)]
        date_wb1=wb1[i][date_index]
        aim_df=clientws[date_index]
        res_aft_range=[i for i in range(len(clientws)) if isinstance(date_wb1[i],datetime.datetime) and date_wb1[i]>date_aimday_start]
        res_bef_range=[i for i in range(len(clientws)) if isinstance(date_wb1[i],datetime.datetime) and date_wb1[i]<date_aimday_start]
        res_curr_range=[i for i in range(len(clientws)) if isinstance(date_wb1[i],datetime.datetime) and date_wb1[i]==date_aimday_start]
        if res_curr_range==[] and res_bef_range==[] and res_aft_range!=[]: #当天没记录,过去没记录，后面有记录，返回之前账户余额。过去余额为0，未来的不算，没有返利，没有账户，还是0，全返回0，
            continue
        if res_curr_range==[] and res_bef_range==[] and res_aft_range==[]: #全没记录，返回0
            continue
        if res_curr_range==[] and res_bef_range!=[] and res_aft_range==[]: #当天没记录，当天前有记录，当天后没记录，返回之前账户余额，有rebate返回rebate，没rebate，返回0
            start=res_bef_range[-1]
            stop=len(aim_df)
            continue
        if res_curr_range==[] and res_bef_range!=[] and res_aft_range!=[]: #当天没记录，之前有记录，之后有记录，返回之前账户余额，如果有rebate返回rebate，没rebate，返回0
            start=res_bef_range[-1]
            stop=res_aft_range[0]
            #if i=='MRAA':
            continue
        if res_curr_range!=[] and res_bef_range!=[] and res_aft_range==[]: #当天有记录，之前有记录,之后没记录，返回当天值和账户余额，区域直接选到frame末尾
            start=res_curr_range[0]
            stop=len(aim_df)
            ares=res[start:stop]
        if res_curr_range!=[] and res_bef_range==[] and res_aft_range==[]: #当天有记录，之前没记录，之后没记录，返回当天的余额，前一天为上一行的账户余额，区域直接选到frame末尾
            start=res_curr_range[0]
            stop=len(aim_df)
            ares=res[start:stop]
            bef_balance=0
        if res_curr_range!=[] and res_bef_range!=[] and res_aft_range!=[]: #当天有记录，之前有记录，之后有记录,
            start=res_curr_range[0]
            stop=res_aft_range[0]
        #    print(start,stop)
        if res_curr_range!=[] and res_bef_range==[] and res_aft_range!=[]: #当天有记录，之前没有记录，之后有记录,前天余额为0，
            start=res_curr_range[0]
            stop=res_aft_range[0]
        ares=clientws[start:stop]
        tt=ares.isin(['REBATE']).any().tolist()
        if True in tt:
            #pdb.set_trace()
            arebate_index=ares.columns[tt.index(True)]
            rebate_row=ares.loc[ares[arebate_index]=='REBATE'].index.values[0]
            end_row=len(ares)-(ares.tail(1).index[0]-rebate_row+1)
            ares=ares[:end_row]
        weight_sum=ares[weight_index].sum()
        price_sum=ares[price_index].sum()
        #if i=='D1':
           #pdb.set_trace()
        if tons_time_sum.get(i,'None')=='None' and weight_sum!=0:
            print(i+"---not pass")
            print(date_aimday_start,None,str(weight_sum),None,str(price_sum))
            #print(i+"---check_pass")
            continue
        if tons_time_sum.get(i,'None')=='None' and weight_sum==0:
            continue
        if abs(tons_time_sum[i]-weight_sum)<0.1 and price_time_sum[i]==price_sum:
            #print(i+"---check_pass")
            continue
        else:
            print(i+"---not pass")
            print(date_aimday_start,str(tons_time_sum[i]),str(weight_sum),str(price_time_sum[i]),str(price_sum))
            
        rownumber=0
        #pdb.set_trace()
        #res.groupby(i).sum()
        #pdb.set_trace()
    
        
    #找到客户分类在天数内的所有拉的总吨数
    
#找到总销量在天数内的该客户的总吨数
    
