#coding:utf-8
import pandas
import datetime
import pdb
today=datetime.date.today()
df = pandas.read_excel(u"C:\\Users\\admin\\Desktop\\客户分类.xls",sheet_name=None)
alist=[i for i in df]
alist.sort()
print(alist)
for i in alist:
    res=df[i]
    #返点前加上返点后
    date_aimday_start_str='2021-5-1'
    date_aimday_start=datetime.datetime.strptime(date_aimday_start_str,'%Y-%m-%d').date
    date_aimday_stop_str='2021-5-14'
    date_aimday_stop=datetime.datetime.strptime(date_aimday_stop_str,'%Y-%m-%d').date
    date_index=res.columns[res.isin(['DATE']).any().tolist().index(True)]
    aim_df=df[i][date_index]
    payment_index=res.columns[res.isin(['PAYMENT']).any().tolist().index(True)]
    balance_index=res.columns[res.isin(['BALANCE']).any().tolist().index(True)]
    amount_index=res.columns[res.isin(['AMOUNT']).any().tolist().index(True)]
    weight_index=res.columns[res.isin(['WEIGHT']).any().tolist().index(True)]
    #row_date=pandas.to_datetime((row_date-25569)*86400.0,unit='s')
    #[aim_df[i] for i in range(len(aim_df)) if pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()==row_date.date()]
    #找到今天，找到下一天，返回两个的行号，取2个行号之间的数值为目标数据，
    #res_aft_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()>date_aimday()]
    #res_bef_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()<date_aimday()]
    #res_curr_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()==date_aimday()]
    res_aft_range=[i for i in range(len(aim_df)) if isinstance(aim_df[i],datetime.datetime) and aim_df[i].date()>date_aimday_stop()]
    res_bef_range=[i for i in range(len(aim_df)) if isinstance(aim_df[i],datetime.datetime) and aim_df[i].date()<date_aimday_start()]
    res_curr_range=[i for i in range(len(aim_df)) if isinstance(aim_df[i],datetime.datetime) and aim_df[i].date()>=date_aimday_start() and aim_df[i].date()<=date_aimday_stop()]
    #如果是一个数值，说明是当天，两个数值，没今天记录,是0给数值，没数据全为0
    #print(res_aft_range,res_bef_range,res_curr_range)
    if res_curr_range==[] and res_bef_range==[] and res_aft_range!=[]: #当天没记录,过去没记录，后面有记录，返回之前账户余额。过去余额为0，未来的不算，没有返利，没有账户，还是0，全返回0，
        #print(i,0,0,0,0,0)
        if i in ('SLK'):
            print('tttttttttt')
        pass
    if res_curr_range==[] and res_bef_range==[] and res_aft_range==[]: #全没记录，返回0
        #print(i,0,0,0,0)
        continue
    if res_curr_range==[] and res_bef_range!=[] and res_aft_range==[]: #当天没记录，当天前有记录，当天后没记录，返回之前账户余额，有rebate返回rebate，没rebate，返回0
        start=res_bef_range[-1]
        stop=len(aim_df)
        bef_balance=res[balance_index][start:stop].dropna().tail(1).values[0]
        #print(i,bef_balance,0,0,bef_balance,0)
        continue
    if res_curr_range==[] and res_bef_range!=[] and res_aft_range!=[]: #当天没记录，之前有记录，之后有记录，返回之前账户余额，如果有rebate返回rebate，没rebate，返回0
        start=res_bef_range[-1]
        stop=res_aft_range[0]
        #if i=='MRAA':
        bef_balance=res[balance_index][start:stop].dropna().tail(1).values[0]
        #print(i,bef_balance,0,0,bef_balance,0)
        continue
    if res_curr_range!=[] and res_bef_range!=[] and res_aft_range==[]: #当天有记录，之前有记录,之后没记录，返回当天值和账户余额，区域直接选到frame末尾
        start=res_curr_range[0]
        stop=len(aim_df)
        ares=res[start:stop]
        bef_balance=res[balance_index][:start].dropna().tail(1).values[0]
    if res_curr_range!=[] and res_bef_range==[] and res_aft_range==[]: #当天有记录，之前没记录，之后没记录，返回当天的余额，前一天为上一行的账户余额，区域直接选到frame末尾
        start=res_curr_range[0]
        stop=len(aim_df)
        ares=res[start:stop]
        bef_balance=0
    if res_curr_range!=[] and res_bef_range!=[] and res_aft_range!=[]: #当天有记录，之前有记录，之后有记录,
        start=res_curr_range[0]
        stop=res_aft_range[0]
        bef_balance=res[balance_index][:start].dropna().tail(1).values[0]
        #if i=='KADEX':
        #    pdb.set_trace()
        balance=res[balance_index][start:stop].dropna().tail(1).values[0]
    #    print(start,stop)
    if res_curr_range!=[] and res_bef_range==[] and res_aft_range!=[]: #当天有记录，之前没有记录，之后有记录,前天余额为0，
        start=res_curr_range[0]
        stop=res_aft_range[0]
        bef_balance=0
        balance=res[balance_index][start:stop].dropna().tail(1).values[0]
    ares=res[start:stop]
    #print(ares)
    #选定区域
    #REBATE
    try:
        rebate_column=res.columns[res.isin(['REBATE']).any().tolist().index(True)]
        t1=[ares.iloc[i] for i in range(len(ares)) if 'REBATE' in ares.iloc[i].tolist()]
        if t1==[]:
            res_rebate=0
        else:
            res_rebate=t1[0][payment_index]
    except ValueError:
        res_rebate=0
    #此日充值是所有payment相加，此日销售是所有amount值相加，返点是rebate行的结果，前日余额是601行的balance，
    tt1=res[weight_index][start:stop].sum()  #计算充值的总和，如果有rebate,减去rebate
    if i in ('SLK'):
        print(i,tt1)

    date_aimday_start_str='2021-5-15'
    date_aimday_start=datetime.datetime.strptime(date_aimday_start_str,'%Y-%m-%d').date
    date_aimday_stop_str='2021-5-31'
    date_aimday_stop=datetime.datetime.strptime(date_aimday_stop_str,'%Y-%m-%d').date
    date_index=res.columns[res.isin(['DATE']).any().tolist().index(True)]
    aim_df=df[i][date_index]
    payment_index=res.columns[res.isin(['PAYMENT']).any().tolist().index(True)]
    balance_index=res.columns[res.isin(['BALANCE']).any().tolist().index(True)]
    amount_index=res.columns[res.isin(['AMOUNT']).any().tolist().index(True)]
    weight_index=res.columns[res.isin(['WEIGHT']).any().tolist().index(True)]
    #row_date=pandas.to_datetime((row_date-25569)*86400.0,unit='s')
    #[aim_df[i] for i in range(len(aim_df)) if pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()==row_date.date()]
    #找到今天，找到下一天，返回两个的行号，取2个行号之间的数值为目标数据，
    #res_aft_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()>date_aimday()]
    #res_bef_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()<date_aimday()]
    #res_curr_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()==date_aimday()]
    res_aft_range=[i for i in range(len(aim_df)) if isinstance(aim_df[i],datetime.datetime) and aim_df[i].date()>date_aimday_stop()]
    res_bef_range=[i for i in range(len(aim_df)) if isinstance(aim_df[i],datetime.datetime) and aim_df[i].date()<date_aimday_start()]
    res_curr_range=[i for i in range(len(aim_df)) if isinstance(aim_df[i],datetime.datetime) and aim_df[i].date()>=date_aimday_start() and aim_df[i].date()<=date_aimday_stop()]
    #如果是一个数值，说明是当天，两个数值，没今天记录,是0给数值，没数据全为0
    #print(res_aft_range,res_bef_range,res_curr_range)
    if res_curr_range==[] and res_bef_range==[] and res_aft_range!=[]: #当天没记录,过去没记录，后面有记录，返回之前账户余额。过去余额为0，未来的不算，没有返利，没有账户，还是0，全返回0，
        #print(i,0,0,0,0,0)
        continue
    if res_curr_range==[] and res_bef_range==[] and res_aft_range==[]: #全没记录，返回0
        #print(i,0,0,0,0)
        continue
    if res_curr_range==[] and res_bef_range!=[] and res_aft_range==[]: #当天没记录，当天前有记录，当天后没记录，返回之前账户余额，有rebate返回rebate，没rebate，返回0
        start=res_bef_range[-1]
        stop=len(aim_df)
        bef_balance=res[balance_index][start:stop].dropna().tail(1).values[0]
        #print(i,bef_balance,0,0,bef_balance,0)
        continue
    if res_curr_range==[] and res_bef_range!=[] and res_aft_range!=[]: #当天没记录，之前有记录，之后有记录，返回之前账户余额，如果有rebate返回rebate，没rebate，返回0
        start=res_bef_range[-1]
        stop=res_aft_range[0]
        #if i=='MRAA':
        bef_balance=res[balance_index][start:stop].dropna().tail(1).values[0]
        #print(i,bef_balance,0,0,bef_balance,0)
        continue
    if res_curr_range!=[] and res_bef_range!=[] and res_aft_range==[]: #当天有记录，之前有记录,之后没记录，返回当天值和账户余额，区域直接选到frame末尾
        start=res_curr_range[0]
        stop=len(aim_df)
        ares=res[start:stop]
        bef_balance=res[balance_index][:start].dropna().tail(1).values[0]
    if res_curr_range!=[] and res_bef_range==[] and res_aft_range==[]: #当天有记录，之前没记录，之后没记录，返回当天的余额，前一天为上一行的账户余额，区域直接选到frame末尾
        start=res_curr_range[0]
        stop=len(aim_df)
        ares=res[start:stop]
        bef_balance=0
    if res_curr_range!=[] and res_bef_range!=[] and res_aft_range!=[]: #当天有记录，之前有记录，之后有记录,
        start=res_curr_range[0]
        stop=res_aft_range[0]
        bef_balance=res[balance_index][:start].dropna().tail(1).values[0]
        #if i=='KADEX':
        #    pdb.set_trace()
        balance=res[balance_index][start:stop].dropna().tail(1).values[0]
    #    print(start,stop)
    if res_curr_range!=[] and res_bef_range==[] and res_aft_range!=[]: #当天有记录，之前没有记录，之后有记录,前天余额为0，
        start=res_curr_range[0]
        stop=res_aft_range[0]
        bef_balance=0
        balance=res[balance_index][start:stop].dropna().tail(1).values[0]
    ares=res[start:stop]
    if i in ('SLK'):
        print(i,tt1)
    #print(ares)
    #选定区域
    #REBATE
    try:
        rebate_column=res.columns[res.isin(['REBATE']).any().tolist().index(True)]
        t1=[ares.iloc[i] for i in range(len(ares)) if 'REBATE' in ares.iloc[i].tolist()]
        if t1==[]:
            res_rebate=0
        else:
            res_rebate=t1[0][payment_index]
    except ValueError:
        res_rebate=0
    #此日充值是所有payment相加，此日销售是所有amount值相加，返点是rebate行的结果，前日余额是601行的balance，
    tt2=res[weight_index][start:stop].sum()  #计算充值的总和，如果有rebate,减去rebate
    #print(i,tt2)
    print(i,tt1,tt2,tt1+tt2,tt1*200,tt2*200,tt1*200+tt2*200)
    


