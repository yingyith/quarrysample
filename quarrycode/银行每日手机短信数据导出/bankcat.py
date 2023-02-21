#coding:utf-8
import pandas
import datetime
import pdb
from accountmap import account_map
def str2date(str2):
    date_aimday=datetime.datetime.strptime(str2,'%Y/%m/%d').date()
    return date_aimday

today=datetime.date.today()
df = pandas.read_excel(u"C:\\Users\\admin\\Documents\\yingyikuang\\bank_pay_record\\sms.xlsx",sheet_name=None)
alist=['sms']
for i in alist:
    #if i not in ('KADEX'):
    #    continue
    res=df[i]
    date_aimday_str='2021/6/28'
    date_aimday=datetime.datetime.strptime(date_aimday_str,'%Y/%m/%d').date()
    aim_df=df[i]['Date']
    #row_date=pandas.to_datetime((row_date-25569)*86400.0,unit='s')
    #[aim_df[i] for i in range(len(aim_df)) if pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()==row_date.date()]
    #找到今天，找到下一天，返回两个的行号，取2个行号之间的数值为目标数据，
    #res_aft_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()>date_aimday()]
    #res_bef_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()<date_aimday()]
    #res_curr_range=[i for i in range(len(aim_df)) if type(aim_df[i])==int and pandas.to_datetime((int(aim_df[i])-25569)*86400,unit='s').date()==date_aimday()]
    res_curr_range=[i for i in range(len(aim_df)) if isinstance(aim_df[i],str) and str2date(aim_df[i].split(" ")[0])==date_aimday]
    #如果是一个数值，说明是当天，两个数值，没今天记录,是0给数值，没数据全为0
    #print(res_curr_range)
    start=res_curr_range[0]
    stop=res_curr_range[-1]+1
    #print(res[start:stop])
    ares=res[start:stop]
    for j in ares['Converted']:
        #print(j)
        if j=='':
            continue
        if type(j)==float:
            continue
        if j[:4]=="Acct":
            item=j.split('\n')
            #print(item)
            flag=True
            for z in account_map:
                for t in account_map[z]:
                    res=[q in item[2] for q in t.split(' ')]
                    #print("t is---")
                    #print(t)
                    #print(res)
                    res=all(res)
                    #print(res)
                    if res==True:
                        #print(item[3],item[4],item[1])
                        tf_time=item[1].split(":")[1]
                        #print(tf_time)
                        amount=item[-3].split(":")[1]
                        print(tf_time,"|",item[2],"|",amount,"|",item[-2],"|",z)
                        #输出匹配到的账号
                        flag=False
                        break
                    else:
                        continue
            if flag==True:
                #没匹配到
                #pdb.set_trace()
                tf_time=item[1].split(":")[1]
                #print(tf_time)
                amount=item[-3].split(":")[1]
                print(tf_time,"|",item[2],"|",amount,"|",item[-2],"|","ZZNone")
        if j[:6]=="Credit":
            item=j.split('\n')
            if item[2]!="Acc:147******633":
            #if item[2]!="Acc:009******335":
                continue
            #print(item)
            flag=True
            for z in account_map:
                for t in account_map[z]:
                    res=[q in item[3] for q in t.split(' ')]
                    res=all(res)
                    if res==True:
                        #print(item[3],item[4],item[1])
                        tf_time=item[4].split(":")[1]
                        #print(tf_time)
                        amount=item[1].split(":")[1][3:]
                        print(tf_time,"|",item[3],"|",amount,"|",item[-2],"|",z)
                        #输出匹配到的账号
                        flag=False
                        break
                    else:
                        continue
            if flag==True:
                #没匹配到
                #pdb.set_trace()
                tf_time=item[4].split(":")[1]
                #print(tf_time)
                amount=item[1].split(":")[1][3:]
                print(tf_time,"|",item[3],"|",amount,"|",item[-2],"|","ZZNone")


    #print("账户:",i,"/n该天充值:",t1,"/n该天返利:",t2,"/n该天销售:",t3,"/n该天账户:",t4,"/n上天账户:",t5)
    


