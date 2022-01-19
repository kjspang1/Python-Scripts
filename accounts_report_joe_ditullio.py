# -*- coding: utf-8 -*-
"""
Created on Thu Dec 23 12:52:40 2021

@author: kevinspang
"""

import gspread 
import datetime as dt
import mysql.connector
import xlwings as xw
from oauth2client.service_account import ServiceAccountCredentials
from login import mysqlLogin

Login = mysqlLogin()
user = Login[0]
password = Login[1]
db = Login[2]
host = Login[3]

cnx = mysql.connector.connect(user = user,
                              password = password,
                              db = db,
                              host = host)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('D:\\Users\\kevinspang\\Documents\\Python_Scripts\\client_secret.json', scope)
client = gspread.authorize(creds)

account_data = client.open("AcctsDepartment_2021_V3")
accounts = account_data.get_worksheet(1)


#################################################################################################################################
#date and quarter logic
today_save = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
today_infile = dt.datetime.strftime(dt.datetime.now(),'%m/%d/%Y')


dformat = '%Y-%m-%d'
today = dt.datetime.strftime(dt.datetime.now(),dformat)
year = dt.datetime.strftime(dt.datetime.now(),'%Y')
q1start = dt.datetime.strftime(dt.datetime.strptime(year + '-01-01',dformat),dformat)
q1end = dt.datetime.strftime(dt.datetime.strptime(year + '-03-31',dformat),dformat)
q2start = dt.datetime.strftime(dt.datetime.strptime(year + '-04-01',dformat),dformat)
q2end = dt.datetime.strftime(dt.datetime.strptime(year + '-06-30',dformat),dformat)
q3start = dt.datetime.strftime(dt.datetime.strptime(year + '-07-01',dformat),dformat)
q3end = dt.datetime.strftime(dt.datetime.strptime(year + '-09-30',dformat),dformat)
q4start = dt.datetime.strftime(dt.datetime.strptime(year + '-10-01',dformat),dformat)
q4end = dt.datetime.strftime(dt.datetime.strptime(year + '-12-31',dformat),dformat)

if q1start <= today <= q1end:
    quartername = 'Q1'
    quarterstart = q1start
    quarterend = q1end
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
elif q2start <= today <= q2end:
    quartername = 'Q2'
    quarterstart = q2start
    quarterend = q2end
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
elif q3start <= today <= q3end:
    quartername = 'Q3'
    quarterstart = q3start
    quarterend = q3end
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
else:
    quartername = 'Q4'
    quarterstart = q4start
    quarterend = q4end 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 


##################################################################################################################################

def Convert(string):
    li = list(string.split(","))
    return li

#MAKE SURE THESE LINE UP WITH ACCOUNTS TABLE!!!
team_ids = accounts.acell('M37').value
team_roles = accounts.acell('N37').value
unsupported = accounts.acell('E41').value
supported = accounts.acell('F41').value

team_list = str(team_ids)
converted_team_list= Convert(team_ids)
roles = Convert(team_roles)
query_result = []

print("Running Joe's Team")
c = cnx.cursor()
while len(converted_team_list) > 0:
    name_id = converted_team_list.pop(0)
    query_individual = ('''#Select name
select concat(first_name, " ", last_name) name 
from fos_user where id in ({})
union all

#Active Accounts
Select count(distinct li.client_id) active_accounts 
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and li.account_managers is not null
and li.account_managers not like '%support%'
union all

#Active Agency Accounts
Select count(distinct li.client_id) active_agency_accounts  
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and lower(type = 'agency')
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day) 
and lower(li.account_managers) not like '%support%'
and li.account_managers is not null
union all

#Active Publisher Accounts
Select count(distinct li.client_id) active_publisher_accounts
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and lower(type = 'publisher')
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and lower(li.account_managers) not like '%support%' 
and li.account_managers is not null
union all

#Solo Accounts
select count(client_id) solo_accounts 
from (select client_id from account_manager 
where client_id in (select id from client where is_active = 1)
and client_id in (select client_id from account_manager where user_id in ({}))
group by client_id having count(client_id) = '1'
union all
select client_id from (select client_id, count(user_id) count, 
sum(case when is_manager = 1 then 1 else 0 end) as manager from account_manager 
where client_id in (select id from client where is_active = 1)
and client_id in (select client_id from account_manager where user_id in ({})) 
group by client_id) a where count > 1 and manager > 0) a
where client_id in (select distinct client_id from line_item where 
(sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day))
union all

#Inactive Accounts
select case when total_accounts - active_accounts < 0 then 0
else total_accounts - active_accounts end as inactive_inactive_accounts
from (Select count(am.id) total_accounts
from account_manager am
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and client_id not in (select client_id from 
line_item where account_managers like '%support%')
and cl.is_active = 1) a,
(Select count(distinct li.client_id) active_accounts 
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and lower(li.account_managers) not like '%support%'
and li.account_managers is not null) b
union all

#Total Accounts 
Select count(am.id) total_accounts
from account_manager am
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and client_id not in (select client_id from 
line_item where account_managers like '%support%')
and cl.is_active = 1
union all

#Active Agency Line items
select count(id) active_line_items_agency 
from line_item where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'agency')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
union all

#active line items agency AC unsupported
select count(id) agency_ac_unsupported
from line_item where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'agency')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
{}
union all

#Active Line Items Agency AC Supported
select count(id) agency_ac_supported
from line_item where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'agency')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
{}
union all

#active line items publisher
select count(id) active_line_item_publisher 
from line_item where client_id in (Select distinct client_id 
from account_manager
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'publisher')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
union all

#total active line items
select count(id) total_active_line_items
from line_item where client_id in (Select distinct client_id 
from account_manager
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
union all

#count of active deal ids
select count(id) active_deal_id_count 
from sales_forecast 
where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_date and end_date
or start_date between sysdate() and sysdate() + interval 7 day)
and client_id not in (select client_id from 
line_item where account_managers like '%support%')
and deal_stage = 4
union all

#active agency revenue
select coalesce(agency_active_revenue,0) as agency_active_revenue from (
select round(agency_active_revenue/7,2) as agency_active_revenue from(select case
when days_left between 0 and 7 then sum(((total_budget)/datediff(end_time, start_time))*days_left)
else  sum(((total_budget)/datediff(end_time, start_time))*7) end as agency_active_revenue from line_item 
where client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between start_time and date_add(end_time, interval 6 day)
and client_id in (select id from client where type = 'agency')
and lower(account_managers) not like '%support%')a)b
union all

#active publisher revenue
select coalesce(publisher_active_revenue,0) as publisher_active_revenue from (
select round(publisher_active_revenue/7,2) as publisher_active_revenue from(select case
when li.days_left between 0 and 7 then sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*days_left)
else  sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*7)
end as publisher_active_revenue from line_item li join client cl
on cl.id = li.client_id
join client_setting cs 
on  cs.id = cl.clientSetting_id
where li.client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between li.start_time and date_add(li.end_time, interval 6 day)
and li.client_id in (select id from client where type = 'publisher')
and lower(li.account_managers) not like '%support%')a)b
union all

#active total revenue
select coalesce(sum(total_active_revenue),0) active_total_revenue from(
select round(total_active_revenue/7,2) as total_active_revenue from(select case
when days_left between 0 and 7 then sum(((total_budget)/datediff(end_time, start_time))*days_left)
else  sum(((total_budget)/datediff(end_time, start_time))*7) end as total_active_revenue from line_item 
where client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between start_time and date_add(end_time, interval 6 day)
and client_id in (select id from client where type = 'agency')
and lower(account_managers) not like '%support%')a union all
select round(total_active_revenue/7,2) as total_active_revenue from(select case
when li.days_left between 0 and 7 then sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*days_left)
else  sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*7)
end as total_active_revenue from line_item li join client cl
on cl.id = li.client_id
join client_setting cs 
on  cs.id = cl.clientSetting_id
where li.client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between li.start_time and date_add(li.end_time, interval 6 day)
and li.client_id in (select id from client where type = 'publisher')
and lower(li.account_managers) not like '%support%')a)b
union all

#Count of expected deal ids
select count(id) expected_deal_ids 
from sales_forecast 
where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and start_date between sysdate() and sysdate() + interval 90 day
and client_id not in (select client_id from line_item where account_managers like '%support%')
and deal_stage=4
union all

#Count of expected subdeal ids
select count(distinct sd.id) expected_subdeal_ids
from sales_forecast_subdeal sd
join sales_forecast sf
on sd.salesForecast_id = sf.id
where sf.client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and sd.start_date between sysdate() and sysdate() + interval 90 day
and client_id not in (select client_id from line_item where account_managers like '%support%')
and deal_stage = 4
union all

#Future Agency Revenue
select coalesce(budget, 0) budget from (
select sum(budget) budget
from sales_forecast 
where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and start_date between sysdate() and sysdate() + interval 90 day
and client_id not in (select client_id from line_item where account_managers like '%support%')
and deal_stage=4) a'''.format(name_id,name_id,name_id,name_id,name_id,name_id,name_id,name_id,name_id,name_id,name_id,
                        unsupported,name_id,supported,name_id,name_id,name_id,name_id,name_id,name_id,name_id,name_id,name_id,name_id))
    c.execute(query_individual)
    result = [item[0] for item in c.fetchall()]
    query_result.append(result)

query_team = ('''#Select name
              Select 'Team Total'
              union all

#Active Accounts
Select count(distinct li.client_id) active_accounts 
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and li.account_managers is not null
and li.account_managers not like '%support%'
union all

#Active Agency Accounts
Select count(distinct li.client_id) active_agency_accounts  
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and lower(type = 'agency')
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day) 
and lower(li.account_managers) not like '%support%'
and li.account_managers is not null
union all

#Active Publisher Accounts
Select count(distinct li.client_id) active_publisher_accounts
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and lower(type = 'publisher')
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and lower(li.account_managers) not like '%support%' 
and li.account_managers is not null
union all

#Solo Accounts
select count(client_id) solo_accounts 
from (select client_id from account_manager 
where client_id in (select id from client where is_active = 1)
and client_id in (select client_id from account_manager where user_id in ({}))
group by client_id having count(client_id) = '1'
union all
select client_id from (select client_id, count(user_id) count, 
sum(case when is_manager = 1 then 1 else 0 end) as manager from account_manager 
where client_id in (select id from client where is_active = 1)
and client_id in (select client_id from account_manager where user_id in ({})) 
group by client_id) a where count > 1 and manager > 0) a
where client_id in (select distinct client_id from line_item where 
(sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day))
union all

#Inactive Accounts
select case when total_accounts - active_accounts < 0 then 0
else total_accounts - active_accounts end as inactive_inactive_accounts
from (Select count(am.id) total_accounts
from account_manager am
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and client_id not in (select client_id from 
line_item where account_managers like '%support%')
and cl.is_active = 1) a,
(Select count(distinct li.client_id) active_accounts 
from line_item li
join account_manager am
on  li.client_id = am.client_id 
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and cl.is_active = 1
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and lower(li.account_managers) not like '%support%'
and li.account_managers is not null) b
union all

#Total Accounts 
Select count(am.id) total_accounts
from account_manager am
join client cl
on cl.id = am.client_id
where am.user_id in ({})
and client_id not in (select client_id from 
line_item where account_managers like '%support%')
and cl.is_active = 1
union all

#Active Agency Line items
select count(id) active_line_items_agency 
from line_item where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'agency')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
union all

#active line items agency AC unsupported
select count(id) agency_ac_unsupported
from line_item where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'agency')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
{}
union all

#Active Line Items Agency AC Supported
select count(id) agency_ac_supported
from line_item where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'agency')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
{}
union all

#active line items publisher
select count(id) active_line_item_publisher 
from line_item where client_id in (Select distinct client_id 
from account_manager
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and client_id in (select id from client where type = 'publisher')
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
union all

#total active line items
select count(id) total_active_line_items
from line_item where client_id in (Select distinct client_id 
from account_manager
where user_id in ({}))
and (sysdate() between start_time and end_time
or start_time between sysdate() and sysdate() + interval 7 day)
and lower(ad_server_status) in ('active','delivering','ready')
and lower(account_managers) not like '%support%'
and account_managers is not null
union all

#count of active deal ids
select count(id) active_deal_id_count 
from sales_forecast 
where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and (sysdate() between start_date and end_date
or start_date between sysdate() and sysdate() + interval 7 day)
and client_id not in (select client_id from 
line_item where account_managers like '%support%')
and deal_stage = 4
union all

#active agency revenue
select coalesce(agency_active_revenue,0) as agency_active_revenue from (
select round(agency_active_revenue/7,2) as agency_active_revenue from(select case
when days_left between 0 and 7 then sum(((total_budget)/datediff(end_time, start_time))*days_left)
else  sum(((total_budget)/datediff(end_time, start_time))*7) end as agency_active_revenue from line_item 
where client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between start_time and date_add(end_time, interval 6 day)
and client_id in (select id from client where type = 'agency')
and lower(account_managers) not like '%support%')a)b
union all

#active publisher revenue
select coalesce(publisher_active_revenue,0) as publisher_active_revenue from (
select round(publisher_active_revenue/7,2) as publisher_active_revenue from(select case
when li.days_left between 0 and 7 then sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*days_left)
else  sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*7)
end as publisher_active_revenue from line_item li join client cl
on cl.id = li.client_id
join client_setting cs 
on  cs.id = cl.clientSetting_id
where li.client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between li.start_time and date_add(li.end_time, interval 6 day)
and li.client_id in (select id from client where type = 'publisher')
and lower(li.account_managers) not like '%support%')a)b
union all

#active total revenue
select coalesce(sum(total_active_revenue),0) active_total_revenue from(
select round(total_active_revenue/7,2) as total_active_revenue from(select case
when days_left between 0 and 7 then sum(((total_budget)/datediff(end_time, start_time))*days_left)
else  sum(((total_budget)/datediff(end_time, start_time))*7) end as total_active_revenue from line_item 
where client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between start_time and date_add(end_time, interval 6 day)
and client_id in (select id from client where type = 'agency')
and lower(account_managers) not like '%support%')a union all
select round(total_active_revenue/7,2) as total_active_revenue from(select case
when li.days_left between 0 and 7 then sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*days_left)
else  sum(((li.goal_impressions/1000 *cs.dpm_cpm)/datediff(li.end_time, li.start_time))*7)
end as total_active_revenue from line_item li join client cl
on cl.id = li.client_id
join client_setting cs 
on  cs.id = cl.clientSetting_id
where li.client_id in (Select distinct client_id from account_manager where user_id in ({}))
and sysdate() between li.start_time and date_add(li.end_time, interval 6 day)
and li.client_id in (select id from client where type = 'publisher')
and lower(li.account_managers) not like '%support%')a)b
union all

#Count of expected deal ids
select count(id) expected_deal_ids 
from sales_forecast 
where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and start_date between sysdate() and sysdate() + interval 90 day
and client_id not in (select client_id from line_item where account_managers like '%support%')
and deal_stage=4
union all

#Count of expected subdeal ids
select count(distinct sd.id) expected_subdeal_ids
from sales_forecast_subdeal sd
join sales_forecast sf
on sd.salesForecast_id = sf.id
where sf.client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and sd.start_date between sysdate() and sysdate() + interval 90 day
and client_id not in (select client_id from line_item where account_managers like '%support%')
and deal_stage = 4
union all

#Future Agency Revenue
select coalesce(budget, 0) budget from (
select sum(budget) budget
from sales_forecast 
where client_id in (Select distinct client_id 
from account_manager 
where user_id in ({}))
and start_date between sysdate() and sysdate() + interval 90 day
and client_id not in (select client_id from line_item where account_managers like '%support%')
and deal_stage=4) a'''.format(team_list,team_list,team_list,team_list,team_list,team_list,team_list,team_list,team_list,team_list,unsupported,team_list,
            supported,team_list,team_list,team_list,team_list,team_list,team_list,team_list,team_list,team_list,team_list))
c.execute(query_team)
team_result = [item[0] for item in c.fetchall()]
query_result.append(team_result)


cnx.close()

#########################################################################################################################################################################
app = xw.App(visible=False)
wb1 = xw.Book("D://Users//kevinspang//Documents//Excel//Accounts Reports//template.xlsx")
wb2 = xw.Book("D://Users//kevinspang//Documents//Excel//Accounts Reports//temp.xlsx")
sheet_master = wb1.sheets['Master']
sheet_temp = wb2.sheets['temp']
top_date = quartername + ": " + today_infile
sheet_master.range("A1").value = top_date
sheet_master.range("A2").value = "Days Remaining in Quarter: " + str(daysRemaining)
sheet_temp.range("A1").options(transpose=True).value = query_result
col_num = int(len(roles) + 5)
col = chr(ord('@')+col_num)
temp = sheet_temp.range(("A2:{}19").format(col)).value
sheet_temp.range(("A2:{}18").format(col)).value = ""
sheet_temp.range("A2").value = roles
sheet_temp.range("A3").value = temp
data =  sheet_temp.range(("A1:{}20").format(col)).value
sheet_temp.range(("A1:Z100").format(col)).value = ""
sheet_master.range("B4").value = data
wb1.save("D://Users//kevinspang//Documents//Excel//Accounts Reports//Joe Ditullio//Accounts_Report_Joe_DiTullio_{}.xlsx".format(today_save))
wb1.close()
wb2.close()
app.kill()