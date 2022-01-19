# -*- coding: utf-8 -*-
"""
Created on Tue Nov 19 14:48:39 2019

@author: kevinspang
"""

import gspread 
import datetime as dt
import mysql.connector
import xlwings as xw
import sys
import pandas as pd
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
director_data = account_data.get_worksheet(2)
manager_data = account_data.get_worksheet(3)
team1_data = account_data.get_worksheet(4)
team2_data = account_data.get_worksheet(5)


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

top_date = quartername + ": " + today_infile
remaining = "Days Remaining in Quarter: " + str(daysRemaining)


##################################################################################################################################

def Convert(string):
    li = list(string.split(","))
    return li

#team 1 = ELi
#team 2 = Jackie
#team 3 = Max

total_team_ids = accounts.acell('B37').value
team1_ids = accounts.acell('E37').value
team2_ids = accounts.acell('G37').value
unsupported = accounts.acell('E41').value
supported = accounts.acell('F41').value

#remove support from total list
try:
    total_team_ids = total_team_ids.replace("'730',","")
except:
    pass

#convert team ids to string for total team queries later
total_team_list = str(total_team_ids)
team1_list = str(team1_ids)
team2_list = str(team2_ids)

#add results to list
total_query_result = []

#######################################################################################################################################
#TEAM order: team1, team2, total

print("Running Eli's Team")
c = cnx.cursor()
query_team1 = ('''
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
and deal_stage=4) a'''.format(team1_list,team1_list,team1_list,team1_list,team1_list,team1_list,team1_list,unsupported,team1_list,supported,
                        team1_list,team1_list,team1_list,team1_list,team1_list,team1_list,team1_list,team1_list,team1_list,team1_list))
c.execute(query_team1)
team_result1 = [item[0] for item in c.fetchall()]
total_query_result.append(team_result1)

#######################################################################################################################################
#TEAM order: team1, team2, total

print("Running Jackies's Team")
c = cnx.cursor()
query_team2 = ('''
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
and deal_stage=4) a'''.format(team2_list,team2_list,team2_list,team2_list,team2_list,team2_list,team2_list,unsupported,team2_list,supported,
                        team2_list,team2_list,team2_list,team2_list,team2_list,team2_list,team2_list,team2_list,team2_list,team2_list))
c.execute(query_team2)
team_result2 = [item[0] for item in c.fetchall()]
total_query_result.append(team_result2)

###############################################################################################################################################################
#TEAM order: team1, team2, total

print("Running Total Team")
c = cnx.cursor()
query_team_total = ('''
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
and deal_stage=4) a'''.format(total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,unsupported,total_team_list,supported,
                              total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,total_team_list,total_team_list))
c.execute(query_team_total)
team_result_total = [item[0] for item in c.fetchall()]
total_query_result.append(team_result_total)
cnx.close()
###############################################################################################################################################################

#Open up sheets to grab invidivual results
print("Grabbing Individual Team Data")
app = xw.App(visible=False)

#get workbook names
eli_wb_name = "D://Users//kevinspang//Documents//Excel//Accounts Reports//eli moger//Accounts_Report_Eli_Moger_{}.xlsx".format(today_save)
jackie_wb_name = "D://Users//kevinspang//Documents//Excel//Accounts Reports//jackie ellingson//Accounts_Report_Jackie_Ellingson_{}.xlsx".format(today_save)
hannah_wb_name = "D://Users//kevinspang//Documents//Excel//Accounts Reports//hannah lucas//Accounts_Report_Hannah_Lucas_{}.xlsx".format(today_save)
sami_wb_name = "D://Users//kevinspang//Documents//Excel//Accounts Reports//sami kaminski//Accounts_Report_Sami_Kaminski_{}.xlsx".format(today_save)
joe_wb_name = "D://Users//kevinspang//Documents//Excel//Accounts Reports//joe ditullio//Accounts_Report_Joe_Ditullio_{}.xlsx".format(today_save)
will_wb_name = "D://Users//kevinspang//Documents//Excel//Accounts Reports//will rice//Accounts_Report_Will_Rice_{}.xlsx".format(today_save)


try:
    wbeli = xw.Book(eli_wb_name)
    eli_sheet = wbeli.sheets['Master']
    eli_sheet_data = eli_sheet.range("B4:R23").value
    wbeli.close()
except:
    print("Eli's sheet not yet run")
    sys.exit()

try:
    wbjackie = xw.Book(jackie_wb_name)
    jackie_sheet = wbjackie.sheets['Master']
    jackie_sheet_data = jackie_sheet.range("B4:R23").value
    wbjackie.close()
except:
    print("Jackie's sheet not yet run")
    sys.exit()
    
try:
    wbhannah = xw.Book(hannah_wb_name)
    hannah_sheet = wbhannah.sheets['Master']
    hannah_sheet_data = hannah_sheet.range("B4:R23").value
    wbhannah.close()
except:
    print("Hannah's sheet not yet run")
    sys.exit()    

try:
    wbsami = xw.Book(sami_wb_name)
    sami_sheet = wbsami.sheets['Master']
    sami_sheet_data = sami_sheet.range("B4:R23").value
    wbsami.close()
except:
    print("Sami's sheet not yet run")
    sys.exit()
    
try:
    wbjoe = xw.Book(joe_wb_name)
    joe_sheet = wbjoe.sheets['Master']
    joe_sheet_data = joe_sheet.range("B4:R23").value
    wbjoe.close()
except:
    print("Joe's sheet not yet run")
    sys.exit()

try:
    wbwill = xw.Book(will_wb_name)
    will_sheet = wbwill.sheets['Master']
    will_sheet_data = will_sheet.range("B4:R23").value
    wbwill.close()
except:
    print("Will's sheet not yet run")
    sys.exit()

####################################################################################################################################################################
#grab team count data
print("Grabbing Team Count Data")    

director_data = director_data.get_all_values()
director_headers = director_data.pop(0)
df_director_count = pd.DataFrame(director_data, columns = director_headers)

manager_data = manager_data.get_all_values()
manager_headers = manager_data.pop(0)
df_manager_count = pd.DataFrame(manager_data, columns = manager_headers)

team1_data_grab = team1_data.get_all_values()
team1_headers = team1_data_grab.pop(0)
df_team1 = pd.DataFrame(team1_data_grab, columns = team1_headers)

team2_data_grab = team2_data.get_all_values()
team2_headers = team2_data_grab.pop(0)
df_team2 = pd.DataFrame(team2_data_grab, columns = team2_headers)


####################################################################################################################################################################
#input into template
print("Inputting Data into Master Sheet")

wbmaster = xw.Book("D://Users//kevinspang//Documents//Excel//Accounts Reports//master//template_master.xlsx")
master_sheet = wbmaster.sheets['Master']
eli_sheet = wbmaster.sheets['Eli Team']
jackie_sheet = wbmaster.sheets['Jackie Team']
hannah_sheet = wbmaster.sheets['Hannah Team']
sami_sheet = wbmaster.sheets['Sami Team']
joe_sheet = wbmaster.sheets['Joe Team']
will_sheet = wbmaster.sheets['Will Team']

eli_sheet.range("A1").value = top_date
eli_sheet.range("A2").value = remaining
eli_sheet.range("B4").value = eli_sheet_data

jackie_sheet.range("A1").value = top_date
jackie_sheet.range("A2").value = remaining
jackie_sheet.range("B4").value = jackie_sheet_data

hannah_sheet.range("A1").value = top_date
hannah_sheet.range("A2").value = remaining
hannah_sheet.range("B4").value = hannah_sheet_data

sami_sheet.range("A1").value = top_date
sami_sheet.range("A2").value = remaining
sami_sheet.range("B4").value = sami_sheet_data

joe_sheet.range("A1").value = top_date
joe_sheet.range("A2").value = remaining
joe_sheet.range("B4").value = joe_sheet_data

will_sheet.range("A1").value = top_date
will_sheet.range("A2").value = remaining
will_sheet.range("B4").value = will_sheet_data

master_sheet.range("A1").value = top_date
master_sheet.range("A2").value = remaining
master_sheet.range("B5").options(transpose=True).value = total_query_result
master_sheet.range("A22").options(index = False).value = df_director_count
master_sheet.range("A27").options(index = False).value = df_manager_count
master_sheet.range("A34").options(index = False).value = df_team1
master_sheet.range("A49").options(index = False).value = df_team2

wbmaster.save("D://Users//kevinspang//Documents//Excel//Accounts Reports//master//Accounts_Report_Master_{}.xlsx".format(today_save))
wbmaster.close()
app.kill()