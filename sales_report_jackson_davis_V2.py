# -*- coding: utf-8 -*-
"""
Created on Tue Feb  4 15:01:49 2020

@author: kevinspang
"""

import mysql.connector
import xlwings as xw
import datetime as dt
import sys
import os.path
import time
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

##############################################################################################################################
#see if todays sheet exists if it does cancel sheet creation to avoid messing up template number    

# if statement to only run on Mondays
if dt.date.today().isoweekday() == 1: #change to one for mondays
    #if monday check if todays sheet has already run.
    file_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    file_path = 'D://users//kevinspang//my documents//excel//sales reports//jackson davis//sales_report_jackson_davis_{}.xlsx'.format(file_date) 
    if os.path.isfile(file_path) == True:
        print("Todays Sheet Already Complete")
        sys.exit(1)
    else:
        print("Running Jackson Davis's sales report")
else:
    print("Report only runs on Mondays")
    sys.exit(1)

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
year_start = year + '-01-01'
year_end = year + '-12-31'

#initialize date logic and add in goal data
app = xw.App(visible=False)  
wbgoals = xw.Book('D://users//kevinspang//my documents//excel//sales reports//Goals.xlsx' )
goals = wbgoals.sheets['Goals']

if q1start <= today <= q1end:
    quartername = 'Q1'
    quarterstart = year + '-01-01'
    quarterend = year + '-03-31'
    last_monday = year + '-03-29'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team2 = float(goals.range('C21').value)
    qgoal_1414 = float(goals.range('C11').value) #Andrew Piekos
    qgoal_1837 = float(goals.range('C12').value) #Joe Paratore
    qgoal_2907 = float(goals.range('C13').value) #Morgan Afshari
    qgoal_3099 = float(goals.range('C14').value) #Mike Shea

elif q2start <= today <= q2end:
    quartername = 'Q2'
    quarterstart = year + '-04-01'
    quarterend = year + '-06-30'
    last_monday = year + '-06-28'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team2 = float(goals.range('D21').value)
    qgoal_1414 = float(goals.range('D11').value) #Andrew Piekos
    qgoal_1837 = float(goals.range('D12').value) #Joe Paratore
    qgoal_2907 = float(goals.range('D13').value) #Morgan Afshari
    qgoal_3099 = float(goals.range('D14').value) #Mike Shea

elif q3start <= today <= q3end:
    quartername = 'Q3'
    quarterstart = year + '-07-01'
    quarterend = year + '-09-30'
    last_monday = year + '-09-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team2 = float(goals.range('E21').value)
    qgoal_1414 = float(goals.range('E11').value) #Andrew Piekos
    qgoal_1837 = float(goals.range('E12').value) #Joe Paratore
    qgoal_2907 = float(goals.range('E13').value) #Morgan Afshari
    qgoal_3099 = float(goals.range('E14').value) #Mike Shea

else:
    quartername = 'Q4'
    quarterstart = year + '-10-01'
    quarterend = year + '-12-31' 
    last_monday = year + '-12-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team2 = float(goals.range('F21').value)
    qgoal_1414 = float(goals.range('F11').value) #Andrew Piekos
    qgoal_1837 = float(goals.range('F12').value) #Joe Paratore
    qgoal_2907 = float(goals.range('F13').value) #Morgan Afshari
    qgoal_3099 = float(goals.range('F14').value) #Mike Shea


wbgoals.close()
################################################################################################################
# open template and grab last weeks closed won and rfp numbers   
wb1 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//jackson davis//sales_report_jackson_davis_template.xlsx' )
    
sheet_team2 = wb1.sheets['Team']
sheet_1414 = wb1.sheets['Andrew Piekos']
sheet_1837 = wb1.sheets['Joe Paratore']
sheet_2907 = wb1.sheets['Morgan Afshari']
sheet_3099 = wb1.sheets['Mike Shea']

#grab last weeks closed won rfp and expired values
last_week_cw_team2 = float(sheet_team2.range('B11').value)
sheet_team2.range('B13').value = last_week_cw_team2
last_week_open_team2 = float(sheet_team2.range('B17').value)
last_week_expired_team2 = float(sheet_team2.range('B19').value)

last_week_cw_1414 = float(sheet_1414.range('B11').value)
sheet_1414.range('B13').value = last_week_cw_1414
last_week_open_1414 = float(sheet_1414.range('B17').value)
last_week_expired_1414 = float(sheet_1414.range('B19').value)

last_week_cw_1837 = float(sheet_1837.range('B11').value)
sheet_1837.range('B13').value = last_week_cw_1837
last_week_open_1837 = float(sheet_1837.range('B17').value)
last_week_expired_1837 = float(sheet_1837.range('B19').value)

last_week_cw_2907 = float(sheet_2907.range('B11').value)
sheet_2907.range('B13').value = last_week_cw_2907
last_week_open_2907 = float(sheet_2907.range('B17').value)
last_week_expired_2907 = float(sheet_2907.range('B19').value)

last_week_cw_3099 = float(sheet_3099.range('B11').value)
sheet_3099.range('B13').value = last_week_cw_3099
last_week_open_3099 = float(sheet_3099.range('B17').value)
last_week_expired_3099 = float(sheet_3099.range('B19').value)


##############################################################################################################################
#Team 2    
    
name_id = "'1402','1414','1837','2907','3099'"      

#Billed revenue from start of date to today
c = cnx.cursor()
c.execute('''select ifnull((SELECT 
    SUM(CASE
        WHEN c.type = 'agency' THEN round(lir.cpm * lir.impressions / 1000,2)
    END) AS revenueToDate
FROM
    line_item_report AS lir
        INNER JOIN
    client c ON lir.client_id = c.id
WHERE
    lir.report_date between '{}' and date_sub(sysdate(), interval 1 day)
    and lir.client_id in (
		Select distinct client_id 
        from sales_forecast 
        where salesPerson_id in ({}) 
        and start_date >= '{}')
GROUP BY date_format(lir.report_date, '%Y')),0) as billed_revenue '''.format(quarterstart,name_id,year_start))
billed_team2 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

#Closed Won for quarter
c.execute('''select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) revenue 
from (
Select split_percentage, salesForecast_id
from revenue_split
where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4'
and id in (select distinct salesForecast_id 
		   from sales_forecast_daily_budget
           where day between '{}' and '{}')
))a,
(select salesForecast_id, sum(budget) budget
from sales_forecast_daily_budget
where day between '{}' and '{}'
group by salesForecast_id) b
where a.salesForecast_id = b.salesForecast_id '''.format(name_id,quarterstart,quarterend,quarterstart,quarterend))
cw_team2 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_team2 = cw_team2 - last_week_cw_team2 #closed won change 

cw_percent_team2 = cw_team2 / qgoal_team2 #CW percent to goal

#total Remaining rfp
c.execute('''select ifnull((
select sum(revenue) from (
Select round(sum(budget)*4,2)  as revenue
from sales_forecast_daily_budget
where day between  '{}' and '{}'
and salesForecast_id in
(select distinct id from sales_forecast
where salesPerson_id in ({}) 
and deal_stage = 1)
union
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage = 2
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a ) b),0) as revenue ;'''.format(quarterstart,quarterend,name_id,name_id))
rfp_team2 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_team2 = rfp_team2 - last_week_open_team2 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_team2 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_team2 = rfp_expired_team2 - last_week_expired_team2 #expired deals change

#list of remaining rfp
c.execute('''Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
round(sum(db.budget)*4,2) rfp_revenue, sf.budget as rfp_total_revenue,
start_date, end_date,
'RFP' as deal_stage
from sales_forecast sf
join sales_forecast_daily_budget db
on sf.id = db.salesForecast_id
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage = 1
and db.day between '{}' and '{}'
and sf.salesPerson_id in ({}) 
group by sf.id 
union
Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
sf.budget as rfp_revenue, sf.budget as rfp_total_revenue, start_date, end_date,
case 
when sf.deal_stage = 2 then 'Negotiating'
when sf.deal_stage = 3 then 'Contract Sent'
else 'Expired' 
end deal_stage
from sales_forecast sf
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage in (2,3,6)
and start_date >= date_format(date_sub(sysdate(), interval 30 day), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id order by 10 desc,2,3;   '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_team2 = c.fetchall() #list of rfp, expired, negotiating, ect.

#Weekly closed won and pipeline for each quarter
c.execute('''select ifnull(rfp1.rfp + cw1.cw,0) pipelineq1, ifnull(cw1.cw,0) closedwonq1,
ifnull(rfp2.rfp + cw2.cw,0)  pipelineq2, ifnull(cw2.cw,0) closedwonq2,  
ifnull(rfp3.rfp + cw3.cw,0)  pipelineq3, ifnull(cw3.cw,0) closedwonq3,  
ifnull(rfp4.rfp + cw4.cw,0)  pipelineq4, ifnull(cw4.cw,0) closedwonq4,  
ifnull(rfpTotal.rfp + cwTotal.cw,0)  pipelineqTotal, ifnull(cwTotal.cw,0) closedwonqTotal from
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp1,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw1,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp2,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw2,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp3,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw3,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp4,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw4,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfpTotal,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cwTotal
'''.format(q1start,q1end,name_id,name_id,q1start,q1end,q1start,q1end,
           q2start,q2end,name_id,name_id,q2start,q2end,q2start,q2end,
           q3start,q3end,name_id,name_id,q3start,q3end,q3start,q3end,
           q4start,q4end,name_id,name_id,q4start,q4end,q4start,q4end,
           year_start,year_end,name_id,name_id,year_start,year_end,year_start,year_end))

weekly_sheet_team2 = c.fetchall()

ideal_pacing_team2 = qgoal_team2/daysInQuarter * daysPast #Ideal Pace number

differential_team2 = (billed_team2 - ideal_pacing_team2) / ideal_pacing_team2 #pacing differential

rfp_needed_team2 = qgoal_team2 - cw_team2 #amount of rfp needed for quarter

try:
    rfp_needed_percent_team2 = rfp_needed_team2/rfp_team2 #percent of open deals needed
except:
    rfp_needed_percent_team2 = 'N/A'
################################################################################################################
#Andrew Piekos  
    
name_id = '1414'    

#Billed revenue from start of date to today
c = cnx.cursor()
c.execute('''select ifnull((SELECT 
    SUM(CASE
        WHEN c.type = 'agency' THEN round(lir.cpm * lir.impressions / 1000,2)
    END) AS revenueToDate
FROM
    line_item_report AS lir
        INNER JOIN
    client c ON lir.client_id = c.id
WHERE
    lir.report_date between '{}' and date_sub(sysdate(), interval 1 day)
    and lir.client_id in (
		Select distinct client_id 
        from sales_forecast 
        where salesPerson_id in ({}) 
        and start_date >= '{}')
GROUP BY date_format(lir.report_date, '%Y')),0) as billed_revenue '''.format(quarterstart,name_id,year_start))
billed_1414 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

#Closed Won for quarter
c.execute('''select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) revenue 
from (
Select split_percentage, salesForecast_id
from revenue_split
where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4'
and id in (select distinct salesForecast_id 
		   from sales_forecast_daily_budget
           where day between '{}' and '{}')
))a,
(select salesForecast_id, sum(budget) budget
from sales_forecast_daily_budget
where day between '{}' and '{}'
group by salesForecast_id) b
where a.salesForecast_id = b.salesForecast_id '''.format(name_id,quarterstart,quarterend,quarterstart,quarterend))
cw_1414 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_1414 = cw_1414 - last_week_cw_1414 #closed won change 

cw_percent_1414 = cw_1414 / qgoal_1414 #CW percent to goal

#total Remaining rfp
c.execute('''select ifnull((
select sum(revenue) from (
Select round(sum(budget)*4,2)  as revenue
from sales_forecast_daily_budget
where day between  '{}' and '{}'
and salesForecast_id in
(select distinct id from sales_forecast
where salesPerson_id in ({}) 
and deal_stage = 1)
union
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage = 2
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a ) b),0) as revenue ;'''.format(quarterstart,quarterend,name_id,name_id))
rfp_1414 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_1414 = rfp_1414 - last_week_open_1414 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_1414 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_1414 = rfp_expired_1414 - last_week_expired_1414 #expired deals change

#list of remaining rfp
c.execute('''Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
round(sum(db.budget)*4,2) rfp_revenue, sf.budget as rfp_total_revenue,
start_date, end_date,
'RFP' as deal_stage
from sales_forecast sf
join sales_forecast_daily_budget db
on sf.id = db.salesForecast_id
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage = 1
and db.day between '{}' and '{}'
and sf.salesPerson_id in ({}) 
group by sf.id 
union
Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
sf.budget as rfp_revenue, sf.budget as rfp_total_revenue, start_date, end_date,
case 
when sf.deal_stage = 2 then 'Negotiating'
when sf.deal_stage = 3 then 'Contract Sent'
else 'Expired' 
end deal_stage
from sales_forecast sf
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage in (2,3,6)
and start_date >= date_format(date_sub(sysdate(), interval 30 day), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id order by 10 desc,2,3;   '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_1414 = c.fetchall() #list of rfp, expired, negotiating, ect.

#Weekly closed won and pipeline for each quarter
c.execute('''select ifnull(rfp1.rfp + cw1.cw,0) pipelineq1, ifnull(cw1.cw,0) closedwonq1,
ifnull(rfp2.rfp + cw2.cw,0)  pipelineq2, ifnull(cw2.cw,0) closedwonq2,  
ifnull(rfp3.rfp + cw3.cw,0)  pipelineq3, ifnull(cw3.cw,0) closedwonq3,  
ifnull(rfp4.rfp + cw4.cw,0)  pipelineq4, ifnull(cw4.cw,0) closedwonq4,  
ifnull(rfpTotal.rfp + cwTotal.cw,0)  pipelineqTotal, ifnull(cwTotal.cw,0) closedwonqTotal from
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp1,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw1,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp2,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw2,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp3,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw3,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp4,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw4,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfpTotal,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cwTotal
'''.format(q1start,q1end,name_id,name_id,q1start,q1end,q1start,q1end,
           q2start,q2end,name_id,name_id,q2start,q2end,q2start,q2end,
           q3start,q3end,name_id,name_id,q3start,q3end,q3start,q3end,
           q4start,q4end,name_id,name_id,q4start,q4end,q4start,q4end,
           year_start,year_end,name_id,name_id,year_start,year_end,year_start,year_end))

weekly_sheet_1414 = c.fetchall()

ideal_pacing_1414 = qgoal_1414/daysInQuarter * daysPast #Ideal Pace number

differential_1414 = (billed_1414 - ideal_pacing_1414) / ideal_pacing_1414 #pacing differential

rfp_needed_1414 = qgoal_1414 - cw_1414 #amount of rfp needed for quarter

try:
    rfp_needed_percent_1414 = rfp_needed_1414/rfp_1414 #percent of open deals needed
except:
    rfp_needed_percent_1414 = 'N/A'
################################################################################################################
#Joe Paratore
    
name_id = '1837'    

#Billed revenue from start of date to today
c = cnx.cursor()
c.execute('''select ifnull((SELECT 
    SUM(CASE
        WHEN c.type = 'agency' THEN round(lir.cpm * lir.impressions / 1000,2)
    END) AS revenueToDate
FROM
    line_item_report AS lir
        INNER JOIN
    client c ON lir.client_id = c.id
WHERE
    lir.report_date between '{}' and date_sub(sysdate(), interval 1 day)
    and lir.client_id in (
		Select distinct client_id 
        from sales_forecast 
        where salesPerson_id in ({}) 
        and start_date >= '{}')
GROUP BY date_format(lir.report_date, '%Y')),0) as billed_revenue '''.format(quarterstart,name_id,year_start))
billed_1837 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

#Closed Won for quarter
c.execute('''select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) revenue 
from (
Select split_percentage, salesForecast_id
from revenue_split
where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4'
and id in (select distinct salesForecast_id 
		   from sales_forecast_daily_budget
           where day between '{}' and '{}')
))a,
(select salesForecast_id, sum(budget) budget
from sales_forecast_daily_budget
where day between '{}' and '{}'
group by salesForecast_id) b
where a.salesForecast_id = b.salesForecast_id '''.format(name_id,quarterstart,quarterend,quarterstart,quarterend))
cw_1837 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_1837 = cw_1837 - last_week_cw_1837 #closed won change 

cw_percent_1837 = cw_1837 / qgoal_1837 #CW percent to goal

#total Remaining rfp
c.execute('''select ifnull((
select sum(revenue) from (
Select round(sum(budget)*4,2)  as revenue
from sales_forecast_daily_budget
where day between  '{}' and '{}'
and salesForecast_id in
(select distinct id from sales_forecast
where salesPerson_id in ({}) 
and deal_stage = 1)
union
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage = 2
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a ) b),0) as revenue ;'''.format(quarterstart,quarterend,name_id,name_id))
rfp_1837 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_1837 = rfp_1837 - last_week_open_1837 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_1837 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_1837 = rfp_expired_1837 - last_week_expired_1837 #expired deals change

#list of remaining rfp
c.execute('''Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
round(sum(db.budget)*4,2) rfp_revenue, sf.budget as rfp_total_revenue,
start_date, end_date,
'RFP' as deal_stage
from sales_forecast sf
join sales_forecast_daily_budget db
on sf.id = db.salesForecast_id
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage = 1
and db.day between '{}' and '{}'
and sf.salesPerson_id in ({}) 
group by sf.id 
union
Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
sf.budget as rfp_revenue, sf.budget as rfp_total_revenue, start_date, end_date,
case 
when sf.deal_stage = 2 then 'Negotiating'
when sf.deal_stage = 3 then 'Contract Sent'
else 'Expired' 
end deal_stage
from sales_forecast sf
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage in (2,3,6)
and start_date >= date_format(date_sub(sysdate(), interval 30 day), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id order by 10 desc,2,3;   '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_1837 = c.fetchall() #list of rfp, expired, negotiating, ect.

#Weekly closed won and pipeline for each quarter
c.execute('''select ifnull(rfp1.rfp + cw1.cw,0) pipelineq1, ifnull(cw1.cw,0) closedwonq1,
ifnull(rfp2.rfp + cw2.cw,0)  pipelineq2, ifnull(cw2.cw,0) closedwonq2,  
ifnull(rfp3.rfp + cw3.cw,0)  pipelineq3, ifnull(cw3.cw,0) closedwonq3,  
ifnull(rfp4.rfp + cw4.cw,0)  pipelineq4, ifnull(cw4.cw,0) closedwonq4,  
ifnull(rfpTotal.rfp + cwTotal.cw,0)  pipelineqTotal, ifnull(cwTotal.cw,0) closedwonqTotal from
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp1,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw1,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp2,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw2,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp3,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw3,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp4,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw4,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfpTotal,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cwTotal
'''.format(q1start,q1end,name_id,name_id,q1start,q1end,q1start,q1end,
           q2start,q2end,name_id,name_id,q2start,q2end,q2start,q2end,
           q3start,q3end,name_id,name_id,q3start,q3end,q3start,q3end,
           q4start,q4end,name_id,name_id,q4start,q4end,q4start,q4end,
           year_start,year_end,name_id,name_id,year_start,year_end,year_start,year_end))

weekly_sheet_1837 = c.fetchall()

ideal_pacing_1837 = qgoal_1837/daysInQuarter * daysPast #Ideal Pace number

differential_1837 = (billed_1837 - ideal_pacing_1837) / ideal_pacing_1837 #pacing differential

rfp_needed_1837 = qgoal_1837 - cw_1837 #amount of rfp needed for quarter

try:
    rfp_needed_percent_1837 = rfp_needed_1837/rfp_1837 #percent of open deals needed
except:
    rfp_needed_percent_1837 = 'N/A'
################################################################################################################    
#Morgan Afshari
    
name_id = '2907'    

#Billed revenue from start of date to today
c = cnx.cursor()
c.execute('''select ifnull((SELECT 
    SUM(CASE
        WHEN c.type = 'agency' THEN round(lir.cpm * lir.impressions / 1000,2)
    END) AS revenueToDate
FROM
    line_item_report AS lir
        INNER JOIN
    client c ON lir.client_id = c.id
WHERE
    lir.report_date between '{}' and date_sub(sysdate(), interval 1 day)
    and lir.client_id in (
		Select distinct client_id 
        from sales_forecast 
        where salesPerson_id in ({}) 
        and start_date >= '{}')
GROUP BY date_format(lir.report_date, '%Y')),0) as billed_revenue '''.format(quarterstart,name_id,year_start))
billed_2907 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

#Closed Won for quarter
c.execute('''select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) revenue 
from (
Select split_percentage, salesForecast_id
from revenue_split
where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4'
and id in (select distinct salesForecast_id 
		   from sales_forecast_daily_budget
           where day between '{}' and '{}')
))a,
(select salesForecast_id, sum(budget) budget
from sales_forecast_daily_budget
where day between '{}' and '{}'
group by salesForecast_id) b
where a.salesForecast_id = b.salesForecast_id '''.format(name_id,quarterstart,quarterend,quarterstart,quarterend))
cw_2907 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_2907 = cw_2907 - last_week_cw_2907 #closed won change 

cw_percent_2907 = cw_2907 / qgoal_2907 #CW percent to goal

#total Remaining rfp
c.execute('''select ifnull((
select sum(revenue) from (
Select round(sum(budget)*4,2)  as revenue
from sales_forecast_daily_budget
where day between  '{}' and '{}'
and salesForecast_id in
(select distinct id from sales_forecast
where salesPerson_id in ({}) 
and deal_stage = 1)
union
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage = 2
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a ) b),0) as revenue ;'''.format(quarterstart,quarterend,name_id,name_id))
rfp_2907 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_2907 = rfp_2907 - last_week_open_2907 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_2907 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_2907 = rfp_expired_2907 - last_week_expired_2907 #expired deals change

#list of remaining rfp
c.execute('''Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
round(sum(db.budget)*4,2) rfp_revenue, sf.budget as rfp_total_revenue,
start_date, end_date,
'RFP' as deal_stage
from sales_forecast sf
join sales_forecast_daily_budget db
on sf.id = db.salesForecast_id
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage = 1
and db.day between '{}' and '{}'
and sf.salesPerson_id in ({}) 
group by sf.id 
union
Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
sf.budget as rfp_revenue, sf.budget as rfp_total_revenue, start_date, end_date,
case 
when sf.deal_stage = 2 then 'Negotiating'
when sf.deal_stage = 3 then 'Contract Sent'
else 'Expired' 
end deal_stage
from sales_forecast sf
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage in (2,3,6)
and start_date >= date_format(date_sub(sysdate(), interval 30 day), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id order by 10 desc,2,3;   '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_2907 = c.fetchall() #list of rfp, expired, negotiating, ect.

#Weekly closed won and pipeline for each quarter
c.execute('''select ifnull(rfp1.rfp + cw1.cw,0) pipelineq1, ifnull(cw1.cw,0) closedwonq1,
ifnull(rfp2.rfp + cw2.cw,0)  pipelineq2, ifnull(cw2.cw,0) closedwonq2,  
ifnull(rfp3.rfp + cw3.cw,0)  pipelineq3, ifnull(cw3.cw,0) closedwonq3,  
ifnull(rfp4.rfp + cw4.cw,0)  pipelineq4, ifnull(cw4.cw,0) closedwonq4,  
ifnull(rfpTotal.rfp + cwTotal.cw,0)  pipelineqTotal, ifnull(cwTotal.cw,0) closedwonqTotal from
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp1,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw1,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp2,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw2,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp3,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw3,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp4,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw4,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfpTotal,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cwTotal
'''.format(q1start,q1end,name_id,name_id,q1start,q1end,q1start,q1end,
           q2start,q2end,name_id,name_id,q2start,q2end,q2start,q2end,
           q3start,q3end,name_id,name_id,q3start,q3end,q3start,q3end,
           q4start,q4end,name_id,name_id,q4start,q4end,q4start,q4end,
           year_start,year_end,name_id,name_id,year_start,year_end,year_start,year_end))

weekly_sheet_2907 = c.fetchall()

ideal_pacing_2907 = qgoal_2907/daysInQuarter * daysPast #Ideal Pace number

differential_2907 = (billed_2907 - ideal_pacing_2907) / ideal_pacing_2907 #pacing differential

rfp_needed_2907 = qgoal_2907 - cw_2907 #amount of rfp needed for quarter

try:
    rfp_needed_percent_2907 = rfp_needed_2907/rfp_2907 #percent of open deals needed
except:
    rfp_needed_percent_2907 = 'N/A'
#################################################################################################################################  
#Mike Shea
    
name_id = '3099'    

#Billed revenue from start of date to today
c = cnx.cursor()
c.execute('''select ifnull((SELECT 
    SUM(CASE
        WHEN c.type = 'agency' THEN round(lir.cpm * lir.impressions / 1000,2)
    END) AS revenueToDate
FROM
    line_item_report AS lir
        INNER JOIN
    client c ON lir.client_id = c.id
WHERE
    lir.report_date between '{}' and date_sub(sysdate(), interval 1 day)
    and lir.client_id in (
		Select distinct client_id 
        from sales_forecast 
        where salesPerson_id in ({}) 
        and start_date >= '{}')
GROUP BY date_format(lir.report_date, '%Y')),0) as billed_revenue '''.format(quarterstart,name_id,year_start))
billed_3099 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

#Closed Won for quarter
c.execute('''select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) revenue 
from (
Select split_percentage, salesForecast_id
from revenue_split
where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4'
and id in (select distinct salesForecast_id 
		   from sales_forecast_daily_budget
           where day between '{}' and '{}')
))a,
(select salesForecast_id, sum(budget) budget
from sales_forecast_daily_budget
where day between '{}' and '{}'
group by salesForecast_id) b
where a.salesForecast_id = b.salesForecast_id '''.format(name_id,quarterstart,quarterend,quarterstart,quarterend))
cw_3099 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_3099 = cw_3099 - last_week_cw_3099 #closed won change 

cw_percent_3099 = cw_3099 / qgoal_3099 #CW percent to goal

#total Remaining rfp
c.execute('''select ifnull((
select sum(revenue) from (
Select round(sum(budget)*4,2)  as revenue
from sales_forecast_daily_budget
where day between  '{}' and '{}'
and salesForecast_id in
(select distinct id from sales_forecast
where salesPerson_id in ({}) 
and deal_stage = 1)
union
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage = 2
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a ) b),0) as revenue ;'''.format(quarterstart,quarterend,name_id,name_id))
rfp_3099 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_3099 = rfp_3099 - last_week_open_3099 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_3099 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_3099 = rfp_expired_3099 - last_week_expired_3099 #expired deals change

#list of remaining rfp
c.execute('''Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
round(sum(db.budget)*4,2) rfp_revenue, sf.budget as rfp_total_revenue,
start_date, end_date,
'RFP' as deal_stage
from sales_forecast sf
join sales_forecast_daily_budget db
on sf.id = db.salesForecast_id
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage = 1
and db.day between '{}' and '{}'
and sf.salesPerson_id in ({}) 
group by sf.id 
union
Select sf.id, concat(fos.first_name, " ", fos.last_name) rep,
cl.name, sf.advertiser_name, sf.campaign_name,
sf.budget as rfp_revenue, sf.budget as rfp_total_revenue, start_date, end_date,
case 
when sf.deal_stage = 2 then 'Negotiating'
when sf.deal_stage = 3 then 'Contract Sent'
else 'Expired' 
end deal_stage
from sales_forecast sf
join fos_user fos
on fos.id = sf.salesPerson_id
join client cl
on cl.id = sf.client_id
where sf.deal_stage in (2,3,6)
and start_date >= date_format(date_sub(sysdate(), interval 30 day), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id order by 10 desc,2,3;   '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_3099 = c.fetchall() #list of rfp, expired, negotiating, ect.

#Weekly closed won and pipeline for each quarter
c.execute('''select ifnull(rfp1.rfp + cw1.cw,0) pipelineq1, ifnull(cw1.cw,0) closedwonq1,
ifnull(rfp2.rfp + cw2.cw,0)  pipelineq2, ifnull(cw2.cw,0) closedwonq2,  
ifnull(rfp3.rfp + cw3.cw,0)  pipelineq3, ifnull(cw3.cw,0) closedwonq3,  
ifnull(rfp4.rfp + cw4.cw,0)  pipelineq4, ifnull(cw4.cw,0) closedwonq4,  
ifnull(rfpTotal.rfp + cwTotal.cw,0)  pipelineqTotal, ifnull(cwTotal.cw,0) closedwonqTotal from
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp1,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw1,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp2,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw2,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp3,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw3,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfp4,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cw4,
(Select round(sum(budget),2) rfp from sales_forecast_daily_budget
where day between '{}' and '{}' 
and salesForecast_id in (select id from sales_forecast 
where salesPerson_id in ({}) and deal_stage = '1')) rfpTotal,
(select ifnull(round(sum(a.split_percentage / 100 * b.budget),2),0) cw 
from (Select split_percentage, salesForecast_id
from revenue_split where salesPerson_id in ({})
and salesForecast_id in (select id from sales_forecast 
where deal_stage = '4' and id in (select distinct salesForecast_id 
from sales_forecast_daily_budget where day between '{}' and '{}' )))a,
(select salesForecast_id, sum(budget) budget from sales_forecast_daily_budget
where day between '{}' and '{}' 
group by salesForecast_id) b where a.salesForecast_id = b.salesForecast_id) cwTotal
'''.format(q1start,q1end,name_id,name_id,q1start,q1end,q1start,q1end,
           q2start,q2end,name_id,name_id,q2start,q2end,q2start,q2end,
           q3start,q3end,name_id,name_id,q3start,q3end,q3start,q3end,
           q4start,q4end,name_id,name_id,q4start,q4end,q4start,q4end,
           year_start,year_end,name_id,name_id,year_start,year_end,year_start,year_end))

weekly_sheet_3099 = c.fetchall()

ideal_pacing_3099 = qgoal_3099/daysInQuarter * daysPast #Ideal Pace number

differential_3099 = (billed_3099 - ideal_pacing_3099) / ideal_pacing_3099 #pacing differential

rfp_needed_3099 = qgoal_3099 - cw_3099 #amount of rfp needed for quarter

try:
    rfp_needed_percent_3099 = rfp_needed_3099/rfp_3099 #percent of open deals needed
except:
    rfp_needed_percent_3099 = 'N/A'
#################################################################################################################################     
#insert new values into template and save

sheet_team2.range('A1').value = today
sheet_team2.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_team2.range('B11').value = cw_team2
sheet_team2.range('B17').value = rfp_team2
sheet_team2.range('B19').value = rfp_expired_team2   

sheet_1414.range('A1').value = today
sheet_1414.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_1414.range('B11').value = cw_1414
sheet_1414.range('B17').value = rfp_1414
sheet_1414.range('B19').value = rfp_expired_1414  

sheet_1837.range('A1').value = today
sheet_1837.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_1837.range('B11').value = cw_1837
sheet_1837.range('B17').value = rfp_1837
sheet_1837.range('B19').value = rfp_expired_1837  

sheet_2907.range('A1').value = today
sheet_2907.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_2907.range('B11').value = cw_2907
sheet_2907.range('B17').value = rfp_2907
sheet_2907.range('B19').value = rfp_expired_2907  

sheet_3099.range('A1').value = today
sheet_3099.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_3099.range('B11').value = cw_3099
sheet_3099.range('B17').value = rfp_3099
sheet_3099.range('B19').value = rfp_expired_3099  

#save template for next week
wb1.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//Sales_Report_Jackson_Davis_Template.xlsx')

#################################################################################################################################
#load in salesforce data

wb2 = xw.Book('D://users//kevinspang//documents//excel//sales reports//salesforce.xlsx' )
salesforce = wb2.sheets['Salesforce']

meetings_team2 = salesforce.range('C21').value
meetings_total_team2 = salesforce.range('D21').value
emails_team2 = salesforce.range('E21').value
emails_total_team2 = salesforce.range('F21').value

meetings_1414 = salesforce.range('C11').value
meetings_total_1414 = salesforce.range('D11').value
emails_1414 = salesforce.range('E11').value
emails_total_1414 = salesforce.range('F11').value

meetings_1837 = salesforce.range('C12').value
meetings_total_1837 = salesforce.range('D12').value
emails_1837 = salesforce.range('E12').value
emails_total_1837 = salesforce.range('F12').value

meetings_2907 = salesforce.range('C13').value
meetings_total_2907 = salesforce.range('D13').value
emails_2907 = salesforce.range('E13').value
emails_total_2907 = salesforce.range('F13').value

meetings_3099 = salesforce.range('C14').value
meetings_total_3099 = salesforce.range('D14').value
emails_3099 = salesforce.range('E14').value
emails_total_3099 = salesforce.range('F14').value

wb2.close()
#################################################################################################################################
#fill out remainder of sheet for distribution   
 
#Team 2 sheet
sheet_team2.range('B5').value = qgoal_team2
sheet_team2.range('B7').value = billed_team2
sheet_team2.range('B8').value = ideal_pacing_team2
sheet_team2.range('B9').value = differential_team2
sheet_team2.range('B12').value = cw_percent_team2
sheet_team2.range('B13').value = last_week_cw_team2
sheet_team2.range('B14').value = cw_change_team2
sheet_team2.range('B16').value = rfp_needed_team2
sheet_team2.range('B18').value = rfp_change_team2
sheet_team2.range('B20').value = expired_change_team2
sheet_team2.range('B21').value = rfp_needed_percent_team2
sheet_team2.range('B23').value = meetings_team2
sheet_team2.range('B24').value = emails_team2
sheet_team2.range('B25').value = meetings_total_team2
sheet_team2.range('B26').value = emails_total_team2
sheet_team2.range('A29:H200').value = ""
sheet_team2.range('A29').value = rfp_list_team2

#Andrew Piekos
sheet_1414.range('B5').value = qgoal_1414
sheet_1414.range('B7').value = billed_1414
sheet_1414.range('B8').value = ideal_pacing_1414
sheet_1414.range('B9').value = differential_1414
sheet_1414.range('B12').value = cw_percent_1414
sheet_1414.range('B13').value = last_week_cw_1414
sheet_1414.range('B14').value = cw_change_1414
sheet_1414.range('B16').value = rfp_needed_1414
sheet_1414.range('B18').value = rfp_change_1414
sheet_1414.range('B20').value = expired_change_1414
sheet_1414.range('B21').value = rfp_needed_percent_1414
sheet_1414.range('B23').value = meetings_1414
sheet_1414.range('B24').value = emails_1414
sheet_1414.range('B25').value = meetings_total_1414 
sheet_1414.range('B26').value = emails_total_1414
sheet_1414.range('A29:H200').value = ""
sheet_1414.range('A29').value = rfp_list_1414
all_1414 = sheet_1414.range('A1:J200').value

#Joe Paratore
sheet_1837.range('B5').value = qgoal_1837
sheet_1837.range('B7').value = billed_1837
sheet_1837.range('B8').value = ideal_pacing_1837
sheet_1837.range('B9').value = differential_1837
sheet_1837.range('B12').value = cw_percent_1837
sheet_1837.range('B13').value = last_week_cw_1837
sheet_1837.range('B14').value = cw_change_1837
sheet_1837.range('B16').value = rfp_needed_1837
sheet_1837.range('B18').value = rfp_change_1837
sheet_1837.range('B20').value = expired_change_1837
sheet_1837.range('B21').value = rfp_needed_percent_1837
sheet_1837.range('B23').value = meetings_1837
sheet_1837.range('B24').value = emails_1837
sheet_1837.range('B25').value = meetings_total_1837
sheet_1837.range('B26').value = emails_total_1837
sheet_1837.range('A29:H200').value = ""
sheet_1837.range('A29').value = rfp_list_1837
all_1837 = sheet_1837.range('A1:J200').value

#Morgan Afshari
sheet_2907.range('B5').value = qgoal_2907
sheet_2907.range('B7').value = billed_2907
sheet_2907.range('B8').value = ideal_pacing_2907
sheet_2907.range('B9').value = differential_2907
sheet_2907.range('B12').value = cw_percent_2907
sheet_2907.range('B13').value = last_week_cw_2907
sheet_2907.range('B14').value = cw_change_2907
sheet_2907.range('B16').value = rfp_needed_2907
sheet_2907.range('B18').value = rfp_change_2907
sheet_2907.range('B20').value = expired_change_2907
sheet_2907.range('B21').value = rfp_needed_percent_2907
sheet_2907.range('B23').value = meetings_2907
sheet_2907.range('B24').value = emails_2907
sheet_2907.range('B25').value = meetings_total_2907
sheet_2907.range('B26').value = emails_total_2907
sheet_2907.range('A29:H200').value = ""
sheet_2907.range('A29').value = rfp_list_2907
all_2907 = sheet_2907.range('A1:J200').value

#Mike Shea
sheet_3099.range('B5').value = qgoal_3099
sheet_3099.range('B7').value = billed_3099
sheet_3099.range('B8').value = ideal_pacing_3099
sheet_3099.range('B9').value = differential_3099
sheet_3099.range('B12').value = cw_percent_3099
sheet_3099.range('B13').value = last_week_cw_3099
sheet_3099.range('B14').value = cw_change_3099
sheet_3099.range('B16').value = rfp_needed_3099
sheet_3099.range('B18').value = rfp_change_3099
sheet_3099.range('B20').value = expired_change_3099
sheet_3099.range('B21').value = rfp_needed_percent_3099
sheet_3099.range('B23').value = meetings_3099
sheet_3099.range('B24').value = emails_3099
sheet_3099.range('B25').value = meetings_total_3099
sheet_3099.range('B26').value = emails_total_3099
sheet_3099.range('A29:H200').value = ""
sheet_3099.range('A29').value = rfp_list_3099
all_3099 = sheet_3099.range('A1:J200').value

####################################################################################  
try:
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    save_path = 'D://users//kevinspang//my documents//excel//sales reports//jackson davis//jackson davis//Sales_Report_Jackson_Davis_{}.xlsx'.format(save_date)
    wb1.save(save_path)
    wb1.close()
except:
    wb1.close()
    
#create individual sheets for each team member
try:
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    wb3 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//jackson davis//Individual Sales Snapshot Template.xlsx')
    sheet_main = wb3.sheets['Main']
    #Andrew Piekos
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_1414
    wb3.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//Andrew Piekos//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Joe Paratore
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_1837
    wb3.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//joe paratore//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Morgan Afshari
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_2907
    wb3.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//morgan afshari//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Mike Shea
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_2907
    wb3.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//mike shea//Sales_Snapshot_{}.xlsx'.format(save_date))
    sheet_main.range('A1:J200').value = ""
    wb3.close()
except:
    print("Individual snapshot sheets could not be created")    

############################################################################################################################################
#Load in data to weekly tracking sheet
    
try:
    wb4 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//jackson davis//jackson davis//jackson team weekly report.xlsx')
    sheet_team2 = wb4.sheets['Team']
    sheet_1414 = wb4.sheets['Andrew Piekos']
    sheet_1837 = wb4.sheets['Joe Paratore']
    sheet_2907 = wb4.sheets['Morgan Afshari']
    sheet_3099 = wb4.sheets['Mike Shea']

    date_dict = { year+'-01-05':'4', year+'-01-10':'7', year+'-01-17':'10', year+'-01-24':'13', year+'-01-31':'16', year+'-02-07':'19', year+'-02-14':'22', year+'-02-21':'25', year+'-02-28':'28', year+'-03-07':'31',
                  year+'-03-14':'34', year+'-03-21':'37', year+'-03-28':'40', year+'-04-04':'43', year+'-04-11':'46', year+'-04-18':'49', year+'-04-25':'52', year+'-05-02':'55', year+'-05-09':'58', year+'-05-16':'61',
                  year+'-05-23':'64', year+'-05-30':'67', year+'-06-06':'70', year+'-06-13':'73', year+'-06-20':'76', year+'-06-27':'79', year+'-07-04':'82', year+'-07-11':'85', year+'-07-18':'88', year+'-07-25':'91',
                  year+'-08-01':'94', year+'-08-08':'97', year+'-08-15':'100', year+'-08-22':'103', year+'-08-29':'106', year+'-09-05':'109', year+'-09-12':'112', year+'-09-19':'115', year+'-09-26':'118',
                  year+'-10-05':'121', year+'-10-10':'124', year+'-10-17':'127', year+'-10-24':'130', year+'-10-31':'133', year+'-11-07':'136', year+'-11-14':'139', year+'-11-21':'142', year+'-11-28':'145',
                  year+'-12-05':'148', year+'-12-12':'151', year+'-12-19':'154', year+'-12-26':'157'}
    
    date_num = date_dict.get(today)
    
    sheet_team2.range('D{}'.format(date_num)).value = weekly_sheet_team2
    sheet_1414.range('D{}'.format(date_num)).value = weekly_sheet_1414
    sheet_1837.range('D{}'.format(date_num)).value = weekly_sheet_1837
    sheet_2907.range('D{}'.format(date_num)).value = weekly_sheet_2907
    sheet_3099.range('D{}'.format(date_num)).value = weekly_sheet_3099
    wb4.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//jackson davis//jackson team weekly report.xlsx')
    wb4.close()
except:
    print("Not a Monday or other issue with the weekly report")
    
#try:
wb5 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//jackson davis//jackson davis//jackson team weekly report.xlsx')
sheet_1414 = wb5.sheets['Andrew Piekos']
sheet_1837 = wb5.sheets['Joe Paratore']
sheet_2907 = wb5.sheets['Morgan Afshari']
sheet_3099 = wb5.sheets['Mike Shea']

#Grab data for individual sheets
weekly_sheet_all_1414 = sheet_1414.range('A1:P160').value
weekly_sheet_all_1837 = sheet_1837.range('A1:P160').value
weekly_sheet_all_2907 = sheet_2907.range('A1:P160').value
weekly_sheet_all_3099 = sheet_3099.range('A1:P160').value


wb5.close()
#Input data into indivdual sheets
save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
wb6 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//jackson davis//Individual Weekly Report.xlsx')
sheet_main = wb6.sheets['Main']
#Andrew Piekos
sheet_main.range('A1:P200').value = ""
sheet_main.range('A1').value = weekly_sheet_all_1414
wb6.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//andrew piekos//Weekly_Sales_{}.xlsx'.format(save_date))
#Joe Paratore
sheet_main.range('A1:P200').value = ""
sheet_main.range('A1').value = weekly_sheet_all_1837
wb6.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//joe paratore//Weekly_Sales_{}.xlsx'.format(save_date))
#Morgan Afshari
sheet_main.range('A1:P200').value = ""
sheet_main.range('A1').value = weekly_sheet_all_2907
wb6.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//morgan afshari//Weekly_Sales_{}.xlsx'.format(save_date))
#Mike Shea
sheet_main.range('A1:P200').value = ""
sheet_main.range('A1').value = weekly_sheet_all_3099
wb6.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//mike shea//Weekly_Sales_{}.xlsx'.format(save_date))
sheet_main.range('A1:P200').value = ""
wb6.close()
#except:
#     print("Individual weekly sheets could not be created")    

############################################################################################################################################
#if last monday of the month, wipe out template data. 

if today == last_monday:
        wb1 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//jackson davis//sales_report_jackson_davis_template.xlsx' )
        sheet_team2 = wb1.sheets['Team']
        sheet_1414 = wb1.sheets['Andrew Piekos']
        sheet_1837 = wb1.sheets['Joe Paratore']
        sheet_2907 = wb1.sheets['Morgan Afshari']
        sheet_3099 = wb1.sheets['Mike Shea']
        #Delete out all previous quarters data
        sheet_team2.range('B11').value = 0
        sheet_team2.range('B13').value = 0
        sheet_team2.range('B17').value = 0
        sheet_team2.range('B19').value = 0
        sheet_1414.range('B11').value = 0
        sheet_1414.range('B13').value = 0
        sheet_1414.range('B17').value = 0
        sheet_1414.range('B19').value = 0
        sheet_1837.range('B11').value = 0
        sheet_1837.range('B13').value = 0
        sheet_1837.range('B17').value = 0
        sheet_1837.range('B19').value = 0
        sheet_2907.range('B11').value = 0
        sheet_2907.range('B13').value = 0
        sheet_2907.range('B17').value = 0
        sheet_2907.range('B19').value = 0
        sheet_3099.range('B11').value = 0
        sheet_3099.range('B13').value = 0
        sheet_3099.range('B17').value = 0
        sheet_3099.range('B19').value = 0
        wb1.save('D://users//kevinspang//my documents//excel//sales reports//jackson davis//Sales_Report_Jackson_Davis_Template.xlsx')
        wb1.close()
        app.kill()
        time.sleep(2)
else:
    app.kill()
    time.sleep(2)