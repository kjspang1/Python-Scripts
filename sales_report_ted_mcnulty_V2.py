# -*- coding: utf-8 -*-
"""
Created on Wed Feb  5 09:45:42 2020

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
    file_path = 'D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//sales_report_ted_mcnulty_{}.xlsx'.format(file_date) 
    if os.path.isfile(file_path) == True:
        print("Todays Sheet Already Complete")
        sys.exit(1)
    else:
        print("Running Ted McNulty's sales report")
else:
    print("Report only runs on Mondays")
    sys.exit(1)
    
##############################################################################################################################
#date logic for scripts
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
    qgoal_team3 = float(goals.range("C22").value)
    qgoal_2473 = float(goals.range("C16").value) #Jefferson Butler 
    qgoal_2265 = float(goals.range("C17").value) #Alan Lai
    qgoal_3019 = float(goals.range("C18").value) #Adam Churilla
elif q2start <= today <= q2end:
    quartername = 'Q2'
    quarterstart = year + '-04-01'
    quarterend = year + '-06-30'
    last_monday = year + '-06-28'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team3 = float(goals.range("D22").value)
    qgoal_2473 = float(goals.range("D16").value) #Jefferson Butler 
    qgoal_2265 = float(goals.range("D17").value) #Alan Lai
    qgoal_3019 = float(goals.range("D18").value) #Adam Churilla
elif q3start <= today <= q3end:
    quartername = 'Q3'
    quarterstart = year + '-07-01'
    quarterend = year + '-09-30'
    last_monday = year + '-09-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team3 = float(goals.range("E22").value)
    qgoal_2473 = float(goals.range("E16").value) #Jefferson Butler 
    qgoal_2265 = float(goals.range("E17").value) #Alan Lai
    qgoal_3019 = float(goals.range("E18").value) #Adam Churilla
else:
    quartername = 'Q4'
    quarterstart = year + '-07-01'
    quarterend = year + '-09-30'
    last_monday = year + '-09-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team3 = float(goals.range("F22").value)
    qgoal_2473 = float(goals.range("F16").value) #Jefferson Butler 
    qgoal_2265 = float(goals.range("F17").value) #Alan Lai
    qgoal_3019 = float(goals.range("F18").value) #Adam Churilla

wbgoals.close()    
##################################################################################################################   
wb1 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//sales_report_ted_mcnulty_template.xlsx' )

sheet_team3 = wb1.sheets['Team']
sheet_2473 = wb1.sheets['Jefferson Butler']
sheet_2265 = wb1.sheets['Alan Lai']
sheet_3019 = wb1.sheets['Adam Churilla']

#grab last weeks closed won rfp and expired values
last_week_cw_team3 = float(sheet_team3.range('B11').value)
sheet_team3.range('B13').value = last_week_cw_team3
last_week_open_team3 = float(sheet_team3.range('B17').value)
last_week_expired_team3 = float(sheet_team3.range('B19').value)

last_week_cw_2473 = float(sheet_2473.range('B11').value)
sheet_2473.range('B13').value = last_week_cw_2473
last_week_open_2473 = float(sheet_2473.range('B17').value)
last_week_expired_2473 = float(sheet_2473.range('B19').value)

last_week_cw_2265 = float(sheet_2265.range('B11').value)
sheet_2265.range('B13').value = last_week_cw_2265
last_week_open_2265 = float(sheet_2265.range('B17').value)
last_week_expired_2265 = float(sheet_2265.range('B19').value)

last_week_cw_3019 = float(sheet_3019.range('B11').value)
sheet_3019.range('B13').value = last_week_cw_3019
last_week_open_3019 = float(sheet_3019.range('B17').value)
last_week_expired_3019 = float(sheet_3019.range('B19').value)

##################################################################################################################   
#Team 3
    
name_id = "'2265','2473','3019'"
    

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
billed_team3 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_team3 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_team3 = cw_team3 - last_week_cw_team3 #closed won change 

cw_percent_team3 = cw_team3 / qgoal_team3 #CW percent to goal

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
rfp_team3 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_team3 = rfp_team3 - last_week_open_team3 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_team3 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_team3 = rfp_expired_team3 - last_week_expired_team3 #expired deals change

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
rfp_list_team3 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_team3 = c.fetchall()

ideal_pacing_team3 = qgoal_team3/daysInQuarter * daysPast #Ideal Pace number

differential_team3 = (billed_team3 - ideal_pacing_team3) / ideal_pacing_team3 #pacing differential

rfp_needed_team3 = qgoal_team3 - cw_team3 #amount of rfp needed for quarter

try:
    rfp_needed_percent_team3 = rfp_needed_team3/rfp_team3 #percent of open deals needed
except:
    rfp_needed_percent_team3 = 'N/A'

################################################################################################################    
#Alan Lai
    
name_id = "'2265'"      

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
billed_2265 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_2265 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_2265 = cw_2265 - last_week_cw_2265 #closed won change 

cw_percent_2265 = cw_2265 / qgoal_2265 #CW percent to goal

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
rfp_2265 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_2265 = rfp_2265 - last_week_open_2265 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_2265 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_2265 = rfp_expired_2265 - last_week_expired_2265 #expired deals change

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
rfp_list_2265 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_2265 = c.fetchall()

ideal_pacing_2265 = qgoal_2265/daysInQuarter * daysPast #Ideal Pace number

differential_2265 = (billed_2265 - ideal_pacing_2265) / ideal_pacing_2265 #pacing differential

rfp_needed_2265 = qgoal_2265 - cw_2265 #amount of rfp needed for quarter

try:
    rfp_needed_percent_2265 = rfp_needed_2265/rfp_2265 #percent of open deals needed
except:
    rfp_needed_percent_2265 = 'N/A'
################################################################################################################
#Jefferson Butler
    
name_id = "'2473'"      

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
billed_2473 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_2473 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_2473 = cw_2473 - last_week_cw_2473 #closed won change 

cw_percent_2473 = cw_2473 / qgoal_2473 #CW percent to goal

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
rfp_2473 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_2473 = rfp_2473 - last_week_open_2473 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_2473 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_2473 = rfp_expired_2473 - last_week_expired_2473 #expired deals change

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
rfp_list_2473 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_2473 = c.fetchall()

ideal_pacing_2473 = qgoal_2473/daysInQuarter * daysPast #Ideal Pace number

differential_2473 = (billed_2473 - ideal_pacing_2473) / ideal_pacing_2473 #pacing differential

rfp_needed_2473 = qgoal_2473 - cw_2473 #amount of rfp needed for quarter

try:
    rfp_needed_percent_2473 = rfp_needed_2473/rfp_2473 #percent of open deals needed
except:
    rfp_needed_percent_2473 = 'N/A'
    
################################################################################################################
#Adam Churilla
    
name_id = "'3019'"      

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
billed_3019 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_3019 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_3019 = cw_3019 - last_week_cw_3019 #closed won change 

cw_percent_3019 = cw_3019 / qgoal_3019 #CW percent to goal

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
rfp_3019 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_3019 = rfp_3019 - last_week_open_3019 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_3019 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_3019 = rfp_expired_3019 - last_week_expired_3019 #expired deals change

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
rfp_list_3019 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_3019 = c.fetchall()

ideal_pacing_3019 = qgoal_3019/daysInQuarter * daysPast #Ideal Pace number

differential_3019 = (billed_3019 - ideal_pacing_3019) / ideal_pacing_3019 #pacing differential

rfp_needed_3019 = qgoal_3019 - cw_3019 #amount of rfp needed for quarter

try:
    rfp_needed_percent_3019 = rfp_needed_3019/rfp_3019 #percent of open deals needed
except:
    rfp_needed_percent_3019 = 'N/A'

################################################################################################################  
#insert new values into template and save

sheet_team3.range('A1').value = today
sheet_team3.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_team3.range('B11').value = cw_team3
sheet_team3.range('B17').value = rfp_team3
sheet_team3.range('B19').value = rfp_expired_team3     

sheet_2473.range('A1').value = today
sheet_2473.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_2473.range('B11').value = cw_2473
sheet_2473.range('B17').value = rfp_2473
sheet_2473.range('B19').value = rfp_expired_2473   

sheet_2265.range('A1').value = today
sheet_2265.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_2265.range('B11').value = cw_2265
sheet_2265.range('B17').value = rfp_2265
sheet_2265.range('B19').value = rfp_expired_2265

sheet_3019.range('A1').value = today
sheet_3019.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_3019.range('B11').value = cw_3019
sheet_3019.range('B17').value = rfp_3019
sheet_3019.range('B19').value = rfp_expired_3019

#save template for next week
wb1.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//Sales_Report_Ted_McNulty_Template.xlsx')

#################################################################################################################################
#load in salesforce data

wb2 = xw.Book('D://users//kevinspang//documents//excel//sales reports//salesforce.xlsx' )
salesforce = wb2.sheets['Salesforce']

meetings_team3 = salesforce.range('C22').value
meetings_total_team3 = salesforce.range('D22').value
emails_team3 = salesforce.range('E22').value
emails_total_team3 = salesforce.range('F22').value

meetings_2265 = salesforce.range('C16').value
meetings_total_2265 = salesforce.range('D16').value
emails_2265 = salesforce.range('E16').value
emails_total_2265 = salesforce.range('F16').value

meetings_2473 = salesforce.range('C17').value
meetings_total_2473 = salesforce.range('D17').value
emails_2473 = salesforce.range('E17').value
emails_total_2473 = salesforce.range('F17').value

meetings_3019 = salesforce.range('C18').value
meetings_total_3019 = salesforce.range('D18').value
emails_3019 = salesforce.range('E18').value
emails_total_3019 = salesforce.range('F18').value

wb2.close()
#################################################################################################################################
#fill out remainder of sheet for distribution   
 
#Team 3 sheet
sheet_team3.range('B5').value = qgoal_team3
sheet_team3.range('B7').value = billed_team3
sheet_team3.range('B8').value = ideal_pacing_team3
sheet_team3.range('B9').value = differential_team3
sheet_team3.range('B12').value = cw_percent_team3
sheet_team3.range('B13').value = last_week_cw_team3
sheet_team3.range('B14').value = cw_change_team3
sheet_team3.range('B16').value = rfp_needed_team3
sheet_team3.range('B18').value = rfp_change_team3
sheet_team3.range('B20').value = expired_change_team3
sheet_team3.range('B21').value = rfp_needed_percent_team3
sheet_team3.range('B23').value = meetings_team3
sheet_team3.range('B24').value = emails_team3
sheet_team3.range('B25').value = meetings_total_team3
sheet_team3.range('B26').value = emails_total_team3
sheet_team3.range('A29:H200').value = ""
sheet_team3.range('A29').value = rfp_list_team3

#Jefferson Butler
sheet_2473.range('B5').value = qgoal_2473
sheet_2473.range('B7').value = billed_2473
sheet_2473.range('B8').value = ideal_pacing_2473
sheet_2473.range('B9').value = differential_2473
sheet_2473.range('B12').value = cw_percent_2473
sheet_2473.range('B13').value = last_week_cw_2473
sheet_2473.range('B14').value = cw_change_2473
sheet_2473.range('B16').value = rfp_needed_2473
sheet_2473.range('B18').value = rfp_change_2473
sheet_2473.range('B20').value = expired_change_2473
sheet_2473.range('B21').value = rfp_needed_percent_2473
sheet_2473.range('B23').value = meetings_2473
sheet_2473.range('B24').value = emails_2473
sheet_2473.range('B25').value = meetings_total_2473
sheet_2473.range('B26').value = emails_total_2473
sheet_2473.range('A29:H200').value = ""
sheet_2473.range('A29').value = rfp_list_2473
all_2473 = sheet_2473.range('A1:J200').value

#Alan Lai
sheet_2265.range('B5').value = qgoal_2265
sheet_2265.range('B7').value = billed_2265
sheet_2265.range('B8').value = ideal_pacing_2265
sheet_2265.range('B9').value = differential_2265
sheet_2265.range('B12').value = cw_percent_2265
sheet_2265.range('B13').value = last_week_cw_2265
sheet_2265.range('B14').value = cw_change_2265
sheet_2265.range('B16').value = rfp_needed_2265
sheet_2265.range('B18').value = rfp_change_2265
sheet_2265.range('B20').value = expired_change_2265
sheet_2265.range('B21').value = rfp_needed_percent_2265
sheet_2265.range('B23').value = meetings_2265
sheet_2265.range('B24').value = emails_2265
sheet_2265.range('B25').value = meetings_total_2265
sheet_2265.range('B26').value = emails_total_2265
sheet_2265.range('A29:H200').value = ""
sheet_2265.range('A29').value = rfp_list_2265
all_2265 = sheet_2265.range('A1:J200').value

#Adam Churilla
sheet_3019.range('B5').value = qgoal_3019
sheet_3019.range('B7').value = billed_3019
sheet_3019.range('B8').value = ideal_pacing_3019
sheet_3019.range('B9').value = differential_3019
sheet_3019.range('B12').value = cw_percent_3019
sheet_3019.range('B13').value = last_week_cw_3019
sheet_3019.range('B14').value = cw_change_3019
sheet_3019.range('B16').value = rfp_needed_3019
sheet_3019.range('B18').value = rfp_change_3019
sheet_3019.range('B20').value = expired_change_3019
sheet_3019.range('B21').value = rfp_needed_percent_3019
sheet_3019.range('B23').value = meetings_3019
sheet_3019.range('B24').value = emails_3019
sheet_3019.range('B25').value = meetings_total_3019
sheet_3019.range('B26').value = emails_total_3019
sheet_3019.range('A29:H200').value = ""
sheet_3019.range('A29').value = rfp_list_3019
all_3019 = sheet_3019.range('A1:J200').value

####################################################################################  

try:
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    save_path = 'D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//ted mcnulty//Sales_Report_Ted_McNulty_{}.xlsx'.format(save_date)
    wb1.save(save_path)
    wb1.close()
except:
    wb1.close()

#create individual sheets for each team member
try:
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    wb3 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//Individual Sales Snapshot Template.xlsx')
    sheet_main = wb3.sheets['Main']
    #Alan Lai
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_2265
    wb3.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//alan lai//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Jefferson Butler
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_2473
    wb3.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//jefferson butler//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Adam Churilla
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_3019
    wb3.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//adam churilla//Sales_Snapshot_{}.xlsx'.format(save_date))
    sheet_main.range('A1:J200').value = ""
    wb3.close()
except:
    print("Individual snapshot sheets could not be created")

############################################################################################################################################
#Load in data to weekly tracking sheet
try:
    wb4 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//ted mcnulty//Ted Team Weekly Report.xlsx')
    sheet_team3 = wb4.sheets['Team']
    sheet_2473 = wb4.sheets['Jefferson Butler']
    sheet_2265 = wb4.sheets['Alan Lai']
    sheet_3019 = wb4.sheets['Adam Churilla']
    
    date_dict = { year+'-01-05':'4', year+'-01-10':'7', year+'-01-17':'10', year+'-01-24':'13', year+'-01-31':'16', year+'-02-07':'19', year+'-02-14':'22', year+'-02-21':'25', year+'-02-28':'28', year+'-03-07':'31',
                  year+'-03-14':'34', year+'-03-21':'37', year+'-03-28':'40', year+'-04-04':'43', year+'-04-11':'46', year+'-04-18':'49', year+'-04-25':'52', year+'-05-02':'55', year+'-05-09':'58', year+'-05-16':'61',
                  year+'-05-23':'64', year+'-05-30':'67', year+'-06-06':'70', year+'-06-13':'73', year+'-06-20':'76', year+'-06-27':'79', year+'-07-04':'82', year+'-07-11':'85', year+'-07-18':'88', year+'-07-25':'91',
                  year+'-08-01':'94', year+'-08-08':'97', year+'-08-15':'100', year+'-08-22':'103', year+'-08-29':'106', year+'-09-05':'109', year+'-09-12':'112', year+'-09-19':'115', year+'-09-26':'118',
                  year+'-10-05':'121', year+'-10-10':'124', year+'-10-17':'127', year+'-10-24':'130', year+'-10-31':'133', year+'-11-07':'136', year+'-11-14':'139', year+'-11-21':'142', year+'-11-28':'145',
                  year+'-12-05':'148', year+'-12-12':'151', year+'-12-19':'154', year+'-12-26':'157'}
    
    date_num = date_dict.get(today)
    
    sheet_team3.range('D{}'.format(date_num)).value = weekly_sheet_team3
    sheet_2473.range('D{}'.format(date_num)).value = weekly_sheet_2473
    sheet_2265.range('D{}'.format(date_num)).value = weekly_sheet_2265
    sheet_3019.range('D{}'.format(date_num)).value = weekly_sheet_3019
    
    wb4.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//ted mcnulty//Ted Team Weekly Report.xlsx')
    wb4.close()
except:
    print("Not a Monday or other issue with the weekly report")
    
#create individual sheets for each team member
try:
    wb5 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//ted mcnulty//Ted Team Weekly Report.xlsx')
    sheet_2265 = wb5.sheets['Alan Lai']
    sheet_2473 = wb5.sheets['Jefferson Butler']
    sheet_3019 = wb5.sheets['Adam Churilla']
    #Grab data for individual sheets
    weekly_sheet_all_2473 = sheet_2473.range('A1:P160').value
    weekly_sheet_all_2265 = sheet_2265.range('A1:P160').value
    weekly_sheet_all_3019 = sheet_3019.range('A1:P160').value

    wb5.close()
    #Input data into indivdual sheets
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    wb6 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//Individual Weekly Report.xlsx')
    sheet_main = wb6.sheets['Main']
    #Alan Lai
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_2265
    wb6.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//alan lai//Weekly_Sales_{}.xlsx'.format(save_date))
    #Jefferson Butler
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_2473
    wb6.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//jefferson butler//Weekly_Sales_{}.xlsx'.format(save_date))
    #Adam Churilla
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_3019
    wb6.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//adam churilla//Weekly_Sales_{}.xlsx'.format(save_date))
    sheet_main.range('A1:P200').value = ""
    wb6.close()
except:
    print("Individual weekly sheets could not be created")
    
############################################################################################################################################
#if last monday of the month, wipe out template data. 

if today == last_monday:
        wb1 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//sales_report_ted_mcnulty_template.xlsx' )
        sheet_team3 = wb1.sheets['Team']
        sheet_2473 = wb1.sheets['Jefferson Butler']
        sheet_2265 = wb1.sheets['Alan Lai']
        sheet_3019 = wb1.sheets['Adam Churilla']
        #Delete out all previous quarters data
        sheet_team3.range('B11').value = 0
        sheet_team3.range('B13').value = 0
        sheet_team3.range('B17').value = 0
        sheet_team3.range('B19').value = 0
        sheet_2265.range('B11').value = 0
        sheet_2265.range('B13').value = 0
        sheet_2265.range('B17').value = 0
        sheet_2265.range('B19').value = 0
        sheet_2473.range('B11').value = 0
        sheet_2473.range('B13').value = 0
        sheet_2473.range('B17').value = 0
        sheet_2473.range('B19').value = 0
        sheet_3019.range('B11').value = 0
        sheet_3019.range('B13').value = 0
        sheet_3019.range('B17').value = 0
        sheet_3019.range('B19').value = 0
        wb1.save('D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//Sales_Report_Ted_McNulty_Template.xlsx')
        wb1.close()
        app.kill()
        time.sleep(2)
else:
    app.kill()
    time.sleep(2)
