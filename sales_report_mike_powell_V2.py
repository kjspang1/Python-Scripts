# -*- coding: utf-8 -*-
"""
Created on Fri Jan 31 10:39:25 2020

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
    file_path = 'D://users//kevinspang//documents//excel//sales reports//mike powell//mike powell//sales_report_mike_powell_{}.xlsx'.format(file_date) 
    if os.path.isfile(file_path) == True:
        print("Todays Sheet Already Complete")
        sys.exit(1)
    else:
        print("Running Mike Powell's sales report")
else:
    print("Report only runs on Mondays")
    sys.exit(1)

########################################################################################################################################
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
wbgoals = xw.Book('D://users//kevinspang//documents//excel//sales reports//Goals.xlsx')
goals = wbgoals.sheets['Goals']


if q1start <= today <= q1end:
    quartername = 'Q1'
    quarterstart = year + '-01-01'
    quarterend = year + '-03-31'
    last_monday = year + '-03-29'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team1 = float(goals.range("C20").value)
    qgoal_1079 = float(goals.range("C3").value)
    qgoal_1587 = float(goals.range("C4").value)
    qgoal_2617 = float(goals.range("C5").value)
    qgoal_2908 = float(goals.range("C6").value)
    qgoal_3162 = float(goals.range("C7").value)
    qgoal_3117 = float(goals.range("C8").value)
elif q2start <= today <= q2end:
    quartername = 'Q2'
    quarterstart = year + '-04-01'
    quarterend = year + '-06-30'
    last_monday = year + '-06-28'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team1 = float(goals.range("D20").value)
    qgoal_1079 = float(goals.range("D3").value)
    qgoal_1587 = float(goals.range("D4").value)
    qgoal_2617 = float(goals.range("D5").value)
    qgoal_2908 = float(goals.range("D6").value)
    qgoal_3162 = float(goals.range("D7").value)
    qgoal_3117 = float(goals.range("D8").value)
elif q3start <= today <= q3end:
    quartername = 'Q3'
    quarterstart = year + '-07-01'
    quarterend = year + '-09-30'
    last_monday = year + '-09-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team1 = float(goals.range("E20").value)
    qgoal_1079 = float(goals.range("E3").value)
    qgoal_1587 = float(goals.range("E4").value)
    qgoal_2617 = float(goals.range("E5").value)
    qgoal_2908 = float(goals.range("E6").value)
    qgoal_3162 = float(goals.range("E7").value)
    qgoal_3117 = float(goals.range("E8").value)
else:
    quartername = 'Q4'
    quarterstart = year + '-10-01'
    quarterend = year + '-12-31' 
    last_monday = year + '-12-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_team1 = float(goals.range("F20").value)
    qgoal_1079 = float(goals.range("F3").value)
    qgoal_1587 = float(goals.range("F4").value)
    qgoal_2617 = float(goals.range("F5").value)
    qgoal_2908 = float(goals.range("F6").value)
    qgoal_3162 = float(goals.range("F7").value)
    qgoal_3117 = float(goals.range("F8").value)

wbgoals.close()
################################################################################################################
#open template and grab previous weeks data
  
wb1 = xw.Book('D://users//kevinspang//documents//excel//sales reports//mike powell//sales_report_mike_powell_template.xlsx' )

sheet_team1 = wb1.sheets['Team']
sheet_1079 = wb1.sheets['Jim Kappos']
sheet_1587 = wb1.sheets['Chris Pedevillano']
sheet_2617 = wb1.sheets['Dan Dechristoforo']
sheet_2908 = wb1.sheets['Patrick McTague']
sheet_3162 = wb1.sheets['Benjamin Crosby']
sheet_3117 = wb1.sheets['Chris Ligas']

#grab last weeks closed won rfp and expired values
last_week_cw_team1 = float(sheet_team1.range('B11').value)
sheet_team1.range('B13').value = last_week_cw_team1
last_week_open_team1 = float(sheet_team1.range('B17').value)
last_week_expired_team1 = float(sheet_team1.range('B19').value)

last_week_cw_1079 = float(sheet_1079.range('B11').value)
sheet_1079.range('B13').value = last_week_cw_1079
last_week_open_1079 = float(sheet_1079.range('B17').value)
last_week_expired_1079 = float(sheet_1079.range('B19').value)

last_week_cw_1587 = float(sheet_1587.range('B11').value)
sheet_1587.range('B13').value = last_week_cw_1587
last_week_open_1587 = float(sheet_1587.range('B17').value)
last_week_expired_1587 = float(sheet_1587.range('B19').value)

last_week_cw_2617 = float(sheet_2617.range('B11').value)
sheet_2617.range('B13').value = last_week_cw_2617
last_week_open_2617 = float(sheet_2617.range('B17').value)
last_week_expired_2617 = float(sheet_2617.range('B19').value)

last_week_cw_2908 = float(sheet_2908.range('B11').value)
sheet_2908.range('B13').value = last_week_cw_2908
last_week_open_2908 = float(sheet_2908.range('B17').value)
last_week_expired_2908 = float(sheet_2908.range('B19').value)

last_week_cw_3162 = float(sheet_3162.range('B11').value)
sheet_3162.range('B13').value = last_week_cw_3162
last_week_open_3162 = float(sheet_3162.range('B17').value)
last_week_expired_3162 = float(sheet_3162.range('B19').value)

last_week_cw_3117 = float(sheet_3117.range('B11').value)
sheet_3117.range('B13').value = last_week_cw_3117
last_week_open_3117 = float(sheet_3117.range('B17').value)
last_week_expired_3117 = float(sheet_3117.range('B19').value)

################################################################################################################
#Team 1
    
name_id = "'892','1079','1587','2617','2908','3162','3117'"     

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
billed_team1 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_team1 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_team1 = cw_team1 - last_week_cw_team1 #closed won change

cw_percent_team1 = cw_team1 / qgoal_team1 #CW percent to goal 

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
rfp_team1 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_team1 = rfp_team1 - last_week_open_team1 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_team1 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_team1 = rfp_expired_team1 - last_week_expired_team1 #expired deals change

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
group by sf.id order by 10 desc,2,3;  '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_team1 = c.fetchall() #list of rfp, expired, negotiating, ect.

#list of remaining rfp
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

weekly_sheet_team1 = c.fetchall()

ideal_pacing_team1 = qgoal_team1/daysInQuarter * daysPast #Ideal Pace number

differential_team1 = (billed_team1 - ideal_pacing_team1) / ideal_pacing_team1 #pacing differential

rfp_needed_team1 = qgoal_team1 - cw_team1 #amount of rfp needed for quarter

try:
    rfp_needed_percent_team1 = rfp_needed_team1/rfp_team1 #percent of open deals needed
except:
    rfp_needed_percent_team1 = 'N/A'
    
################################################################################################################
#Jim Kappos
    
name_id = '1079'    

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
billed_1079 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_1079 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_1079 = cw_1079 - last_week_cw_1079 #closed won change 

cw_percent_1079 = cw_1079 / qgoal_1079 #CW percent to goal

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
rfp_1079 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_1079 = rfp_1079 - last_week_open_1079 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_1079 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_1079 = rfp_expired_1079 - last_week_expired_1079 #expired deals change

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
group by sf.id order by 10 desc,2,3;  '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_1079 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_1079 = c.fetchall()

ideal_pacing_1079 = qgoal_1079/daysInQuarter * daysPast #Ideal Pace number

differential_1079 = (billed_1079 - ideal_pacing_1079) / ideal_pacing_1079 #pacing differential

rfp_needed_1079 = qgoal_1079 - cw_1079 #amount of rfp needed for quarter

try:
    rfp_needed_percent_1079 = rfp_needed_1079/rfp_1079 #percent of open deals needed
except:
    rfp_needed_percent_1079 = 'N/A'
################################################################################################################
#Chris Pedevillano
    
name_id = '1587'    

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
billed_1587 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_1587 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_1587 = cw_1587 - last_week_cw_1587 #closed won change 

cw_percent_1587 = cw_1587 / qgoal_1587 #CW percent to goal

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
rfp_1587 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_1587 = rfp_1587 - last_week_open_1587 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_1587 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_1587 = rfp_expired_1587 - last_week_expired_1587 #expired deals change

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
group by sf.id order by 10 desc,2,3;  '''.format(quarterstart,quarterend,name_id,name_id))
rfp_list_1587 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_1587 = c.fetchall()

ideal_pacing_1587 = qgoal_1587/daysInQuarter * daysPast #Ideal Pace number

differential_1587 = (billed_1587 - ideal_pacing_1587) / ideal_pacing_1587 #pacing differential

rfp_needed_1587 = qgoal_1587 - cw_1587 #amount of rfp needed for quarter

try:
    rfp_needed_percent_1587 = rfp_needed_1587/rfp_1587 #percent of open deals needed
except:
    rfp_needed_percent_1587 = 'N/A'
    
################################################################################################################
#Dan Dechristoforo
    
name_id = '2617'    

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
billed_2617 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_2617 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_2617 = cw_2617 - last_week_cw_2617 #closed won change 

cw_percent_2617 = cw_2617 / qgoal_2617 #CW percent to goal

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
rfp_2617 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_2617 = rfp_2617 - last_week_open_2617 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_2617 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_2617 = rfp_expired_2617 - last_week_expired_2617 #expired deals change

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
rfp_list_2617 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_2617 = c.fetchall()

ideal_pacing_2617 = qgoal_2617/daysInQuarter * daysPast #Ideal Pace number

differential_2617 = (billed_2617 - ideal_pacing_2617) / ideal_pacing_2617 #pacing differential

rfp_needed_2617 = qgoal_2617 - cw_2617 #amount of rfp needed for quarter

try:
    rfp_needed_percent_2617 = rfp_needed_2617/rfp_2617 #percent of open deals needed
except:
    rfp_needed_percent_2617 = 'N/A'
    
################################################################################################################
#Patrick McTague
    
name_id = '2908'    

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
billed_2908 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_2908 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_2908 = cw_2908 - last_week_cw_2908 #closed won change 

cw_percent_2908 = cw_2908 / qgoal_2908 #CW percent to goal

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
rfp_2908 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_2908 = rfp_2908 - last_week_open_2908 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_2908 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_2908 = rfp_expired_2908 - last_week_expired_2908 #expired deals change

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
rfp_list_2908 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_2908 = c.fetchall()

ideal_pacing_2908 = qgoal_2908/daysInQuarter * daysPast #Ideal Pace number

differential_2908 = (billed_2908 - ideal_pacing_2908) / ideal_pacing_2908 #pacing differential

rfp_needed_2908 = qgoal_2908 - cw_2908 #amount of rfp needed for quarter

try:
    rfp_needed_percent_2908 = rfp_needed_2908/rfp_2908 #percent of open deals needed
except:
    rfp_needed_percent_2908 = 'N/A'

################################################################################################################
#Benjamin Crosby
    
name_id = '3162'    

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
billed_3162 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_3162 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_3162 = cw_3162 - last_week_cw_3162 #closed won change 

cw_percent_3162 = cw_3162 / qgoal_3162 #CW percent to goal

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
rfp_3162 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_3162 = rfp_3162 - last_week_open_3162 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_3162 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_3162 = rfp_expired_3162 - last_week_expired_3162 #expired deals change

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
rfp_list_3162 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_3162 = c.fetchall()

ideal_pacing_3162 = qgoal_3162/daysInQuarter * daysPast #Ideal Pace number

differential_3162 = (billed_3162 - ideal_pacing_3162) / ideal_pacing_3162 #pacing differential

rfp_needed_3162 = qgoal_3162 - cw_3162 #amount of rfp needed for quarter

try:
    rfp_needed_percent_3162 = rfp_needed_3162/rfp_3162 #percent of open deals needed
except:
    rfp_needed_percent_3162 = 'N/A'
    
################################################################################################################
#Chris Ligas
    
name_id = '3117'    

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
billed_3117 = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_3117 = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_3117 = cw_3117 - last_week_cw_3117 #closed won change 

cw_percent_3117 = cw_3117 / qgoal_3117 #CW percent to goal

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
rfp_3117 = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_3117 = rfp_3117 - last_week_open_3117 #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_3117 = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_3117 = rfp_expired_3117 - last_week_expired_3117 #expired deals change

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
rfp_list_3117 = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_3117 = c.fetchall()

ideal_pacing_3117 = qgoal_3117/daysInQuarter * daysPast #Ideal Pace number

differential_3117 = (billed_3117 - ideal_pacing_3117) / ideal_pacing_3117 #pacing differential

rfp_needed_3117 = qgoal_3117 - cw_3117 #amount of rfp needed for quarter

try:
    rfp_needed_percent_3117 = rfp_needed_3117/rfp_3117 #percent of open deals needed
except:
    rfp_needed_percent_3117 = 'N/A'
    
################################################################################################################    
#insert new values into template and save

sheet_team1.range('A1').value = today
sheet_team1.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_team1.range('B11').value = cw_team1
sheet_team1.range('B17').value = rfp_team1
sheet_team1.range('B19').value = rfp_expired_team1    

sheet_1079.range('A1').value = today
sheet_1079.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_1079.range('B11').value = cw_1079
sheet_1079.range('B17').value = rfp_1079
sheet_1079.range('B19').value = rfp_expired_1079

sheet_1587.range('A1').value = today
sheet_1587.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_1587.range('B11').value = cw_1587
sheet_1587.range('B17').value = rfp_1587
sheet_1587.range('B19').value = rfp_expired_1587

sheet_2617.range('A1').value = today
sheet_2617.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_2617.range('B11').value = cw_2617
sheet_2617.range('B17').value = rfp_2617
sheet_2617.range('B19').value = rfp_expired_2617

sheet_2908.range('A1').value = today
sheet_2908.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_2908.range('B11').value = cw_2908
sheet_2908.range('B17').value = rfp_2908
sheet_2908.range('B19').value = rfp_expired_2908

sheet_3162.range('A1').value = today
sheet_3162.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_3162.range('B11').value = cw_3162
sheet_3162.range('B17').value = rfp_3162
sheet_3162.range('B19').value = rfp_expired_3162

sheet_3117.range('A1').value = today
sheet_3117.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_3117.range('B11').value = cw_3117
sheet_3117.range('B17').value = rfp_3117
sheet_3117.range('B19').value = rfp_expired_3117

#save template for next week
wb1.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Sales_Report_Mike_Powell_Template.xlsx')

#################################################################################################################################
#load in salesforce data

wb2 = xw.Book('D://users//kevinspang//documents//excel//sales reports//salesforce.xlsx' )
salesforce = wb2.sheets['Salesforce']

meetings_team1 = salesforce.range('C20').value
meetings_total_team1 = salesforce.range('D20').value
emails_team1 = salesforce.range('E20').value
emails_total_team1 = salesforce.range('F20').value

meetings_1079 = salesforce.range('C3').value
meetings_total_1079 = salesforce.range('D3').value
emails_1079 = salesforce.range('E3').value
emails_total_1079 = salesforce.range('F3').value

meetings_1587 = salesforce.range('C4').value
meetings_total_1587 = salesforce.range('D4').value
emails_1587 = salesforce.range('E4').value
emails_total_1587 = salesforce.range('F4').value

meetings_2617 = salesforce.range('C5').value
meetings_total_2617 = salesforce.range('D5').value
emails_2617 = salesforce.range('E5').value
emails_total_2617 = salesforce.range('F5').value

meetings_2908 = salesforce.range('C6').value
meetings_total_2908 = salesforce.range('D6').value
emails_2908 = salesforce.range('E6').value
emails_total_2908 = salesforce.range('F6').value

meetings_3162 = salesforce.range('C7').value
meetings_total_3162 = salesforce.range('D7').value
emails_3162 = salesforce.range('E7').value
emails_total_3162 = salesforce.range('F7').value

meetings_3117 = salesforce.range('C8').value
meetings_total_3117 = salesforce.range('D8').value
emails_3117 = salesforce.range('E8').value
emails_total_3117 = salesforce.range('F8').value

wb2.close()

#################################################################################################################################
#fill out remainder of sheet for distribution    
#Team 1 sheet
sheet_team1.range('B5').value = qgoal_team1
sheet_team1.range('B7').value = billed_team1
sheet_team1.range('B8').value = ideal_pacing_team1
sheet_team1.range('B9').value = differential_team1
sheet_team1.range('B12').value = cw_percent_team1
sheet_team1.range('B13').value = last_week_cw_team1
sheet_team1.range('B14').value = cw_change_team1
sheet_team1.range('B16').value = rfp_needed_team1
sheet_team1.range('B18').value = rfp_change_team1
sheet_team1.range('B20').value = expired_change_team1
sheet_team1.range('B21').value = rfp_needed_percent_team1
sheet_team1.range('B23').value = meetings_team1
sheet_team1.range('B24').value = emails_team1
sheet_team1.range('B25').value = meetings_total_team1
sheet_team1.range('B26').value = emails_total_team1
sheet_team1.range('A29:J200').value = ""
sheet_team1.range('A29').value = rfp_list_team1

#Jim Kappos individual sheet
sheet_1079.range('B5').value = qgoal_1079
sheet_1079.range('B7').value = billed_1079
sheet_1079.range('B8').value = ideal_pacing_1079
sheet_1079.range('B9').value = differential_1079
sheet_1079.range('B12').value = cw_percent_1079
sheet_1079.range('B13').value = last_week_cw_1079
sheet_1079.range('B14').value = cw_change_1079
sheet_1079.range('B16').value = rfp_needed_1079
sheet_1079.range('B18').value = rfp_change_1079
sheet_1079.range('B20').value = expired_change_1079
sheet_1079.range('B21').value = rfp_needed_percent_1079
sheet_1079.range('B23').value = meetings_1079
sheet_1079.range('B24').value = emails_1079
sheet_1079.range('B25').value = meetings_total_1079
sheet_1079.range('B26').value = emails_total_1079
sheet_1079.range('A29:J200').value = ""
sheet_1079.range('A29').value = rfp_list_1079
all_1079 = sheet_1079.range('A1:J200').value

#Chris Pedevillano individual sheet
sheet_1587.range('B5').value = qgoal_1587
sheet_1587.range('B7').value = billed_1587
sheet_1587.range('B8').value = ideal_pacing_1587
sheet_1587.range('B9').value = differential_1587
sheet_1587.range('B12').value = cw_percent_1587
sheet_1587.range('B13').value = last_week_cw_1587
sheet_1587.range('B14').value = cw_change_1587
sheet_1587.range('B16').value = rfp_needed_1587
sheet_1587.range('B18').value = rfp_change_1587
sheet_1587.range('B20').value = expired_change_1587
sheet_1587.range('B21').value = rfp_needed_percent_1587
sheet_1587.range('B23').value = meetings_1587
sheet_1587.range('B24').value = emails_1587
sheet_1587.range('B25').value = meetings_total_1587
sheet_1587.range('B26').value = emails_total_1587
sheet_1587.range('A29:J200').value = ""
sheet_1587.range('A29').value = rfp_list_1587
all_1587 = sheet_1587.range('A1:J200').value

#Dan Dechristoforo individual sheet
sheet_2617.range('B5').value = qgoal_2617
sheet_2617.range('B7').value = billed_2617
sheet_2617.range('B8').value = ideal_pacing_2617
sheet_2617.range('B9').value = differential_2617
sheet_2617.range('B12').value = cw_percent_2617
sheet_2617.range('B13').value = last_week_cw_2617
sheet_2617.range('B14').value = cw_change_2617
sheet_2617.range('B16').value = rfp_needed_2617
sheet_2617.range('B18').value = rfp_change_2617
sheet_2617.range('B20').value = expired_change_2617
sheet_2617.range('B21').value = rfp_needed_percent_2617
sheet_2617.range('B23').value = meetings_2617
sheet_2617.range('B24').value = emails_2617
sheet_2617.range('B25').value = meetings_total_2617
sheet_2617.range('B26').value = emails_total_2617
sheet_2617.range('A29:J200').value = ""
sheet_2617.range('A29').value = rfp_list_2617
all_2617 = sheet_2617.range('A1:J200').value

#Patrick McTague individual sheet
sheet_2908.range('B5').value = qgoal_2908
sheet_2908.range('B7').value = billed_2908
sheet_2908.range('B8').value = ideal_pacing_2908
sheet_2908.range('B9').value = differential_2908
sheet_2908.range('B12').value = cw_percent_2908
sheet_2908.range('B13').value = last_week_cw_2908
sheet_2908.range('B14').value = cw_change_2908
sheet_2908.range('B16').value = rfp_needed_2908
sheet_2908.range('B18').value = rfp_change_2908
sheet_2908.range('B20').value = expired_change_2908
sheet_2908.range('B21').value = rfp_needed_percent_2908
sheet_2908.range('B23').value = meetings_2908
sheet_2908.range('B24').value = emails_2908
sheet_2908.range('B25').value = meetings_total_2908
sheet_2908.range('B26').value = emails_total_2908
sheet_2908.range('A29:J200').value = ""
sheet_2908.range('A29').value = rfp_list_2908
all_2908 = sheet_2908.range('A1:J200').value

#Benjamin Crosby individual sheet
sheet_3162.range('B5').value = qgoal_3162
sheet_3162.range('B7').value = billed_3162
sheet_3162.range('B8').value = ideal_pacing_3162
sheet_3162.range('B9').value = differential_3162
sheet_3162.range('B12').value = cw_percent_3162
sheet_3162.range('B13').value = last_week_cw_3162
sheet_3162.range('B14').value = cw_change_3162
sheet_3162.range('B16').value = rfp_needed_3162
sheet_3162.range('B18').value = rfp_change_3162
sheet_3162.range('B20').value = expired_change_3162
sheet_3162.range('B21').value = rfp_needed_percent_3162
sheet_3162.range('B23').value = meetings_3162
sheet_3162.range('B24').value = emails_3162
sheet_3162.range('B25').value = meetings_total_3162
sheet_3162.range('B26').value = emails_total_3162
sheet_3162.range('A29:J200').value = ""
sheet_3162.range('A29').value = rfp_list_3162
all_3162 = sheet_3162.range('A1:J200').value

#Chris Ligas individual sheet
sheet_3117.range('B5').value = qgoal_3117
sheet_3117.range('B7').value = billed_3117
sheet_3117.range('B8').value = ideal_pacing_3117
sheet_3117.range('B9').value = differential_3117
sheet_3117.range('B12').value = cw_percent_3117
sheet_3117.range('B13').value = last_week_cw_3117
sheet_3117.range('B14').value = cw_change_3117
sheet_3117.range('B16').value = rfp_needed_3117
sheet_3117.range('B18').value = rfp_change_3117
sheet_3117.range('B20').value = expired_change_3117
sheet_3117.range('B21').value = rfp_needed_percent_3117
sheet_3117.range('B23').value = meetings_3117
sheet_3117.range('B24').value = emails_3117
sheet_3117.range('B25').value = meetings_total_3117
sheet_3117.range('B26').value = emails_total_3117
sheet_3117.range('A29:J200').value = ""
sheet_3117.range('A29').value = rfp_list_3117
all_3117 = sheet_3117.range('A1:J200').value

#############################################################################################################################################
#save sales snapshot sheet for manager
try:
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    save_path = 'D://users//kevinspang//documents//excel//sales reports//mike powell//mike powell//Sales_Report_Mike_Powell_{}.xlsx'.format(save_date)
    wb1.save(save_path)
    wb1.close()
except:
    wb1.close()
    
#create individual sheets for each team member
try:
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    wb3 = xw.Book('D://users//kevinspang//documents//excel//sales reports//mike powell//Individual Sales Snapshot Template.xlsx')
    sheet_main = wb3.sheets['Main']
    #Jim Kappos
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_1079
    wb3.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Jim Kappos//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Chris Pedevillano
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_1587
    wb3.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Chris Pedevillano//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Dan Dichristoforo
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_2617
    wb3.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Dan Dechristoforo//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Patrick McTague
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_2908
    wb3.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Patrick Mctague//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Benjamin Crosby
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_3162
    wb3.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Benjamin Crosby//Sales_Snapshot_{}.xlsx'.format(save_date))
    #Chris Ligas
    sheet_main.range('A1:J200').value = ""
    sheet_main.range('A1').value = all_3117
    wb3.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Chris Ligas//Sales_Snapshot_{}.xlsx'.format(save_date))
    sheet_main.range('A1:J200').value = ""
    wb3.close()
except:
    print("Individual snapshot sheets could not be created")

############################################################################################################################################
#Load in data to weekly tracking sheet
try:
    wb4 = xw.Book('D://users//kevinspang//documents//excel//sales reports//mike powell//mike powell//Mike team weekly report.xlsx')
    sheet_team1 = wb4.sheets['Team']
    sheet_1079 = wb4.sheets['Jim Kappos']
    sheet_1587 = wb4.sheets['Chris Pedevillano']
    sheet_2617 = wb4.sheets['Dan Dechristoforo']
    sheet_2908 = wb4.sheets['Patrick McTague']
    sheet_3162 = wb4.sheets['Benjamin Crosby']
    sheet_3117 = wb4.sheets['Chris Ligas']
    
    date_dict = { year+'-01-05':'4', year+'-01-10':'7', year+'-01-17':'10', year+'-01-24':'13', year+'-01-31':'16', year+'-02-07':'19', year+'-02-14':'22', year+'-02-21':'25', year+'-02-28':'28', year+'-03-07':'31',
                  year+'-03-14':'34', year+'-03-21':'37', year+'-03-28':'40', year+'-04-04':'43', year+'-04-11':'46', year+'-04-18':'49', year+'-04-25':'52', year+'-05-02':'55', year+'-05-09':'58', year+'-05-16':'61',
                  year+'-05-23':'64', year+'-05-30':'67', year+'-06-06':'70', year+'-06-13':'73', year+'-06-20':'76', year+'-06-27':'79', year+'-07-04':'82', year+'-07-11':'85', year+'-07-18':'88', year+'-07-25':'91',
                  year+'-08-01':'94', year+'-08-08':'97', year+'-08-15':'100', year+'-08-22':'103', year+'-08-29':'106', year+'-09-05':'109', year+'-09-12':'112', year+'-09-19':'115', year+'-09-26':'118',
                  year+'-10-05':'121', year+'-10-10':'124', year+'-10-17':'127', year+'-10-24':'130', year+'-10-31':'133', year+'-11-07':'136', year+'-11-14':'139', year+'-11-21':'142', year+'-11-28':'145',
                  year+'-12-05':'148', year+'-12-12':'151', year+'-12-19':'154', year+'-12-26':'157'}
    
    date_num = date_dict.get(today)
    
    sheet_team1.range('D{}'.format(date_num)).value = weekly_sheet_team1
    sheet_1079.range('D{}'.format(date_num)).value = weekly_sheet_1079
    sheet_1587.range('D{}'.format(date_num)).value = weekly_sheet_1587
    sheet_2617.range('D{}'.format(date_num)).value = weekly_sheet_2617
    sheet_2908.range('D{}'.format(date_num)).value = weekly_sheet_2908
    sheet_3162.range('D{}'.format(date_num)).value = weekly_sheet_3162
    sheet_3117.range('D{}'.format(date_num)).value = weekly_sheet_3117
    
    wb4.save('D://users//kevinspang//documents//excel//sales reports//mike powell//mike powell//Mike Team Weekly Report.xlsx')
    wb4.close()
except:
    print("Not a Monday or other issue with the weekly report")

#create individual sheets for each team member
try:
    wb5 = xw.Book('D://users//kevinspang//documents//excel//sales reports//mike powell//mike powell//Mike team weekly report.xlsx')
    sheet_1079 = wb5.sheets['Jim Kappos']
    sheet_1587 = wb5.sheets['Chris Pedevillano']
    sheet_2617 = wb5.sheets['Dan Dechristoforo']
    sheet_2908 = wb5.sheets['Patrick McTague']
    sheet_3162 = wb5.sheets['Benjamin Crosby']
    sheet_3117 = wb5.sheets['Chris Ligas']
    #Grab data for individual sheets
    weekly_sheet_all_1079 = sheet_1079.range('A1:P160').value
    weekly_sheet_all_1587 = sheet_1587.range('A1:P160').value
    weekly_sheet_all_2617 = sheet_2617.range('A1:P160').value
    weekly_sheet_all_2908 = sheet_2908.range('A1:P160').value
    weekly_sheet_all_3162 = sheet_3162.range('A1:P160').value
    weekly_sheet_all_3117 = sheet_3117.range('A1:P160').value
    wb5.close()
    #Input data into indivdual sheets
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    wb6 = xw.Book('D://users//kevinspang//documents//excel//sales reports//mike powell//Individual Weekly Report.xlsx')
    sheet_main = wb6.sheets['Main']
    #Jim Kappos
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_1079
    wb6.save('D://users//kevinspang//documents//excel//sales reports//mike powell//jim kappos//Weekly_Sales_{}.xlsx'.format(save_date))
    #Chris Pedevillano
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_1587
    wb6.save('D://users//kevinspang//documents//excel//sales reports//mike powell//chris pedevillano//Weekly_Sales_{}.xlsx'.format(save_date))
    #Dan Dichristoforo
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_2617
    wb6.save('D://users//kevinspang//documents//excel//sales reports//mike powell//dan dechristoforo//Weekly_Sales_{}.xlsx'.format(save_date))
    #Patrick McTague
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_2908
    wb6.save('D://users//kevinspang//documents//excel//sales reports//mike powell//patrick mctague//Weekly_Sales_{}.xlsx'.format(save_date))
    #Benjamin Crosby
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_3162
    wb6.save('D://users//kevinspang//documents//excel//sales reports//mike powell//benjamin crosby//Weekly_Sales_{}.xlsx'.format(save_date))
    #Chris Ligas
    sheet_main.range('A1:P200').value = ""
    sheet_main.range('A1').value = weekly_sheet_all_3117
    wb6.save('D://users//kevinspang//documents//excel//sales reports//mike powell//chris ligas//Weekly_Sales_{}.xlsx'.format(save_date))
    sheet_main.range('A1:P200').value = ""
    wb6.close()
except:
    print("Individual weekly sheets could not be created")

############################################################################################################################################
#if last monday of the month, wipe out template data and kill excel. 

if today == last_monday:
        wb1 = xw.Book('D://users//kevinspang//documents//excel//sales reports//mike powell//sales_report_mike_powell_template.xlsx' )
        sheet_team1 = wb1.sheets['Team']
        sheet_1079 = wb1.sheets['Jim Kappos']
        sheet_1587 = wb1.sheets['Chris Pedevillano']
        sheet_2617 = wb1.sheets['Dan Dechristoforo']
        sheet_2908 = wb1.sheets['Patrick McTague']
        sheet_3162 = wb1.sheets['Benjamin Crosby']
        sheet_3117 = wb1.sheets['Chris Ligas']
        #Delete out all previous quarters data
        sheet_team1.range('B11').value = 0
        sheet_team1.range('B13').value = 0
        sheet_team1.range('B17').value = 0
        sheet_team1.range('B19').value = 0
        sheet_1079.range('B11').value = 0
        sheet_1079.range('B13').value = 0
        sheet_1079.range('B17').value = 0
        sheet_1079.range('B19').value = 0
        sheet_1587.range('B11').value = 0
        sheet_1587.range('B13').value = 0
        sheet_1587.range('B17').value = 0
        sheet_1587.range('B19').value = 0
        sheet_2617.range('B11').value = 0
        sheet_2617.range('B13').value = 0
        sheet_2617.range('B17').value = 0
        sheet_2617.range('B19').value = 0
        sheet_2908.range('B11').value = 0
        sheet_2908.range('B13').value = 0
        sheet_2908.range('B17').value = 0
        sheet_2908.range('B19').value = 0
        sheet_3162.range('B11').value = 0
        sheet_3162.range('B13').value = 0
        sheet_3162.range('B17').value = 0
        sheet_3162.range('B19').value = 0
        sheet_3117.range('B11').value = 0
        sheet_3117.range('B13').value = 0
        sheet_3117.range('B17').value = 0
        sheet_3117.range('B19').value = 0
        wb1.save('D://users//kevinspang//documents//excel//sales reports//mike powell//Sales_Report_Mike_Powell_Template.xlsx' )
        wb1.close()
        app.kill()
        time.sleep(2)
else:
    app.kill()
    time.sleep(2)