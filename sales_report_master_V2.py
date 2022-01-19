# -*- coding: utf-8 -*-
"""
Created on Tue Nov  5 13:37:53 2019

@author: Kevin_Spang
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
# see if todays sheet exists if it does cancel sheet creation to avoid messing up template number    
# will then run scripts to get total agency information
# will grab teams and individuals from snapshots sheet and save
# will then grab and run weekly sheets and save second sheet

# if statement to only run on Mondays
if dt.date.today().isoweekday() == 1: #change to one for mondays
    #if monday check if todays sheet has already run.
    file_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    file_path = 'D://users//kevinspang//my documents//excel//sales reports//master//master//sales_report_master_{}.xlsx'.format(file_date) 
    if os.path.isfile(file_path) == True:
        print("Todays Sheet Already Complete")
        time.sleep(10)
        sys.exit(1)
    else:
        print("Running Master Sales Report")
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
    qgoal_master = float(goals.range("C23").value)
elif q2start <= today <= q2end:
    quartername = 'Q2'
    quarterstart = year + '-04-01'
    quarterend = year + '-06-30'
    last_monday = year + '-06-28'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_master = float(goals.range("D23").value)
elif q3start <= today <= q3end:
    quartername = 'Q3'
    quarterstart = year + '-07-01'
    quarterend = year + '-09-30'
    last_monday = year + '-09-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days 
    qgoal_master = float(goals.range("E23").value)
else:
    quartername = 'Q4'
    quarterstart = year + '-10-01'
    quarterend = year + '-12-31' 
    last_monday = year + '-12-27'
    daysInQuarter = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(quarterstart,dformat)).days
    daysPast = (dt.datetime.strptime(today,dformat) - dt.datetime.strptime(quarterstart,dformat)).days 
    daysRemaining = (dt.datetime.strptime(quarterend,dformat) - dt.datetime.strptime(today,dformat)).days
    qgoal_master = float(goals.range("F23").value)

wbgoals.close()

#####################################################################################################################################  
wb1 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//master//sales_report_master_template.xlsx' )

sheet_master = wb1.sheets['Master']
sheet_team1 = wb1.sheets['Mike Powell Team']
sheet_team2 = wb1.sheets['Jackson Davis Team']
sheet_team3 = wb1.sheets['Ted McNulty Team']

sheet_1079 = wb1.sheets['Jim Kappos']
sheet_1587 = wb1.sheets['Chris Pedevillano']
sheet_2617 = wb1.sheets['Dan Dechristoforo']
sheet_2908 = wb1.sheets['Patrick McTague']
sheet_3162 = wb1.sheets['Benjamin Crosby']
sheet_3117 = wb1.sheets['Chris Ligas']

sheet_1414 = wb1.sheets['Andrew Piekos']
sheet_1837 = wb1.sheets['Joe Paratore']
sheet_2907 = wb1.sheets['Morgan Afshari']
sheet_3099 = wb1.sheets['Mike Shea']

sheet_2265 = wb1.sheets['Alan Lai']
sheet_2473 = wb1.sheets['Jefferson Butler']
sheet_3019 = wb1.sheets['Adam Churilla']

#grab last weeks closed won rfp and expired values
last_week_cw_master = float(sheet_master.range('B11').value)
sheet_master.range('B13').value = last_week_cw_master
last_week_open_master = float(sheet_master.range('B17').value)
last_week_expired_master = float(sheet_master.range('B19').value)

#####################################################################################################################################  

name_id = "'1402','1414','1837','2907','3099','892','1079','1587','2617','2908','3162','3117','2265','2473','3019'"  

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
billed_master = [float(i[0]) for i in c.fetchall()].pop(0) #Billed to date

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
cw_master = [float(i[0]) for i in c.fetchall()].pop(0) #Total closed Won

cw_change_master = cw_master - last_week_cw_master #closed won change 

cw_percent_master = cw_master / qgoal_master #CW percent to goal

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
rfp_master = [float(i[0]) for i in c.fetchall()].pop(0) #open Deals

rfp_change_master = rfp_master - last_week_open_master #open Deals Change

#total expired RFP
c.execute('''select ifnull((
Select sum(budget) from (
Select sf.budget 
from sales_forecast sf
where sf.deal_stage in (3,6)
and start_date >= date_format(date_sub(sysdate(), interval 1 Month), '%Y-%m-%d')
and sf.salesPerson_id in ({}) 
group by sf.id) a),0) expired_rfp'''.format(name_id))
rfp_expired_master = [float(i[0]) for i in c.fetchall()].pop(0) #expired deals

expired_change_master = rfp_expired_master - last_week_expired_master #expired deals change

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
rfp_list_master = c.fetchall() #list of rfp, expired, negotiating, ect.

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

weekly_sheet_master = c.fetchall()

ideal_pacing_master = qgoal_master/daysInQuarter * daysPast #Ideal Pace number

differential_master = (billed_master - ideal_pacing_master) / ideal_pacing_master #pacing differential

rfp_needed_master = qgoal_master - cw_master #amount of rfp needed for quarter

try:
    rfp_needed_percent_master = rfp_needed_master/rfp_master #percent of open deals needed
except:
    rfp_needed_percent_master = 'N/A'


print("Completed running SQL scripts")    
#insert new values into template and save
    
sheet_master.range('A1').value = today
sheet_master.range('A2').value = 'Days Remaining in Quarter: {}'.format(daysRemaining)    
sheet_master.range('B11').value = cw_master
sheet_master.range('B17').value = rfp_master
sheet_master.range('B19').value = rfp_expired_master 

#save template for next week
wb1.save('D://users//kevinspang//my documents//excel//sales reports//master//Sales_Report_Master_Template.xlsx')

#load in salesforce data
wb2 = xw.Book('D://users//kevinspang//documents//excel//sales reports//salesforce.xlsx' )
salesforce = wb2.sheets['Salesforce']

meetings_master = salesforce.range('C23').value
meetings_total_master = salesforce.range('D23').value
emails_master = salesforce.range('E23').value
emails_total_master = salesforce.range('F23').value

wb2.close()

#insert into master tab 
sheet_master.range('B5').value = qgoal_master
sheet_master.range('B7').value = billed_master
sheet_master.range('B8').value = ideal_pacing_master
sheet_master.range('B9').value = differential_master
sheet_master.range('B12').value = cw_percent_master
sheet_master.range('B13').value = last_week_cw_master
sheet_master.range('B14').value = cw_change_master
sheet_master.range('B16').value = rfp_needed_master
sheet_master.range('B18').value = rfp_change_master
sheet_master.range('B20').value = expired_change_master
sheet_master.range('B21').value = rfp_needed_percent_master
sheet_master.range('B23').value = meetings_master
sheet_master.range('B24').value = emails_master
sheet_master.range('B25').value = meetings_total_master
sheet_master.range('B26').value = emails_total_master
sheet_master.range('A29:H200').value = ""
sheet_master.range('A29').value = rfp_list_master

############################################################################################################################################ 
#grab data from other sheets

#get workbook names
save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
mike_team_snapshot = 'D://users//kevinspang//my documents//excel//sales reports//mike powell//mike powell//Sales_Report_Mike_Powell_{}.xlsx'.format(save_date)
jackson_team_snapshot = 'D://users//kevinspang//my documents//excel//sales reports//jackson davis//jackson davis//Sales_Report_Jackson_Davis_{}.xlsx'.format(save_date)
ted_team_snapshot = 'D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//ted mcnulty//Sales_Report_Ted_McNulty_{}.xlsx'.format(save_date)

#grab mikes team snapshot data
try:
    wb3 = xw.Book(mike_team_snapshot)
    sheet_team1_snap = wb3.sheets['Team']
    sheet_1079_snap = wb3.sheets['Jim Kappos']
    sheet_1587_snap = wb3.sheets['Chris Pedevillano']
    sheet_2617_snap = wb3.sheets['Dan Dechristoforo']
    sheet_2908_snap = wb3.sheets['Patrick McTague']
    sheet_3162_snap = wb3.sheets['Benjamin Crosby']
    sheet_3117_snap = wb3.sheets['Chris Ligas']
    snapshot_team1 = sheet_team1_snap .range('A1:J200').value
    snapshot_1079 = sheet_1079_snap.range('A1:J200').value
    snapshot_1587 = sheet_1587_snap.range('A1:J200').value
    snapshot_2617 = sheet_2617_snap.range('A1:J200').value
    snapshot_2908 = sheet_2908_snap.range('A1:J200').value
    snapshot_3162 = sheet_3162_snap.range('A1:J200').value
    snapshot_3117 = sheet_3117_snap.range('A1:J200').value
    wb3.close()    
except:
    print("Issue pulling Mike's team data,check that sheet was run")
    sys.exit()
    
#grab jacksons team snapshot data
try:
    wb4 = xw.Book(jackson_team_snapshot)
    sheet_team2_snap = wb4.sheets['Team']
    sheet_1414_snap = wb4.sheets['Andrew Piekos']
    sheet_1837_snap = wb4.sheets['Joe Paratore']
    sheet_2907_snap = wb4.sheets["Morgan Afshari"]
    sheet_3099_snap = wb4.sheets["Mike Shea"]
    snapshot_team2 = sheet_team2_snap.range('A1:J200').value
    snapshot_1414 = sheet_1414_snap.range('A1:J200').value
    snapshot_1837 = sheet_1837_snap.range('A1:J200').value
    snapshot_2907 = sheet_2907_snap.range('A1:J200').value
    snapshot_3099 = sheet_3099_snap.range('A1:J200').value
    wb4.close()    
except:
    print("Issue pulling Jackson's team data,check that sheet was run")
    sys.exit()

#grab tims team snapshot data
try:
    wb5 = xw.Book(ted_team_snapshot)
    sheet_team3_snap = wb5.sheets['Team']
    sheet_2265_snap = wb5.sheets['Alan Lai']
    sheet_2473_snap = wb5.sheets['Jefferson Butler']
    sheet_3019_snap = wb5.sheets['Adam Churilla']
    snapshot_team3 = sheet_team3_snap.range('A1:J200').value
    snapshot_2265 = sheet_2265_snap.range('A1:J200').value
    snapshot_2473 = sheet_2473_snap.range('A1:J200').value
    snapshot_3019 = sheet_3019_snap.range('A1:J200').value
    wb5.close()    
except:
    print("Issue pulling Tim's team data,check that sheet was run")
    sys.exit()
    
############################################################################################################################################ 
#load everything into template and save
    
try:
    sheet_team1.range('A1').value = snapshot_team1
    sheet_team2.range('A1').value = snapshot_team2
    sheet_team3.range('A1').value = snapshot_team3
    sheet_1079.range('A1').value = snapshot_1079
    sheet_1587.range('A1').value = snapshot_1587
    sheet_2617.range('A1').value = snapshot_2617
    sheet_2908.range('A1').value = snapshot_2908
    sheet_3162.range('A1').value = snapshot_3162
    sheet_3117.range('A1').value = snapshot_3117
    sheet_1414.range('A1').value = snapshot_1414
    sheet_1837.range('A1').value = snapshot_1837
    sheet_2907.range('A1').value = snapshot_2907
    sheet_3099.range('A1').value = snapshot_3099
    sheet_2265.range('A1').value = snapshot_2265
    sheet_2473.range('A1').value = snapshot_2473
    sheet_3019.range('A1').value = snapshot_3019
    save_date = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
    save_path1 = 'D://users//kevinspang//my documents//excel//sales reports//master//master//Sales_Report_Master_{}.xlsx'.format(save_date)
    save_path2 = 'D://users//kevinspang//my documents//excel//sales reports//Master//Master//Sales_Report_Master_{}.xlsx'.format(save_date)
    wb1.save(save_path1)
    wb1.save(save_path2)
    wb1.close()
except:
    print("Issue inputting team and individual data into template")
    sys.exit()  

############################################################################################################################################ 
#Grab data from weekly sheets
    
#get workbook names
mike_team_weekly = 'D://users//kevinspang//my documents//excel//sales reports//mike powell//mike powell//Mike team weekly report.xlsx'
jackson_team_weekly = 'D://users//kevinspang//my documents//excel//sales reports//jackson davis//jackson davis//jackson team weekly report.xlsx'
ted_team_weekly = 'D://users//kevinspang//my documents//excel//sales reports//ted mcnulty//ted mcnulty//ted team weekly report.xlsx'  
    
try:
    wb6 = xw.Book(mike_team_weekly)
    sheet_weekly_team1 = wb6.sheets['Team']
    sheet_weekly_1079 = wb6.sheets['Jim Kappos']
    sheet_weekly_1587 = wb6.sheets['Chris Pedevillano']
    sheet_weekly_2617 = wb6.sheets['Dan Dechristoforo']
    sheet_weekly_2908 = wb6.sheets['Patrick McTague']
    sheet_weekly_3162 = wb6.sheets['Benjamin Crosby']
    sheet_weekly_3117 = wb6.sheets['Chris Ligas']
    weekly_team1 = sheet_weekly_team1.range('A1:P160').value
    weekly_1079 = sheet_weekly_1079.range('A1:P160').value
    weekly_1587 = sheet_weekly_1587.range('A1:P160').value
    weekly_2617 = sheet_weekly_2617.range('A1:P160').value
    weekly_2908 = sheet_weekly_2908.range('A1:P160').value
    weekly_3162 = sheet_weekly_3162.range('A1:P160').value
    weekly_3117 = sheet_weekly_3117.range('A1:P160').value 
    wb6.close() 
except:
    print("Issue pulling Mike's weekly data")
    sys.exit()      
    
#grab jacksons team snapshot data
try:
    wb7 = xw.Book(jackson_team_weekly)
    sheet_weekly_team2 = wb7.sheets['Team']
    sheet_weekly_1414 = wb7.sheets['Andrew Piekos']
    sheet_weekly_1837 = wb7.sheets['Joe Paratore']
    sheet_weekly_2907 = wb7.sheets['Morgan Afshari']
    sheet_weekly_3099 = wb7.sheets['Mike Shea']
    weekly_team2 = sheet_weekly_team2.range('A1:P160').value
    weekly_1414 = sheet_weekly_1414.range('A1:P160').value
    weekly_1837 = sheet_weekly_1837.range('A1:P160').value
    weekly_2907 = sheet_weekly_2907.range('A1:P160').value
    weekly_3099 = sheet_weekly_3099.range('A1:P160').value
    wb7.close()    
except:
    print("Issue pulling Jackson's weekly data") 
    sys.exit()    
    
#grab tims team snapshot data
try:
    wb8 = xw.Book(ted_team_weekly)
    sheet_weekly_team3 = wb8.sheets['Team']
    sheet_weekly_2473 = wb8.sheets['Jefferson Butler']
    sheet_weekly_2265 = wb8.sheets['Alan Lai']
    sheet_weekly_3019 = wb8.sheets['Adam Churilla']
    weekly_team3 = sheet_weekly_team3.range('A1:P160').value
    weekly_2473 = sheet_weekly_2473.range('A1:P160').value
    weekly_2265 = sheet_weekly_2265.range('A1:P160').value
    weekly_3019 = sheet_weekly_3019.range('A1:P160').value
    wb8.close()    
except:
    print("Issue pulling Tim's team data,check that sheet was run")
    sys.exit()    
    
############################################################################################################################################    
#Load in data to weekly tracking sheet
    
try:
    wb9 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//master//master//master team weekly report.xlsx')
    sheet_master = wb9.sheets['Master']
    sheet_team1 = wb9.sheets['Mike Powell Team']
    sheet_team2 = wb9.sheets['Jackson Davis Team']
    sheet_team3 = wb9.sheets['Ted McNulty Team']

    sheet_1079 = wb9.sheets['Jim Kappos']
    sheet_1587 = wb9.sheets['Chris Pedevillano']
    sheet_2617 = wb9.sheets['Dan Dechristoforo']
    
    sheet_2908 = wb9.sheets['Patrick McTague']
    sheet_3162 = wb9.sheets['Benjamin Crosby']
    sheet_3117 = wb9.sheets['Chris Ligas']

    sheet_1414 = wb9.sheets['Andrew Piekos']
    sheet_1837 = wb9.sheets['Joe Paratore']
    sheet_2907 = wb9.sheets["Morgan Afshari"]
    sheet_3099 = wb9.sheets['Mike Shea']
    
    sheet_2473 = wb9.sheets['Jefferson Butler']
    sheet_2265 = wb9.sheets['Alan Lai']
    sheet_3019 = wb9.sheets['Adam Churilla']
    
    
    date_dict = { year+'-01-05':'4', year+'-01-10':'7', year+'-01-17':'10', year+'-01-24':'13', year+'-01-31':'16', year+'-02-07':'19', year+'-02-14':'22', year+'-02-21':'25', year+'-02-28':'28', year+'-03-07':'31',
                  year+'-03-14':'34', year+'-03-21':'37', year+'-03-28':'40', year+'-04-04':'43', year+'-04-11':'46', year+'-04-18':'49', year+'-04-25':'52', year+'-05-02':'55', year+'-05-09':'58', year+'-05-16':'61',
                  year+'-05-23':'64', year+'-05-30':'67', year+'-06-06':'70', year+'-06-13':'73', year+'-06-20':'76', year+'-06-27':'79', year+'-07-04':'82', year+'-07-11':'85', year+'-07-18':'88', year+'-07-25':'91',
                  year+'-08-01':'94', year+'-08-08':'97', year+'-08-15':'100', year+'-08-22':'103', year+'-08-29':'106', year+'-09-05':'109', year+'-09-12':'112', year+'-09-19':'115', year+'-09-26':'118',
                  year+'-10-05':'121', year+'-10-10':'124', year+'-10-17':'127', year+'-10-24':'130', year+'-10-31':'133', year+'-11-07':'136', year+'-11-14':'139', year+'-11-21':'142', year+'-11-28':'145',
                  year+'-12-05':'148', year+'-12-12':'151', year+'-12-19':'154', year+'-12-26':'157'}
    
    date_num = date_dict.get(today)
    
    sheet_master.range('D{}'.format(date_num)).value = weekly_sheet_master
    sheet_team1.range('A1').value = weekly_team1
    sheet_team2.range('A1').value = weekly_team2
    sheet_team3.range('A1').value = weekly_team3
    
    sheet_1079.range('A1').value = weekly_1079
    sheet_1587.range('A1').value = weekly_1587
    sheet_2617.range('A1').value = weekly_2617
    sheet_2908.range('A1').value = weekly_2908
    sheet_3162.range('A1').value = weekly_3162
    sheet_3117.range('A1').value = weekly_3117
    
    sheet_1414.range('A1').value = weekly_1414
    sheet_1837.range('A1').value= weekly_1837
    sheet_2907.range('A1').value = weekly_2907
    sheet_3099.range('A1').value= weekly_3099
    
    sheet_2473.range('A1').value = weekly_2473
    sheet_2265.range('A1').value = weekly_2265
    sheet_3019.range('A1').value = weekly_3019
    wb9.save('D://users//kevinspang//my documents//excel//sales reports//master//master//Master Team Weekly Report.xlsx')
    wb9.close()
except:
    print("Could not load team data into weekly report")

############################################################################################################################################
#if last monday of the month, wipe out template data. 

if today == last_monday:
        wb1 = xw.Book('D://users//kevinspang//my documents//excel//sales reports//master//sales_report_master_template.xlsx')
        sheet_master = wb1.sheets['Master']
        #Delete out all previous quarters data
        sheet_master.range('B11').value = 0
        sheet_master.range('B13').value = 0
        sheet_master.range('B17').value = 0
        sheet_master.range('B19').value = 0
        wb1.save('D://users//kevinspang//my documents//excel//sales reports//master//Sales_Report_Ted_Mcnulty_Template.xlsx')
        wb1.close()
        app.kill()
        time.sleep(2)
else:
    app.kill()
    time.sleep(2)