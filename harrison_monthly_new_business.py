# -*- coding: utf-8 -*-
"""
Created on Tue Jan  4 12:44:04 2022

@author: kevinspang
"""

import mysql.connector
import xlwings as xw
import datetime as dt
import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
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


#Date Logic

#Last day of previous Month
lastMonth = dt.datetime.strftime(dt.date.today().replace(day=1) - dt.timedelta(days=1),'%Y-%m-%d')

#First day of previous Month
firstMonth = dt.datetime.strftime(dt.date.today().replace(day=1) - dt.timedelta((dt.date.today().replace(day=1) - dt.timedelta(days=1)).day),'%Y-%m-%d')

#first day of the year
#Special case in January need to grab first of previous year
if '01-01' <= dt.datetime.strftime(dt.date.today(),'%m-%d') <= '01-31':
    firstOfYear = dt.datetime.strftime(dt.date.today() - dt.timedelta(days=365),'%Y') + '-01-01'
else:
    firstOfYear = dt.datetime.strftime(dt.date.today(),'%Y') + '-01-01'

titleMonth = dt.datetime.strftime(dt.date.today().replace(day=1) - dt.timedelta(days=1),'%B %Y')

##############################################################################################################################

#set included sales ids
ids = "'1402','1414','1837','2907','892','1079','1587','2265','2473','2617','2908','3019','3162'"
#set minimum Budget
budget = 5000

c = cnx.cursor()
c.execute('''select fu.salesperson,
ifnull(rfp.rfp_count,0) RFP_Count,
ifnull(cw.cw_count,0) Closed_Won_Count,
ifnull(cw.cw_count/rfp.rfp_count,0) Count_Win_rate,
ifnull(round(rfp.total_budget,2),0) RFP_Total_Budget,
ifnull(round(cw.cw_budget,2),0) Closed_Won_Budget,
ifnull(round(cw.cw_budget/rfp.total_budget,2),0) Budget_Win_Rate
from
(select id, concat(first_name," ",last_name) salesperson 
from fos_user 
where id in ({})) fu
left join 
(select salesPerson_id, count(client_id) RFP_Count, sum(budget) total_budget from (
Select salesPerson_id, client_id, budget, min(createdAt)
from sales_forecast
where createdAt between '{}' and '{}'
and client_id not in (select distinct client_id from sales_forecast where createdAt < '{}')
group by client_id) a
group by salesPerson_id) rfp
on rfp.salesPerson_id = fu.id
left join
(select salesPerson_id, count(client_id) cw_count, sum(budget) cw_budget from (
Select salesPerson_id, client_id, budget, min(createdAt)
from sales_forecast
where createdAt between '{}' and '{}'
and client_id not in (select distinct client_id from sales_forecast where createdAt < '{}')
and deal_stage = '4'
group by client_id) b
group by salesPerson_id) cw
on cw.salesPerson_id = fu.id
order by 1'''.format(ids,firstMonth,lastMonth,firstMonth,firstMonth,lastMonth,firstMonth))
newBusiness = c.fetchall()

c.execute('''select fu.salesperson,
ifnull(rfp.rfp_count,0) RFP_Count,
ifnull(cw.cw_count,0) Closed_Won_Count,
ifnull(cw.cw_count/rfp.rfp_count,0) Count_Win_rate,
ifnull(round(rfp.total_budget,2),0) RFP_Total_Budget,
ifnull(round(cw.cw_budget,2),0) Closed_Won_Budget,
ifnull(round(cw.cw_budget/rfp.total_budget,2),0) Budget_Win_Rate
from
(select id, concat(first_name," ",last_name) salesperson 
from fos_user 
where id in ({})) fu
left join 
(select salesPerson_id, count(client_id) RFP_Count, sum(budget) total_budget from (
Select salesPerson_id, client_id, budget, min(createdAt)
from sales_forecast
where createdAt between '{}' and '{}'
and client_id not in (select distinct client_id from sales_forecast where createdAt < '{}')
group by client_id) a
group by salesPerson_id) rfp
on rfp.salesPerson_id = fu.id
left join
(select salesPerson_id, count(client_id) cw_count, sum(budget) cw_budget from (
Select salesPerson_id, client_id, budget, min(createdAt)
from sales_forecast
where createdAt between '{}' and '{}'
and client_id not in (select distinct client_id from sales_forecast where createdAt < '{}')
and deal_stage = '4'
group by client_id) b
group by salesPerson_id) cw
on cw.salesPerson_id = fu.id
order by 1'''.format(ids,firstOfYear,lastMonth,firstOfYear,firstOfYear,lastMonth,firstOfYear))
newBusinessTotal = c.fetchall()

c.execute('''select
concat(fu.first_name," ",fu.last_name) salesperson,
cl.name agency,
min.counts_min Total_Count_Below_Min,
min.RFP_Below_Minimum,
min.Closed_Won_Below_Minimum,
ifnull(max.counts_max,0) Total_Count_Above_Min,
ifnull(max.RFP_Above_Minimum,0) RFP_Above_Minimum,
ifnull(max.Closed_Won_Above_Minimum,0) Closed_Won_Above_Minimum
from
(Select 
sf.client_id,
sf.salesPerson_id,count(sf.id) counts_min,
sum(case when deal_stage = '1' then 1 else 0 end) as RFP_Below_Minimum,
sum(case when deal_stage = '4' then 1 else 0 end) as Closed_Won_Below_Minimum
from sales_forecast sf
where sf.deal_stage in ('4','1')
and sf.budget < {} 
and sf.client_id not in ('62','112')
and sf.end_date >= '{}'
group by sf.client_id) min
left join
(Select 
sf.client_id,
ifnull(count(sf.id),0) counts_max,
ifnull(sum(case when deal_stage = '1' then 1 else 0 end),0) as RFP_Above_Minimum,
ifnull(sum(case when deal_stage = '4' then 1 else 0 end),0) as Closed_Won_Above_Minimum
from sales_forecast sf
where sf.deal_stage in ('4','1')
and sf.budget >= {}
and sf.client_id not in ('62','112')
and sf.end_date >= '{}'
group by sf.client_id) max
on min.client_id = max.client_id
join client cl
on cl.id = min.client_id
join fos_user fu
on fu.id = min.salesPerson_id
order by 1'''.format(budget,lastMonth,budget,lastMonth))
minimumAccounts = c.fetchall()

a = """Select concat(fu.first_name," ",fu.last_name) salesperson,
cl.name agency,
sf.advertiser_name advertiser,
round(sf.budget,2) budget, 
sf.start_date,
sf.end_date
from sales_forecast sf
join client cl
on cl.id = sf.client_id
join fos_user fu
on fu.id = sf.salesPerson_id
where sf.id in (
Select distinct salesForecast_id 
from sales_forecast_audit_log
where date_format(created_at, '%Y-%m-%d') between '{}' and '{}'
and message like '%old\":1,\"new\":4%')
and sf.deal_stage = '4'
order by 4 desc
limit 1""".format(firstMonth,lastMonth)

c.execute(a)
largestIO = c.fetchall()
c.close()

##############################################################################################################################
#Create Sheet
app = xw.App(visible=False)  

try:
    wb1 = xw.Book('D://users//kevinspang//my documents//excel//harrison monthly//sales update template.xlsx')
    new_business_sheet = wb1.sheets['New Business RFPs (Monthly)']
    new_total_sheet = wb1.sheets['New Business RFPS (Cumulative)']
    minimum_sheet = wb1.sheets['Accounts Below Minimum']
    io_sheet = wb1.sheets['Largest signed IO']
    
    new_business_sheet.range('A2').value = newBusiness
    new_total_sheet.range('A2').value = newBusinessTotal
    minimum_sheet.range('A2').value = minimumAccounts
    io_sheet.range('A2').value = largestIO
    
    wb1.save('D://users//kevinspang//my documents//excel//harrison monthly//Sales {} Update.xlsx'.format(titleMonth))
    wb1.close()
    app.kill()
except:
    print("Could not create Sheet")
    app.kill()

##############################################################################################################################
#Email Sheet
Login = mysqlLogin()
pw = Login[4]
datemessage = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')
fromaddr = "kspang@addaptive.com"
toaddr = "hgraham@addaptive.com"
msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Sales {} Update".format(titleMonth)  
body = "Please see the attached excel sheet for the Sales {} Update. ".format(titleMonth) 
msg.attach(MIMEText(body, 'plain')) 
filename = 'Sales {} Update.xlsx'.format(titleMonth)
filepath = 'D://Users//kevinspang//Documents//Excel//harrison monthly//Sales {} Update.xlsx'.format(titleMonth)
attachment = open(filepath, "rb") 
p = MIMEBase('application', 'octet-stream') 
p.set_payload((attachment).read()) 
encoders.encode_base64(p) 
p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
msg.attach(p) 
s = smtplib.SMTP('smtp.gmail.com', 587) 
s.starttls() 
s.login(fromaddr, pw) 
text = msg.as_string() 
s.sendmail(fromaddr, toaddr, text) 
s.quit()  






















