# -*- coding: utf-8 -*-
"""
Created on Wed Jun  2 08:08:21 2021

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

###############################################################################################################################
#datelogic and quarter selection
dformat = '%Y-%m-%d'
today = dt.datetime.strftime(dt.datetime.now(),dformat)
yesterday = dt.datetime.strftime(dt.date.today() - dt.timedelta(days=1),dformat)
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
    quarter = 'Q1'
    quarterstart = year + '-01-01'
    quarterend = year + '-03-31'
elif q2start <= today <= q2end:
    quarter = 'Q2'
    quarterstart = year + '-04-01'
    quarterend = year + '-06-30'
elif q3start <= today <= q3end:
    quarter = 'Q3'
    quarterstart = year + '-07-01'
    quarterend = year + '-09-30'
else:
    quarter = 'Q4'
    quarterstart = year + '-10-01'
    quarterend = year + '-12-31' 

##############################################################################################################################################################################################################################################################
    
#teams
Powell = "'892','1079','1587','2617','2908','3162','3117'"
Jackson = "'1402','1414','1837','2907','3099'"
Ted = "'2265','2473','3019'"
All  = "'1402','1414','1837','2907','3099','892','1079','1587','2617','2908','3162','3117','2265','2473','3019'"  

individual  = ['1414','1837','2907','3099',
               '1079','1587','2617','2908','3162','3117',
               '2265','2473','3019'] 

emails = ['apiekos@addaptive.com','jparatore@addaptive.com','mafshari@addaptive.com','mshea@addaptive.com',
          'jkappos@addaptive.com','cpedevillano@addaptive.com','ddechristoforo@addaptive.com','pmctague@addaptive.com','bcrosby@addaptive.com','cligas@addaptive.com',
          'alai@addaptive.com','jbutler@addaptive.com','achurilla@addaptive.com']
    
names = ['Andrew_Piekos','Joe_Paratore','Morgan_Afshari','Mike Shea',
         'Jim_Kappos','Chris_Pedevillano','Dan_Dechristoforo','Patrick_McTague','Benjamin_Crosby','Chris_Ligas',
         'Alan_Lai','Jeff_Butler','Adam_Churilla']    

##############################################################################################################################################################################################################################################################
    
#Run Mike Powells Team
c = cnx.cursor()
c.execute('''Select cw.client_id, 
cw.name, 
cw.salesPerson,
ifnull(round(pace.revenueToDate,2),0) revenueToDate, 
ifnull(round(cw.ideal_pace,2),0) ideal_pace, 
ifnull(round(pace.revenueToDate,2),0) - ifnull(round(cw.ideal_pace,2),0) Difference,
ifnull(round(cw.closed_won,2),0) closed_won, 
round(ifnull(cw.closed_won,0) - ifnull(pace.revenueToDate,0),2) remaining_closed_won from
(#get closed won and current pace
select client_id, name, salesPerson, sum(pace_budget) ideal_pace, sum(eoq_budget) Closed_won from (
select a.*, b.eoq_budget, c.pace_budget
from (
Select sf.id, sf.client_id, concat(fu.first_name," ",fu.last_name) salesperson,
cl.name
from sales_forecast sf
join client cl
on cl.id = sf.client_id
join fos_user fu
on fu.id = sf.salesPerson_id
where sf.id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and sf.salesPerson_id in ({})
and sf.deal_stage = '4' ) a,
#calculate closed won for the quarter
(Select  salesForecast_id, round(sum(budget),2) eoq_budget  from sales_forecast_daily_budget
where day between '{}' and '{}' #change to quarter date range
group by salesForecast_id) b,
#calculate closed won through current date -1
(Select  salesForecast_id, round(sum(budget),2) pace_budget  from sales_forecast_daily_budget
where day between '{}' and date_sub(sysdate(), interval 1 day) #change to previous day close
group by salesForecast_id) c
where a.id = b.salesForecast_id
and a.id = c.salesForecast_id
) d group by client_id) cw,
(#get current pace to date
select lir.client_id, SUM(CASE
WHEN c.type = 'agency' 
THEN round(lir.cpm * lir.impressions / 1000,2)
END) AS revenueToDate 
from line_item_report lir
join client c
on c.id = lir.client_id
WHERE report_date between '{}' and date_sub(sysdate(), interval 1 day) #Must change start of quarter
and client_id in (Select distinct client_id
from sales_forecast where id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and salesPerson_id in ({})
and deal_stage = '4' )
group by lir.client_id) pace
where pace.client_id = cw.client_id
order by 6'''.format(quarterstart,quarterend,Powell,quarterstart,quarterend,quarterstart,quarterstart,quarterstart,quarterend,Powell))   
powell_results = c.fetchall() 
    
#Run Jacksons team
c.execute('''Select cw.client_id, 
cw.name, 
cw.salesPerson,
ifnull(round(pace.revenueToDate,2),0) revenueToDate, 
ifnull(round(cw.ideal_pace,2),0) ideal_pace, 
ifnull(round(pace.revenueToDate,2),0) - ifnull(round(cw.ideal_pace,2),0) Difference,
ifnull(round(cw.closed_won,2),0) closed_won, 
round(ifnull(cw.closed_won,0) - ifnull(pace.revenueToDate,0),2) remaining_closed_won from
(#get closed won and current pace
select client_id, name, salesPerson, sum(pace_budget) ideal_pace, sum(eoq_budget) Closed_won from (
select a.*, b.eoq_budget, c.pace_budget
from (
Select sf.id, sf.client_id, concat(fu.first_name," ",fu.last_name) salesperson,
cl.name
from sales_forecast sf
join client cl
on cl.id = sf.client_id
join fos_user fu
on fu.id = sf.salesPerson_id
where sf.id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and sf.salesPerson_id in ({})
and sf.deal_stage = '4' ) a,
#calculate closed won for the quarter
(Select  salesForecast_id, round(sum(budget),2) eoq_budget  from sales_forecast_daily_budget
where day between '{}' and '{}' #change to quarter date range
group by salesForecast_id) b,
#calculate closed won through current date -1
(Select  salesForecast_id, round(sum(budget),2) pace_budget  from sales_forecast_daily_budget
where day between '{}' and date_sub(sysdate(), interval 1 day) #change to previous day close
group by salesForecast_id) c
where a.id = b.salesForecast_id
and a.id = c.salesForecast_id
) d group by client_id) cw,
(#get current pace to date
select lir.client_id, SUM(CASE
WHEN c.type = 'agency' 
THEN round(lir.cpm * lir.impressions / 1000,2)
END) AS revenueToDate 
from line_item_report lir
join client c
on c.id = lir.client_id
WHERE report_date between '{}' and date_sub(sysdate(), interval 1 day) #Must change start of quarter
and client_id in (Select distinct client_id
from sales_forecast where id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and salesPerson_id in ({})
and deal_stage = '4' )
group by lir.client_id) pace
where pace.client_id = cw.client_id
order by 6 '''.format(quarterstart,quarterend,Jackson,quarterstart,quarterend,quarterstart,quarterstart,quarterstart,quarterend,Jackson))   
jackson_results = c.fetchall()    

#Run Teds team
c.execute('''Select cw.client_id, 
cw.name, 
cw.salesPerson,
ifnull(round(pace.revenueToDate,2),0) revenueToDate, 
ifnull(round(cw.ideal_pace,2),0) ideal_pace, 
ifnull(round(pace.revenueToDate,2),0) - ifnull(round(cw.ideal_pace,2),0) Difference,
ifnull(round(cw.closed_won,2),0) closed_won, 
round(ifnull(cw.closed_won,0) - ifnull(pace.revenueToDate,0),2) remaining_closed_won from
(#get closed won and current pace
select client_id, name, salesPerson, sum(pace_budget) ideal_pace, sum(eoq_budget) Closed_won from (
select a.*, b.eoq_budget, c.pace_budget
from (
Select sf.id, sf.client_id, concat(fu.first_name," ",fu.last_name) salesperson,
cl.name
from sales_forecast sf
join client cl
on cl.id = sf.client_id
join fos_user fu
on fu.id = sf.salesPerson_id
where sf.id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and sf.salesPerson_id in ({})
and sf.deal_stage = '4' ) a,
#calculate closed won for the quarter
(Select  salesForecast_id, round(sum(budget),2) eoq_budget  from sales_forecast_daily_budget
where day between '{}' and '{}' #change to quarter date range
group by salesForecast_id) b,
#calculate closed won through current date -1
(Select  salesForecast_id, round(sum(budget),2) pace_budget  from sales_forecast_daily_budget
where day between '{}' and date_sub(sysdate(), interval 1 day) #change to previous day close
group by salesForecast_id) c
where a.id = b.salesForecast_id
and a.id = c.salesForecast_id
) d group by client_id) cw,
(#get current pace to date
select lir.client_id, SUM(CASE
WHEN c.type = 'agency' 
THEN round(lir.cpm * lir.impressions / 1000,2)
END) AS revenueToDate 
from line_item_report lir
join client c
on c.id = lir.client_id
WHERE report_date between '{}' and date_sub(sysdate(), interval 1 day) #Must change start of quarter
and client_id in (Select distinct client_id
from sales_forecast where id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and salesPerson_id in ({})
and deal_stage = '4' )
group by lir.client_id) pace
where pace.client_id = cw.client_id
order by 6 '''.format(quarterstart,quarterend,Ted,quarterstart,quarterend,quarterstart,quarterstart,quarterstart,quarterend,Ted))   
ted_results = c.fetchall()  
    

#Run ALL 
c.execute('''Select cw.client_id, 
cw.name, 
cw.salesPerson,
ifnull(round(pace.revenueToDate,2),0) revenueToDate, 
ifnull(round(cw.ideal_pace,2),0) ideal_pace, 
ifnull(round(pace.revenueToDate,2),0) - ifnull(round(cw.ideal_pace,2),0) Difference,
ifnull(round(cw.closed_won,2),0) closed_won, 
round(ifnull(cw.closed_won,0) - ifnull(pace.revenueToDate,0),2) remaining_closed_won from
(#get closed won and current pace
select client_id, name, salesPerson, sum(pace_budget) ideal_pace, sum(eoq_budget) Closed_won from (
select a.*, b.eoq_budget, c.pace_budget
from (
Select sf.id, sf.client_id, concat(fu.first_name," ",fu.last_name) salesperson,
cl.name
from sales_forecast sf
join client cl
on cl.id = sf.client_id
join fos_user fu
on fu.id = sf.salesPerson_id
where sf.id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and sf.salesPerson_id in ({})
and sf.deal_stage = '4' ) a,
#calculate closed won for the quarter
(Select  salesForecast_id, round(sum(budget),2) eoq_budget  from sales_forecast_daily_budget
where day between '{}' and '{}' #change to quarter date range
group by salesForecast_id) b,
#calculate closed won through current date -1
(Select  salesForecast_id, round(sum(budget),2) pace_budget  from sales_forecast_daily_budget
where day between '{}' and date_sub(sysdate(), interval 1 day) #change to previous day close
group by salesForecast_id) c
where a.id = b.salesForecast_id
and a.id = c.salesForecast_id
) d group by client_id) cw,
(#get current pace to date
select lir.client_id, SUM(CASE
WHEN c.type = 'agency' 
THEN round(lir.cpm * lir.impressions / 1000,2)
END) AS revenueToDate 
from line_item_report lir
join client c
on c.id = lir.client_id
WHERE report_date between '{}' and date_sub(sysdate(), interval 1 day) #Must change start of quarter
and client_id in (Select distinct client_id
from sales_forecast where id in 
(Select  salesForecast_id from sales_forecast_daily_budget
where day between '{}' and '{}') #change to quarter date range
and salesPerson_id in ({})
and deal_stage = '4' )
group by lir.client_id) pace
where pace.client_id = cw.client_id
order by 6 '''.format(quarterstart,quarterend,All,quarterstart,quarterend,quarterstart,quarterstart,quarterstart,quarterend,All))   
full_results = c.fetchall() 


#######################################################################################################################################################################  
    
#Create Mikes Team 
try:   
    app = xw.App(visible=False)
    file_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Template.xlsx'
    wb1 = xw.Book(file_path)
    main_sheet = wb1.sheets['Main']
    main_sheet.range('A2').value = powell_results
    save_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Mike Powell//Delivered_vs_Closed_Won_MP_{}_{}.xlsx'.format(quarter,today)
    wb1.save(save_path)
    wb1.close
    app.kill()      
except:
    print("Powells Sheet already created") 
    
#Create Jacksons Team    
try:
    app = xw.App(visible=False)
    file_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Template.xlsx'
    wb1 = xw.Book(file_path)
    main_sheet = wb1.sheets['Main']
    main_sheet.range('A2').value = jackson_results
    save_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Jackson Davis//Delivered_vs_Closed_Won_JD_{}_{}.xlsx'.format(quarter,today)
    wb1.save(save_path)
    wb1.close
    app.kill()   
except:
    print("Jacksons Sheet already created")
    
#Create Master  
try:  
    app = xw.App(visible=False)
    file_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Template_Master.xlsx'
    wb1 = xw.Book(file_path)
    main_sheet = wb1.sheets['Main']
    powell_sheet = wb1.sheets['Powell']
    jackson_sheet = wb1.sheets['Jackson']
    ted_sheet = wb1.sheets['Ted']
    main_sheet.range('A2').value = full_results
    powell_sheet.range('A2').value = powell_results
    jackson_sheet.range('A2').value = jackson_results
    ted_sheet.range('A2').value = ted_results
    save_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Master//Delivered_vs_Closed_Won_Master_{}_{}.xlsx'.format(quarter,today)
    wb1.save(save_path)
    wb1.close
    app.kill()   
except:
    print("Master Sheet already created")
######################################################################################################################################################################

#send email to mike powell
Login = mysqlLogin()
pw = Login[4]
fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"
toaddr = "mpowell@addaptive.com"
msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Delivered vs. Closed Won {}".format(today)  
body = "Please see the attached excel sheet for the {} Delivered vs. Closed Won report. ".format(today)  
msg.attach(MIMEText(body, 'plain')) 
filename = 'Delivered_vs_Closed_Won_MP_{}_{}.xlsx'.format(quarter,today)
filepath = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Mike Powell//Delivered_vs_Closed_Won_MP_{}_{}.xlsx'.format(quarter,today)
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

#send email Jackson Davis
Login = mysqlLogin()
pw = Login[4]
fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"
toaddr = "jdavis@addaptive.com"
msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Delivered vs. Closed Won {}".format(today)  
body = "Please see the attached excel sheet for the {} Delivered vs. Closed Won report. ".format(today)  
msg.attach(MIMEText(body, 'plain')) 
filename = 'Delivered_vs_Closed_Won_JD_{}_{}.xlsx'.format(quarter,today)
filepath = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Jackson Davis//Delivered_vs_Closed_Won_JD_{}_{}.xlsx'.format(quarter,today)
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

#Send Master Out
master_emails = ["mmahoney@addaptive.com","hgraham@addaptive.com","tmcnulty@addaptive.com"]
for i in master_emails:
    Login = mysqlLogin()
    pw = Login[4]
    datemessage = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')
    fromaddr = "kspang@addaptive.com"
    toaddr = i
    msg = MIMEMultipart() 
    msg['From'] = fromaddr   
    msg['To'] = toaddr
    msg['Subject'] = "Delivered vs. Closed Won {}".format(today)  
    body = "Please see the attached excel sheet for the {} Delivered vs. Closed Won report. ".format(today) 
    msg.attach(MIMEText(body, 'plain')) 
    filename = 'Delivered_vs_Closed_Won_Master_{}_{}.xlsx'.format(quarter,today)
    filepath = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Master//Delivered_vs_Closed_Won_Master_{}_{}.xlsx'.format(quarter,today)
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
    
#######################################################################################################################################################################
#Send individual sheets to reps

while len(individual) > 0:
    app = xw.App(visible=False)
    i = individual.pop(0)
    email = emails.pop(0)
    name = names.pop(0)
    
    c.execute('''Select cw.client_id, 
    cw.name, 
    cw.salesPerson,
    ifnull(round(pace.revenueToDate,2),0) revenueToDate, 
    ifnull(round(cw.ideal_pace,2),0) ideal_pace, 
    ifnull(round(pace.revenueToDate,2),0) - ifnull(round(cw.ideal_pace,2),0) Difference,
    ifnull(round(cw.closed_won,2),0) closed_won, 
    round(ifnull(cw.closed_won,0) - ifnull(pace.revenueToDate,0),2) remaining_closed_won from
    (#get closed won and current pace
    select client_id, name, salesPerson, sum(pace_budget) ideal_pace, sum(eoq_budget) Closed_won from (
    select a.*, b.eoq_budget, c.pace_budget
    from (
    Select sf.id, sf.client_id, concat(fu.first_name," ",fu.last_name) salesperson,
    cl.name
    from sales_forecast sf
    join client cl
    on cl.id = sf.client_id
    join fos_user fu
    on fu.id = sf.salesPerson_id
    where sf.id in 
    (Select  salesForecast_id from sales_forecast_daily_budget
    where day between '{}' and '{}') #change to quarter date range
    and sf.salesPerson_id in ({})
    and sf.deal_stage = '4' ) a,
    #calculate closed won for the quarter
    (Select  salesForecast_id, round(sum(budget),2) eoq_budget  from sales_forecast_daily_budget
    where day between '{}' and '{}' #change to quarter date range
    group by salesForecast_id) b,
    #calculate closed won through current date -1
    (Select  salesForecast_id, round(sum(budget),2) pace_budget  from sales_forecast_daily_budget
    where day between '{}' and date_sub(sysdate(), interval 1 day) #change to previous day close
    group by salesForecast_id) c
    where a.id = b.salesForecast_id
    and a.id = c.salesForecast_id
    ) d group by client_id) cw,
    (#get current pace to date
    select lir.client_id, SUM(CASE
    WHEN c.type = 'agency' 
    THEN round(lir.cpm * lir.impressions / 1000,2)
    END) AS revenueToDate 
    from line_item_report lir
    join client c
    on c.id = lir.client_id
    WHERE report_date between '{}' and date_sub(sysdate(), interval 1 day) #Must change start of quarter
    and client_id in (Select distinct client_id
    from sales_forecast where id in 
    (Select  salesForecast_id from sales_forecast_daily_budget
    where day between '{}' and '{}') #change to quarter date range
    and salesPerson_id in ({})
    and deal_stage = '4' )
    group by lir.client_id) pace
    where pace.client_id = cw.client_id
    order by 6'''.format(quarterstart,quarterend,i,quarterstart,quarterend,quarterstart,quarterstart,quarterstart,quarterend,i))   
    individual_results = c.fetchall() 
  
    file_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Template.xlsx'
    wb1 = xw.Book(file_path)
    main_sheet = wb1.sheets['Main']
    main_sheet.range('A2:H1000').value = ""
    main_sheet.range('A2').value = individual_results
    save_path = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//Individual//Delivered_vs_Closed_Won_{}.xlsx'.format(name)
    wb1.save(save_path)
    wb1.close     
    app.kill()

    Login = mysqlLogin()
    pw = Login[4]
    fromaddr = "kspang@addaptive.com"
    toaddr = email
    msg = MIMEMultipart() 
    msg['From'] = fromaddr   
    msg['To'] = toaddr
    msg['Subject'] = "Delivered vs. Closed Won {}".format(today)  
    body = "Please see the attached excel sheet for your {} Delivered vs. Closed Won report. ".format(today)  
    msg.attach(MIMEText(body, 'plain')) 
    filename = 'Delivered_vs_Closed_Won_{}.xlsx'.format(name)
    filepath = 'D://Users//kevinspang//Documents//Excel//Delivered vs Closed Won//individual//Delivered_vs_Closed_Won_{}.xlsx'.format(name)
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

cnx.close()   
