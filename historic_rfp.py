# -*- coding: utf-8 -*-
"""
Created on Mon Dec 28 08:54:56 2020

@author: kevinspang
"""

#Returns the list of all rfps 11 months ago

import datetime as dt
from dateutil.relativedelta import relativedelta
import mysql.connector
import xlwings as xw
from login import mysqlLogin
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

#date logic
#today = dt.datetime.strftime(dt.datetime.now(),'%b %Y')
delta = relativedelta(months=11)
year = dt.datetime.strftime(dt.datetime.now() - delta,'%Y')
today = dt.datetime.strftime(dt.datetime.now(),'%m')

if today == '10':
    date = 'Q1 ' + year 
elif today == '01':
    date = 'Q2 ' + year 
elif today == '04':
    date = 'Q3 ' + year 
elif today == '07':
    date = 'Q4 ' + year 

#run script for rfp results
Login = mysqlLogin()
user = Login[0]
password = Login[1]
db = Login[2]
host = Login[3]

cnx = mysql.connector.connect(user = user,
                              password = password,
                              db = db,
                              host = host)
c = cnx.cursor()
script = '''select a.*, 
ifnull(cw1.flag," ") closed_won_flag,
ifnull(cw2.flag," ") cw_advertiser_flag,
ifnull(rfp1.flag," ") rfp_flag,
ifnull(rfp2.flag," ") rfp_advertiser_flag
from
(select sf.client_id,
cl.name agency,
sf.advertiser_name, 
case 
when concat(fu.first_name," ",fu.last_name) = 'Andrew Fox' then 'Former rep'
when concat(fu.first_name," ",fu.last_name) = 'Angie Waters' then 'Former rep'
when concat(fu.first_name," ",fu.last_name) = 'Tyler Meyer' then 'Former rep'
else concat(fu.first_name," ",fu.last_name)
end salesperson,
case 
when deal_stage = '1' then 'RFP'  
when deal_stage = '2' then 'Negotiating'  
when deal_stage = '3' then 'Contract Sent'  
when deal_stage = '4' then 'Closed Won'  
when deal_stage = '5' then 'Closed Lost'  
when deal_stage = '6' then 'Expired'  
when deal_stage = '7' then 'RFI'  
when deal_stage = '8' then 'Winback'  
else 'Other'
End deal_stage,
sf.impression_goal,
sf.cpm,
sf.budget,
start_date,
end_date
from sales_forecast sf
join client cl
on sf.client_id = cl.id
join fos_user fu
on sf.salesPerson_id = fu.id
where sf.client_id not in ('112','299','62') 
and createdAt between
date_add(DATE_FORMAT(CURDATE(),'%Y-%m-01'), interval -9 Month )
and 
date_add(LAST_DAY(CURDATE()), interval -7 Month)
) a

#client has a closed won campaign in the system
left join
(select distinct client_id,
Case
when deal_stage = '4' then 'X'
end flag
from sales_forecast
where end_date >= DATE_FORMAT(curdate() ,'%Y-%m-01')
and deal_stage = 4) cw1
on a.client_id = cw1.client_id

#client is closed won with same advertiser
left join
(select distinct client_id,
advertiser_name,
Case
when deal_stage = '4' then 'X'
end flag
from sales_forecast
where end_date >= DATE_FORMAT(curdate() ,'%Y-%m-01')
and deal_stage = 4) cw2
on a.client_id = cw2.client_id
and a.advertiser_name = cw2.advertiser_name

#Client has any rfp in the system
left join
(select distinct client_id,
Case
when deal_stage = '1' then 'X'
end flag
from sales_forecast
where end_date >= DATE_FORMAT(curdate() ,'%Y-%m-01')
and deal_stage = 1) rfp1
on a.client_id = rfp1.client_id

#client has an rfp with the same advertiser in the system
left join
(select distinct client_id,
advertiser_name,
Case
when deal_stage = '1' then 'X'
end flag
from sales_forecast
where end_date >= DATE_FORMAT(curdate() ,'%Y-%m-01')
and deal_stage = 1) rfp2
on a.client_id = rfp2.client_id
and a.advertiser_name = rfp2.advertiser_name
order by 4,2,3'''
c.execute(script)
rfp = c.fetchall()
c.close()


######################################################################################
#load into excel

app = xw.App(visible=False)  
#Open hubspot and extract emails and meetings date
historic_rfp = xw.Book('D://users//kevinspang//my documents//excel//historic rfp//historic rfp.xlsx')
main = historic_rfp.sheets['RFP List']

main.range('A2').value = rfp

file_path = 'D://users//kevinspang//my documents//excel//historic rfp//Historic RFP - {}.xlsx'.format(date)

try:
    historic_rfp.save(file_path)
    historic_rfp.close()
except:
    historic_rfp.close()
    print("Sheet already saved")
app.kill()

############################################################################################################################################
#Email Results to Ashley
pw = Login[4]
emails = ["atrudeau@addaptive.com","mmahoney@addaptive.com"]

for i in emails:
    fromaddr = "kspang@addaptive.com"
    toaddr = i
       
    #message creation and attachment.
    msg = MIMEMultipart() 
    msg['From'] = fromaddr   
    msg['To'] = toaddr
    msg['Subject'] = "Historic RFP {}".format(date)  
    body = "Please see the attached Excel sheet for the {} Historic RFP list.".format(date)  
    msg.attach(MIMEText(body, 'plain')) 
    filename = "Historic RFP - {}.xlsx".format(date)
    filepath = "'D://users//kevinspang//my documents//excel//historic rfp//{}".format(filename)
    attachment = open(file_path, "rb") 
      
    # instance of MIMEBase and named as p 
    p = MIMEBase('application', 'octet-stream') 
    p.set_payload((attachment).read()) 
    encoders.encode_base64(p) 
    p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
    msg.attach(p) 
      
    # creates SMTP session 
    s = smtplib.SMTP('smtp.gmail.com', 587) 
    s.starttls() 
    s.login(fromaddr, pw) 
    text = msg.as_string() 
    s.sendmail(fromaddr, toaddr, text) 
    s.quit()  




