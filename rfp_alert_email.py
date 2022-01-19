# -*- coding: utf-8 -*-
"""
Created on Thu Dec 12 10:43:57 2019

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

emails = ["hgraham@addaptive.com","pmctague@addaptive.com","mafshari@addaptive.com","bcrosby@addaptive.com",
              "mshea@addaptive.com","cligas@addaptive.com","mmodoono@addaptive.com"]

c = cnx.cursor()
c.execute('''Select cl.name, concat(fu.first_name, " ", fu.last_name) Sales_Person, 
                sf.advertiser_name, sf.start_date, sf.end_date, sf.budget,
                case 
                when deal_stage = 1 then 'RFP'
                when deal_stage = 7 then 'RFI' 
                when deal_stage = 2 then  'Negotiating'
                when deal_stage = 3 then 'Contract Sent'
                end as deal_stage
                from sales_forecast sf
                join client cl
                on cl.id = sf.client_id
                join fos_user fu
                on fu.id = sf.salesPerson_id 
                where date_format(sf.createdAt, '%Y-%m-%d') = 
                date_format(subdate(sysdate(), INTERVAL 1 Day), '%Y-%m-%d')
                and sf.deal_stage in (1,2,3,7)''')
results = c.fetchall()
cnx.close()

if len(results) == 0:
    print('No New RFPs')
else:    
    app = xw.App(visible=False)  
    file_path = 'D://users//kevinspang//my documents//excel//RFP Alert//New RFPs.xlsx'
    wb1 = xw.Book(file_path)
    main_sheet = wb1.sheets['Main']
    main_sheet.range('A2:G1000').value = ""
    main_sheet.range('A2').value = results
    wb1.save(file_path)
    wb1.close
    app.kill()
    
    for i in emails:
        Login = mysqlLogin()
        pw = Login[4]
        datemessage = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')
        fromaddr = "kspang@addaptive.com"
        toaddr = i
        msg = MIMEMultipart() 
        msg['From'] = fromaddr   
        msg['To'] = toaddr
        msg['Subject'] = "New RFP List {}".format(datemessage)  
        body = "Please see the attached excel sheet for the list of new RFPs.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        filename = 'New RFPs.xlsx'
        filepath = 'D://users//kevinspang//my documents//excel//RFP Alert//New RFPs.xlsx'
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
