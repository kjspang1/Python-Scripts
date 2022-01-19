# -*- coding: utf-8 -*-
"""
Created on Tue Aug 31 08:22:36 2021

@author: kevinspang
"""

import datetime as dt
import smtplib
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 
from login import mysqlLogin

Login = mysqlLogin()
pw = Login[4]

datefile = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
today = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')

fromaddr = "kspang@addaptive.com"

#Master
#emails = ["kspang@addaptive.com","kspang@addaptive.com","kspang@addaptive.com"]
emails = ["egaines@addaptive.com","emoger@addaptive.com","jellingson@addaptive.com"]
for i in emails:
    Login = mysqlLogin()
    pw = Login[4]
    datemessage = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')
    fromaddr = "kspang@addaptive.com"
    toaddr = i
    msg = MIMEMultipart() 
    msg['From'] = fromaddr   
    msg['To'] = toaddr
    msg['Subject'] = "Accounts Report {}".format(today)  
    body = "Please see the attached Excel sheet for the {} accounts report.".format(today) 
    msg.attach(MIMEText(body, 'plain')) 
    filename = 'Accounts_Report_Master_{}.xlsx'.format(datefile)
    filepath = 'D://users//kevinspang//my documents//excel//accounts reports//master//Accounts_Report_Master_{}.xlsx'.format(datefile)
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
    
#Hannah Lucas
Login = mysqlLogin()
pw = Login[4]
fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"
toaddr = "hlucas@addaptive.com"
msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Accounts Report {}".format(today)  
body = "Please see the attached Excel sheet for the {} accounts report.".format(today) 
msg.attach(MIMEText(body, 'plain')) 
filename = 'Accounts_Report_Hannah_Lucas_{}.xlsx'.format(datefile)
filepath = 'D://users//kevinspang//my documents//excel//accounts reports//hannah lucas//Accounts_Report_Hannah_Lucas_{}.xlsx'.format(datefile)
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

#Sami Kaminski
Login = mysqlLogin()
pw = Login[4]
fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"
toaddr = "skaminski@addaptive.com"
msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Accounts Report {}".format(today)  
body = "Please see the attached Excel sheet for the {} accounts report.".format(today) 
msg.attach(MIMEText(body, 'plain')) 
filename = 'Accounts_Report_Sami_Kaminski_{}.xlsx'.format(datefile)
filepath = 'D://users//kevinspang//my documents//excel//accounts reports//sami kaminski//Accounts_Report_Sami_Kaminski_{}.xlsx'.format(datefile)
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

#Joe Ditullio
Login = mysqlLogin()
pw = Login[4]
fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"
toaddr = "jditullio@addaptive.com"
msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Accounts Report {}".format(today)  
body = "Please see the attached Excel sheet for the {} accounts report.".format(today) 
msg.attach(MIMEText(body, 'plain')) 
filename = 'Accounts_Report_Joe_Ditullio_{}.xlsx'.format(datefile)
filepath = 'D://users//kevinspang//my documents//excel//accounts reports//Joe Ditullio//Accounts_Report_Joe_Ditullio_{}.xlsx'.format(datefile)
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

#Will Rice
Login = mysqlLogin()
pw = Login[4]
fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"
toaddr = "wrice@addaptive.com"
msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Accounts Report {}".format(today)  
body = "Please see the attached Excel sheet for the {} accounts report.".format(today) 
msg.attach(MIMEText(body, 'plain')) 
filename = 'Accounts_Report_Will_Rice_{}.xlsx'.format(datefile)
filepath = 'D://users//kevinspang//my documents//excel//accounts reports//Will Rice//Accounts_Report_Will_Rice_{}.xlsx'.format(datefile)
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