# -*- coding: utf-8 -*-
"""
Created on Thu Oct  8 08:32:28 2020

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

date = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')

cnx = mysql.connector.connect(user = user,
                              password = password,
                              db = db,
                              host = host)

c = cnx.cursor()
c.execute('''Select distinct li.client_id, cl.name, cl.type, date_format(max(end_time),'%Y-%m-%d') last_campaign_end_date, 
            am.user_id, concat(fu.first_name," ",fu.last_name) account_manager
            from line_item li
            join client cl
            on li.client_id = cl.id
            join account_manager am
            on am.client_id = li.client_id
            join fos_user fu 
            on fu.id = am.user_id 
            where am.user_id is not null
            group by client_id
            having max(end_time) between '2017-01-01' 
            and date_format(date_sub(sysdate(), Interval 12 Month), '%Y-%m-%d')
            order by 4 desc''')
results = c.fetchall()
cnx.close()


app = xw.App()
file_path = 'D://users//kevinspang//my documents//excel//Inactive Accounts//Inactive Accounts.xlsx'
wb1 = xw.Book(file_path)
main_sheet = wb1.sheets['Main']
main_sheet.range('A2:F1000').value = ""
main_sheet.range('A2').value = results
date = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')
wb1.save('D://users//kevinspang//my documents//excel//Inactive Accounts//Inactive Accounts {}.xlsx'.format(date))
wb1.close
app.kill()

pw = Login[4]
fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"
toaddr = "manageops@addaptive.com"

msg = MIMEMultipart() 
msg['From'] = fromaddr   
msg['To'] = toaddr
msg['Subject'] = "Inactive Accounts {}".format(date)  
body = "Please see the attached excel sheet for the {} inactive accounts list".format(date)  

msg.attach(MIMEText(body, 'plain')) 
filename = 'Inactive Accounts {}.xlsx'.format(date)
filepath = 'D://users//kevinspang//my documents//excel//Inactive Accounts//Inactive Accounts {}.xlsx'.format(date)
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
