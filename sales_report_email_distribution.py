# -*- coding: utf-8 -*-
"""
Created on Wed Feb 19 10:46:19 2020

@author: kevinspang
"""

import datetime as dt
import os.path
import smtplib 
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.application import MIMEApplication
from login import mysqlLogin

Login = mysqlLogin()
pw = Login[4]

datefile = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
datemessage = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d')

fromaddr = "kspang@addaptive.com"
#toaddr = "kspang@addaptive.com"

snapshot = 'Sales_Snapshot_{}.xlsx'.format(datefile)
weekly = 'Weekly_Sales_{}.xlsx'.format(datefile)

############################################################################################################################
def mike_team_run():
    #email to mike powell and team
    #Mike Powell
    try:
        toaddr = "mpowell@addap▼tive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//Mike Powell//Mike Powell"
        files = ["Sales_Report_Mike_Powell_{}.xlsx".format(datefile),"Mike Team Weekly Report.xlsx"]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Mike Powell")
    
    #Jim Kappos    
    try:
        toaddr = "jkappos@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//Mike Powell//Jim kappos"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Jim kappos")
    
    #Chris Pedevillano    
    try:
        toaddr = "cpedevillano@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//Mike Powell//chris pedevillano"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Chris Pedevillano")
        
    
    #Dan Dechristoforo   
    try:
        toaddr = "ddechristoforo@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//Mike Powell//dan dechristoforo"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Dan Dechristoforo ")
        
    #Patrick McTague 
    try:
        toaddr = "pmctague@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//Mike Powell//patrick mctague"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Patrick McTague ")
        
    #Benjamin Crosby 
    try:
        toaddr = "bcrosby@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//Mike Powell//benjamin crosby"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Benjamin Crosby ")
        
    #Chris Ligas
    try:
        toaddr = "cligas@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//Mike Powell//chris ligas"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Chris Ligas ")
        
    
######################################################################################################################
#email to Jackson Davis and team
#Jackson Davis
def jackson_team_run():
    try:
        toaddr = "jdavis@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//jackson davis//jackson davis"
        files = ["Sales_Report_Jackson_Davis_{}.xlsx".format(datefile),"Jackson Team Weekly Report.xlsx"]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Jackson Davis")
    
    #Andrew Piekos    
    try:
        toaddr = "apiekos@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//jackson davis//andrew piekos"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Andrew Piekos")
    
    #Joe Paratore    
    try:
        toaddr = "jparatore@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//jackson davis//joe paratore"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Joe Paratore")
        
    #Morgan Afshari 
    try:
        toaddr = "mafshari@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//jackson davis//morgan afshari"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Morgan Afshari")
        
    #Mike Shea 
    try:
        toaddr = "mshea@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//jackson davis//mike shea"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Mike Shea")

######################################################################################################################
def ted_team_run():
    #Email out to ted's team (Ted gets master not team sheet)
    
    #Alan Lai   
    try:
        toaddr = "alai@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//tim pedersen//alan lai"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Alan Lai")
    
    #Jefferson  Butler 
    try:
        toaddr = "jbutler@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//tim pedersen//jefferson butler"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Jefferson Butler")
        
    #Adam Churilla
    try:
        toaddr = "achurilla@addaptive.com"
        dir_path = "D://users//kevinspang//my documents//excel//sales reports//tim pedersen//adam churilla"
        files = [snapshot, weekly]
        
        msg = MIMEMultipart() 
        msg['To'] = toaddr
        msg['From'] = fromaddr   
        msg['Subject'] = "Sales Report {}".format(datemessage)  
        body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
        msg.attach(MIMEText(body, 'plain')) 
        
        for f in files:
            file_path = os.path.join(dir_path, f)
            attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
            attachment.add_header('Content-Disposition','attachment', filename=f)
            msg.attach(attachment) 
          
        # creates SMTP session 
        s = smtplib.SMTP('smtp.gmail.com', 587) 
        s.starttls() 
        s.login(fromaddr, pw) 
        text = msg.as_string() 
        s.sendmail(fromaddr, toaddr, text) 
        s.quit() 
    except:
        print("Unable to send email to Adam Churilla")
######################################################################################################################
#Send Sheets to management
def master_team_run():
    emails = ["komalley@addaptive.com","pshea@addaptive.com","tmcnulty@addaptive.com","hgraham@addaptive.com","mmahoney@addaptive.com"]
    for i in emails:
        try:
            toaddr = i
            dir_path = "D://users//kevinspang//my documents//excel//sales reports//master//master"
            files = ["Sales_Report_Master_{}.xlsx".format(datefile)]
            
            msg = MIMEMultipart() 
            msg['To'] = toaddr
            msg['From'] = fromaddr   
            msg['Subject'] = "Sales Report {}".format(datemessage)  
            body = "Please see the attached Excel sheets for the {} weekly sales reports.".format(datemessage)  
            msg.attach(MIMEText(body, 'plain')) 
            
            for f in files:
                file_path = os.path.join(dir_path, f)
                attachment = MIMEApplication(open(file_path, "rb").read(), _subtype="txt")
                attachment.add_header('Content-Disposition','attachment', filename=f)
                msg.attach(attachment) 
              
            # creates SMTP session 
            s = smtplib.SMTP('smtp.gmail.com', 587) 
            s.starttls() 
            s.login(fromaddr, pw) 
            text = msg.as_string() 
            s.sendmail(fromaddr, toaddr, text) 
            s.quit() 
        except:
            print("Unable to send master email")     
        
   
######################################################################################################################            
mike_team_run()
jackson_team_run()
ted_team_run()
master_team_run()        















