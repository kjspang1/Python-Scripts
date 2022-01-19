# -*- coding: utf-8 -*-
"""
Created on Mon Oct 26 10:12:09 2020

@author: kevinspang
"""

import datetime as dt
import mysql.connector
import xlwings as xw
import time
from login import mysqlLogin

#date logic
today = dt.datetime.strftime(dt.datetime.now(),'%Y_%m_%d')
x = dt.date.fromisoformat(dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d'))
qbegin = dt.date(x.year, 3 *((x.month - 1) // 3)+1, 1)
###################################################################################################################
app = xw.App(visible=False)  
#Open hubspot and extract emails and meetings date
crm = xw.Book('D://users//kevinspang//my documents//excel//sales reports//salesforce.xlsx')
salesforce = crm.sheets['salesforce']

#grab emails and meetings
meetings = salesforce.range('H2:I14').value
emails = salesforce.range('K2:L14').value
crm.close()

###################################################################################################################

#get RFP numbers
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
script = '''select fu.name, ifnull(sf.count,0) count from (
select id, concat(first_name," ",last_name) name
from fos_user
where id in ('1414','1837','2907','3099',
				'1079','1587','2617','2908','3162','3117',
				'2265','2473','3019')) fu     
left join                 
(select salesPerson_id, count(id) count
from sales_forecast 
where createdAt >= '{}'
group by salesPerson_id) sf
on sf.salesPerson_id = fu.id
order by 2 desc'''.format(qbegin)
c.execute(script)
rfp = c.fetchall()
c.close()

###################################################################################################################

#extract sales report data

try:
    sales = xw.Book('D://users//kevinspang//my documents//excel//Sales Reports//Master//Master//Sales_Report_Master_{}.xlsx'.format(today))
except:
    print('Issue with sales report master sheet')
sheet1 = sales.sheets['Jim Kappos']
sheet2 = sales.sheets['Chris Pedevillano']
sheet3 = sales.sheets['Dan Dechristoforo']
sheet4 = sales.sheets['Patrick McTague']
sheet5 = sales.sheets['Benjamin Crosby']
sheet6 = sales.sheets['Chris Ligas']
sheet7 = sales.sheets['Andrew Piekos']
sheet8 = sales.sheets['Joe Paratore']
sheet9 = sales.sheets["Morgan Afshari"]
sheet10 = sales.sheets["Mike Shea"]
sheet11 = sales.sheets['Alan Lai']
sheet12 = sales.sheets['Jefferson Butler']
sheet13 = sales.sheets['Adam Churilla']

sheet1_name = sheet1.range('B4').value
sheet1_cw = sheet1.range('B12').value
sheet1_od = sheet1.range('B21').value

if sheet1_od == 'N/A':
    sheet1_od = 9999999
else:
    pass

sheet2_name = sheet2.range('B4').value
sheet2_cw = sheet2.range('B12').value
sheet2_od = sheet2.range('B21').value

if sheet2_od == 'N/A':
    sheet2_od = 9999999
else:
    pass

sheet3_name = sheet3.range('B4').value
sheet3_cw = sheet3.range('B12').value
sheet3_od = sheet3.range('B21').value

if sheet3_od == 'N/A':
    sheet3_od = 9999999
else:
    pass

sheet4_name = sheet4.range('B4').value
sheet4_cw = sheet4.range('B12').value
sheet4_od = sheet4.range('B21').value

if sheet4_od == 'N/A':
    sheet4_od = 9999999
else:
    pass

sheet5_name = sheet5.range('B4').value
sheet5_cw = sheet5.range('B12').value
sheet5_od = sheet5.range('B21').value

if sheet5_od == 'N/A':
    sheet5_od = 9999999
else:
    pass

sheet6_name = sheet6.range('B4').value
sheet6_cw = sheet6.range('B12').value
sheet6_od = sheet6.range('B21').value

if sheet6_od == 'N/A':
    sheet6_od = 9999999
else:
    pass

sheet7_name = sheet7.range('B4').value
sheet7_cw = sheet7.range('B12').value
sheet7_od = sheet7.range('B21').value

if sheet7_od == 'N/A':
    sheet7_od = 9999999
else:
    pass

sheet8_name = sheet8.range('B4').value
sheet8_cw = sheet8.range('B12').value
sheet8_od = sheet8.range('B21').value

if sheet8_od == 'N/A':
    sheet8_od = 9999999
else:
    pass

sheet9_name = sheet9.range('B4').value
sheet9_cw = sheet9.range('B12').value
sheet9_od = sheet9.range('B21').value

if sheet9_od == 'N/A':
    sheet9_od = 9999999
else:
    pass
    
sheet10_name = sheet10.range('B4').value
sheet10_cw = sheet10.range('B12').value
sheet10_od = sheet10.range('B21').value

if sheet10_od == 'N/A':
    sheet10_od = 9999999
else:
    pass

sheet11_name = sheet11.range('B4').value
sheet11_cw = sheet11.range('B12').value
sheet11_od = sheet11.range('B21').value

if sheet11_od == 'N/A':
    sheet11_od = 9999999
else:
    pass

sheet12_name = sheet12.range('B4').value
sheet12_cw = sheet12.range('B12').value
sheet12_od = sheet12.range('B21').value

if sheet12_od == 'N/A':
    sheet12_od = 9999999
else:
    pass

sheet13_name = sheet13.range('B4').value
sheet13_cw = sheet13.range('B12').value
sheet13_od = sheet13.range('B21').value

if sheet13_od == 'N/A':
    sheet13_od = 9999999
else:
    pass


cw_list = [[sheet1_name,sheet1_cw],[sheet2_name,sheet2_cw],[sheet3_name,sheet3_cw],
            [sheet4_name,sheet4_cw],[sheet5_name,sheet5_cw],[sheet6_name,sheet6_cw],
            [sheet7_name,sheet7_cw],[sheet8_name,sheet8_cw],[sheet9_name,sheet9_cw],
            [sheet10_name,sheet10_cw],[sheet11_name,sheet11_cw],[sheet12_name,sheet12_cw],
            [sheet13_name,sheet13_cw]]

od_list = [[sheet1_name,sheet1_od],[sheet2_name,sheet2_od],[sheet3_name,sheet3_od],
            [sheet4_name,sheet4_od],[sheet5_name,sheet5_od],[sheet6_name,sheet6_od],
            [sheet7_name,sheet7_od],[sheet8_name,sheet8_od],[sheet9_name,sheet9_od],
            [sheet10_name,sheet10_od],[sheet11_name,sheet11_od],[sheet12_name,sheet12_od],
            [sheet13_name,sheet13_od]]

sales.close()



###################################################################################################################

#open template
leaderboard = xw.Book('D://users//kevinspang//my documents//excel//Sales Reports//leaderboard//leaderboard_template.xlsx')
main = leaderboard.sheets['Leaderboard']


p1 = [13,12,11,10,9,8,7,6,5,4,3,2,1]
p2 = [i * 2 for i in p1]
p3 = [i * 3 for i in p1]
p8 = [i * 8 for i in p1]  
p0 = [i * 0 for i in p1] 

meetings.sort(key = lambda x: x[1], reverse = True)
emails.sort(key = lambda x: x[1], reverse = True)
cw_list.sort(key = lambda x: x[1], reverse = True)
od_list.sort(key = lambda x: x[1], reverse = False)

#load in data
main.range('A2').value = meetings
main.range('E2').value = emails
main.range('I2').value = rfp
main.range('M2').value = cw_list
main.range('Q2').value = od_list

#load in points
main.range('C2').options(transpose=True).value = p2
main.range('G2').options(transpose=True).value = p1
main.range('K2').options(transpose=True).value = p3
main.range('O2').options(transpose=True).value = p8
main.range('S2').options(transpose=True).value = p0

time.sleep(3)

total = main.range('V2:W14').value
total.sort(key = lambda x: x[1], reverse = True)
main.range('V2').value = total

for i in range(2,15):
    if main.range('R{}'.format(i)).value == 9999999:
        main.range('R{}'.format(i)).value = 'N/A'
    else:
        pass

#save sheet in archive
leaderboard.save('D://users//kevinspang//my documents//excel//Sales Reports//leaderboard//Leaderboard_{}.xlsx'.format(today))
allData = main.range('A2:X14').value
leaderboard.close()

###################################################################################################################

#load data into master sheet
masterBook = xw.Book('D://users//kevinspang//my documents//excel//sales reports//master//master//Sales_Report_Master_{}.xlsx'.format(today))
mastersheet = masterBook.sheets['Leaderboard']
mastersheet.range('A2').value = allData
masterBook.save('D://users//kevinspang//my documents//excel//sales reports//master//master//Sales_Report_Master_{}.xlsx'.format(today))
masterBook.close()

#load to mikes sheet
mikeBook = xw.Book('D://users//kevinspang//my documents//excel//sales reports//Mike Powell//Mike Powell//Sales_Report_Mike_Powell_{}.xlsx'.format(today))
mikesheet = mikeBook.sheets['Leaderboard']
mikesheet.range('A2').value = allData
mikeBook.save('D://users//kevinspang//my documents//excel//sales reports//Mike Powell//Mike Powell//Sales_Report_Mike_Powell_{}.xlsx'.format(today))  
mikeBook.close()

#load to jacksons sheet
jBook = xw.Book('D://users//kevinspang//my documents//excel//sales reports//Jackson Davis//Jackson Davis//Sales_Report_Jackson_Davis_{}.xlsx'.format(today))
jsheet = jBook.sheets['Leaderboard']
jsheet.range('A2').value = allData
jBook.save('D://users//kevinspang//my documents//excel//sales reports//Jackson Davis//Jackson Davis//Sales_Report_Jackson_Davis_{}.xlsx'.format(today))  
jBook.close()

#load to teds sheet
tBook = xw.Book('D://users//kevinspang//my documents//excel//sales reports//Ted McNulty//Ted McNulty//Sales_Report_Ted_McNulty_{}.xlsx'.format(today))
tsheet = tBook.sheets['Leaderboard']
tsheet.range('A2').value = allData
tBook.save('D://users//kevinspang//my documents//excel//sales reports//Ted McNulty//Ted McNulty//Sales_Report_Ted_McNulty_{}.xlsx'.format(today))  
tBook.close()

app.kill()

    































