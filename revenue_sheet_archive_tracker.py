# -*- coding: utf-8 -*-
"""
Created on Mon Jan 13 11:06:04 2020

@author: kevinspang
"""
import datetime as dt
import xlwings as xw
import pandas as pd
import gspread 
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('D:\\Users\\kevinspang\\Documents\\Python_Scripts\\client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open("2020 Revenue")

revenue_data_2020 = sheet.get_worksheet(1)
revenue_data_2019 = sheet.get_worksheet(3)
revenue_data_2018 = sheet.get_worksheet(5)
revenue_data_2017 = sheet.get_worksheet(8)
revenue_data_2016 = sheet.get_worksheet(10)
revenue_data_2015 = sheet.get_worksheet(12)
revenue_data_2014 = sheet.get_worksheet(13)

total_revenue_data_2020 = revenue_data_2020.get_all_values()
total_revenue_data_2019 = revenue_data_2019.get_all_values()
total_revenue_data_2018 = revenue_data_2018.get_all_values()
total_revenue_data_2017 = revenue_data_2017.get_all_values()
total_revenue_data_2016 = revenue_data_2016.get_all_values()
total_revenue_data_2015 = revenue_data_2015.get_all_values()
total_revenue_data_2014 = revenue_data_2014.get_all_values()

headers_2020 = total_revenue_data_2020.pop(0)
headers_2019 = total_revenue_data_2019.pop(0)
headers_2018 = total_revenue_data_2018.pop(0)
headers_2017 = total_revenue_data_2017.pop(0)
headers_2016 = total_revenue_data_2016.pop(0)
headers_2015 = total_revenue_data_2015.pop(0)
headers_2014 = total_revenue_data_2014.pop(0)

df_revenue_sheet_2020 = pd.DataFrame(total_revenue_data_2020, columns = headers_2020)
df_revenue_sheet_2019 = pd.DataFrame(total_revenue_data_2019, columns = headers_2019)
df_revenue_sheet_2018 = pd.DataFrame(total_revenue_data_2018, columns = headers_2018)
df_revenue_sheet_2017 = pd.DataFrame(total_revenue_data_2017, columns = headers_2017)
df_revenue_sheet_2016 = pd.DataFrame(total_revenue_data_2016, columns = headers_2016)
df_revenue_sheet_2015 = pd.DataFrame(total_revenue_data_2015, columns = headers_2015)
df_revenue_sheet_2014 = pd.DataFrame(total_revenue_data_2014, columns = headers_2014)

app = xw.App()
wb1 = xw.Book('D://users//kevinspang//documents//excel//revenue sheet archive//Revenue_temp.xlsx' )
    
sheet_2020 = wb1.sheets['2020']
sheet_2019 = wb1.sheets['2019']
sheet_2018 = wb1.sheets['2018']
sheet_2017 = wb1.sheets['2017']
sheet_2016 = wb1.sheets["2016"]
sheet_2015 = wb1.sheets["2015"]
sheet_2014 = wb1.sheets["2014"]

sheet_2020.range('A1').options(index = False).value = df_revenue_sheet_2020
sheet_2019.range('A1').options(index = False).value = df_revenue_sheet_2019
sheet_2018.range('A1').options(index = False).value = df_revenue_sheet_2018
sheet_2017.range('A1').options(index = False).value = df_revenue_sheet_2017
sheet_2016.range('A1').options(index = False).value = df_revenue_sheet_2016
sheet_2015.range('A1').options(index = False).value = df_revenue_sheet_2015
sheet_2014.range('A1').options(index = False).value = df_revenue_sheet_2014

today = dt.datetime.strftime(dt.datetime.now(),"%Y_%m_%d")
save_path = 'D://users//kevinspang//documents//excel//revenue sheet archive//Revenue_{}.xlsx'.format(today)
wb1.save(save_path)
wb1.close()
app.kill()