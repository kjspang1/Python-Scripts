# -*- coding: utf-8 -*-
"""
Created on Thu Nov 14 10:42:58 2019

@author: kevinspang
"""

import gspread 
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

account_data = client.open("AcctsDepartment_2019_V2")
accounts = account_data.get_worksheet(0)


eli_team_ids = accounts.acell('I23').value
jackie_team_ids = accounts.acell('J23').value#matt_team_ids = accounts.acell('K23').value
max_team_ids = accounts.acell('L23').value
unsupported_string = accounts.acell('L28').value
supported_string = accounts.acell('M28').value