# -*- coding: utf-8 -*-
"""
Created on Thu Mar  4 13:01:14 2021

@author: kevinspang
"""

import requests 
import pandas as pd
import xlwings as xw

#extract NPI Codes
#wb1 = xw.Book('C://Users//Kevin_Spang//Documents//Python_Scripts//NPI Data extract//list.xlsx' )
#npi_numbers = wb1.sheets[0].range('A:A').value





NPI_num = ['1477556413']

company_list = {}
for item in NPI_num:
    try:
        NPI_data = requests.get(f'https://npiregistry.cms.hhs.gov/api/?version=2.0&number={item}')
        NPI_data = NPI_data.json()
        
        number = NPI_data.get('results', {})[0].get('number')
        name = NPI_data.get('results', {})[0].get('basic', {}).get('name')
        address = NPI_data.get('results', {})[0].get('addresses', {})[0].get('address_1')
        city = NPI_data.get('results', {})[0].get('addresses', {})[0].get('city')
        state = NPI_data.get('results', {})[0].get('addresses', {})[0].get('state')
        postalcode = NPI_data.get('results', {})[0].get('addresses', {})[0].get('postal_code')
        
        company_list[item] = {}
        company_list[item]['Number'] = number
        company_list[item]['Name'] = name
        company_list[item]['Address'] = address
        company_list[item]['City'] = city
        company_list[item]['State'] = state
        company_list[item]['Postal_Code'] = postalcode
        company_list.pop('result_count', None)
    except:
      pass

NPI_list = pd.DataFrame.from_dict(company_list, orient='index')
NPI_list.reset_index(drop = True, inplace = True )

#app = xw.App()
#w1 = xw.Book()
#sheet = w1.sheets[0]
#sheet.range('A1').value = NPI_list
#w1.save('C://Users//Kevin_Spang//Documents//Python_Scripts//NPI Data extract//NPI_Return.xlsx')
#w1.close()
#app.kill()