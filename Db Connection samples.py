# -*- coding: utf-8 -*-
"""
Created on Wed Mar  1 13:44:22 2023

@author: Kevin Spang
"""

from Login import snowflakeLogin
from Login import mysqlLogin
import mysql.connector

import pandas as pd

from snowflake.sqlalchemy import URL
from sqlalchemy import create_engine
from sqlalchemy.dialects import registry


#assign for mysql
mysqlLog = mysqlLogin()
mysql_user = mysqlLog[0]
mysql_password = mysqlLog[1]
mysql_db = mysqlLog[2]
mysql_host = mysqlLog[3]

#MySQL connection
cnx = mysql.connector.connect(user = mysql_user,
                              password = mysql_password,
                              db = mysql_db,
                              host = mysql_host
                              )

#######################################################################################################
#Grab Platform Data
c = cnx.cursor()

c.execute('''#executive dash closed won 
SELECT date_format(sysdate(), '%Y-%m-%d') date, a.quarter, a.closed_won, b.rfp, c.forecast FROM
    (SELECT 'Q1' quarter, ROUND(SUM(budget), 2) closed_won
    FROM sales_forecast_daily_budget
    WHERE day BETWEEN '2023-01-01' AND '2023-03-31'
	AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 4 )
	UNION ALL 
	SELECT 'Q2' quarter, ROUND(SUM(budget), 2)
    FROM sales_forecast_daily_budget
    WHERE day BETWEEN '2023-04-01' AND '2023-06-30'
	AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 4 ) 
    UNION ALL 
    SELECT 'Q3' quarter, ROUND(SUM(budget), 2)
    FROM sales_forecast_daily_budget
    WHERE day BETWEEN '2023-07-01' AND '2023-09-30'
	AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 4 ) 
	UNION ALL 
    SELECT 'Q4' quarter, ROUND(SUM(budget), 2)
    FROM sales_forecast_daily_budget
    WHERE day BETWEEN '2023-10-01' AND '2023-12-31'
	AND salesForecast_id IN (SELECT id FROM sales_forecast
	WHERE deal_stage = 4 )
    UNION ALL 
    SELECT 'Total' quarter, ROUND(SUM(budget), 2)
    FROM sales_forecast_daily_budget
    WHERE day BETWEEN '2023-01-01' AND '2023-12-31'
	AND salesForecast_id IN (SELECT id FROM sales_forecast
	WHERE deal_stage = 4 )
) a,
    
    (select  'Q1' quarter, ifnull(round(a.rfp + b.mni_rfp,2),0) rfp from
		(SELECT IFNULL(SUM(budget) * 4, 0) rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-01-01' AND '2023-03-31'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id <> '112')) a,
		(SELECT IFNULL(SUM(budget) * 4, 0) mni_rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-01-01' AND '2023-03-31'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id = '112' and id in (
		Select salesForecast_id
		from sales_forecast_audit_log 
		where message like '%"startDate":%"old":%'
		group by salesForecast_id
		having round(count(salesForecast_id)/ (count(distinct salesForecastSubdeal_id)+1),0) < 2)))b
		union all
		select  'Q2' quarter, ifnull(round(a.rfp + b.mni_rfp,2),0) rfp from
		(SELECT IFNULL(SUM(budget) * 4, 0) rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-04-01' AND '2023-06-30'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id <> '112')) a,
		(SELECT IFNULL(SUM(budget) * 4, 0) mni_rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-04-01' AND '2023-06-30'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id = '112' and id in (
		Select salesForecast_id
		from sales_forecast_audit_log 
		where message like '%"startDate":%"old":%'
		group by salesForecast_id
		having round(count(salesForecast_id)/ (count(distinct salesForecastSubdeal_id)+1),0) < 2)))b
		union all
		select  'Q3' quarter, ifnull(round(a.rfp + b.mni_rfp,2),0) rfp from
		(SELECT IFNULL(SUM(budget) * 4, 0) rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-07-01' AND '2023-09-30'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id <> '112')) a,
		(SELECT IFNULL(SUM(budget) * 4, 0) mni_rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-07-01' AND '2023-09-30'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id = '112' and id in (
		Select salesForecast_id
		from sales_forecast_audit_log 
		where message like '%"startDate":%"old":%'
		group by salesForecast_id
		having round(count(salesForecast_id)/ (count(distinct salesForecastSubdeal_id)+1),0) < 2)))b
		union all
		select  'Q4' quarter, ifnull(round(a.rfp + b.mni_rfp,2),0) rfp from
		(SELECT IFNULL(SUM(budget) * 4, 0) rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-10-01' AND '2023-12-31'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id <> '112')) a,
		(SELECT IFNULL(SUM(budget) * 4, 0) mni_rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-10-01' AND '2023-12-31'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id = '112' and id in (
		Select salesForecast_id
		from sales_forecast_audit_log 
		where message like '%"startDate":%"old":%'
		group by salesForecast_id
		having round(count(salesForecast_id)/ (count(distinct salesForecastSubdeal_id)+1),0) < 2)))b
        union all
		select  'Total' quarter, ifnull(round(a.rfp + b.mni_rfp,2),0) rfp from
		(SELECT IFNULL(SUM(budget) * 4, 0) rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-01-01' AND '2023-12-31'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id <> '112')) a,
		(SELECT IFNULL(SUM(budget) * 4, 0) mni_rfp
		FROM sales_forecast_daily_budget
		WHERE day BETWEEN '2023-01-01' AND '2023-12-31'
		AND salesForecast_id IN (SELECT id FROM sales_forecast WHERE deal_stage = 1
		and client_id = '112' and id in (
		Select salesForecast_id
		from sales_forecast_audit_log 
		where message like '%"startDate":%"old":%'
		group by salesForecast_id
		having round(count(salesForecast_id)/ (count(distinct salesForecastSubdeal_id)+1),0) < 2)))b
	) b,
    (Select 'Q1' quarter, 3777333.46 forecast
    union all
    Select 'Q2' quarter, 4542043.22 forecast
    union all
    Select 'Q3' quarter, 4426200.69 forecast
    union all
    Select 'Q4' quarter, 4387900.67 forecast
    union all
    Select 'Total' quarter, 17133478.04 forecast
    ) c
WHERE a.quarter = b.quarter
and c.quarter = a.quarter''')
print("Scripts Done")
results = c.fetchall()
col_list = list(c.column_names)
df = pd.DataFrame(results, columns = col_list)

##########################################################################################
#pipe into snowflake
sfLog = snowflakeLogin()
account = sfLog[0]
user = sfLog[1]
password = sfLog[2]
role = sfLog[3]
warehouse = sfLog[4]
database = sfLog[5]
schema = sfLog[6]

registry.register('snowflake', 'snowflake.sqlalchemy', 'dialect')

engine = create_engine(URL(
    account = account,
    user = user,
    password = password,
    database = database,
    schema = schema,
    warehouse = warehouse,
    role = role,
    autocommit = True
    ))

df.to_sql(con = engine,
          name = 'forecast_table_sample',
          if_exists = 'replace',
          index = False)
























