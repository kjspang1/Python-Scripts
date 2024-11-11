# -*- coding: utf-8 -*-
"""
Created on Thu Jun  6 15:39:21 2024

@author: Kevin Spang
"""

import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt

pd.options.display.float_format = '{:.0f}'.format

pd.options.mode.chained_assignment = None

from Login import snowflakeLogin
import snowflake.connector as sf

#########################################################################################################################################
sfLog = snowflakeLogin()

ctx = sf.connect(
    user = sfLog[1],
    password = sfLog[2],
    account = sfLog[0],
    role = sfLog[3],
    warehouse = sfLog[4],
    database = sfLog[5],
    schema = sfLog[6]
   )

cs = ctx.cursor()

#get meeting test data
try:
    cs.execute('''
                Select 
                dayname(to_date(date)) day,
                to_varchar(date, '%H') time,
                count(to_varchar(date, '%H')) freq
                from (
                    Select to_timestamp(to_varchar(dateadd(hour,-4, created_at), 'YYYY-MM-DD HH24:MI:SS')) date
                    from outreach.public.calls
                    where dayname(to_date(dateadd(hour,-4,dialed_at))) not in ('Sat','Sun')
                    and prospect_id is not null
                    ) a
                where to_varchar(date, '%H') not in (08,09,19,20)
                group by day, time;
                ''')
    calls_m = cs.fetch_pandas_all()
    
    cs.execute('''
                Select 
                dayname(to_date(date)) day,
                to_varchar(date, '%H') time,
                count(to_varchar(date, '%H')) freq
                from (
                    Select to_timestamp(to_varchar(dateadd(hour,-4, created_at), 'YYYY-MM-DD HH24:MI:SS')) date
                    from outreach.public.calls
                    where dayname(to_date(dateadd(hour,-4,dialed_at))) not in ('Sat','Sun')
                    and prospect_id is not null
                    and call_disposition_id in (1,2,3,4,5,10)
                    ) a
                where to_varchar(date, '%H') not in (08,09,19,20)
                group by day, time;
                ''')    
    calls_a = cs.fetch_pandas_all()
finally:
    cs.close()
ctx.close()    
#########################################################################################################################################
# Heat map of calls made
calls_made = calls_m.pivot(index="TIME",columns="DAY",values="FREQ")
calls_made  = calls_made.reindex(columns=['Mon','Tue','Wed','Thu','Fri'])
calls_made = calls_made.loc[::-1]
calls_made = calls_made.rename(index={"08":"8-9am",
                                      "09":"9-10am",
                                      "10":"10-11am",
                                      "11":"11-12am",
                                      "12":"12-1pm",
                                      "13":"1-2pm",
                                      "14":"2-3pm",
                                      "15":"3-4pm",
                                      "16":"4-5pm",
                                      "17":"5-6pm",
                                      "18":"6-7pm",
                                      "19":"7-8pm",
                                      "20":"8-9pm"
                                          })

calls_answered = calls_a.pivot(index="TIME",columns="DAY",values="FREQ")
calls_answered  = calls_answered.reindex(columns=['Mon','Tue','Wed','Thu','Fri'])
calls_answered = calls_answered.loc[::-1]
calls_answered = calls_answered.rename(index={"08":"8-9am",
                                      "09":"9-10am",
                                      "10":"10-11am",
                                      "11":"11-12am",
                                      "12":"12-1pm",
                                      "13":"1-2pm",
                                      "14":"2-3pm",
                                      "15":"3-4pm",
                                      "16":"4-5pm",
                                      "17":"5-6pm",
                                      "18":"6-7pm",
                                      "19":"7-8pm",
                                      "20":"8-9pm"
                                          })

difference = calls_made.subtract(calls_answered, fill_value=0.0)
ratio = difference.div(calls_made, axis=0, fill_value=0.0)
win_pct = calls_answered.div(calls_made, axis = 0, fill_value = 0.0)

#format library https://docs.python.org/3/library/string.html#formatspec

f, ax = plt.subplots(2,2, figsize=(20,10))
ax1 = sns.heatmap(calls_made, ax=ax[0,0], annot=True, cmap="coolwarm", fmt=".0f", linewidths=1)
ax2 = sns.heatmap(calls_answered, ax=ax[0,1], annot=True, cmap="coolwarm", fmt=".0f", linewidths=1)
ax3 = sns.heatmap(win_pct, ax=ax[1,0], annot=True, cmap="coolwarm", fmt=".2%", linewidths=1)
#ax4 = sns.heatmap(ratio, ax=ax[1,1], annot=True, cmap="coolwarm", fmt=".2%", linewidths=1)
ax1.set_title("Calls Made")
ax2.set_title("Calls Answered")
ax3.set_title("Win Percentage")
#ax4.set_title("Ratio")
plt.show()




# f, ax = plt.subplots(1,2, figsize=(15,6))
# ax5 = sns.heatmap(ratio, ax=ax[0], annot=True, cmap="coolwarm", fmt=".2%")
# ax6 = sns.heatmap(win_pct, ax=ax[1], annot=True, cmap="coolwarm", fmt=".2%")
# ax5.set_title("Ratio")
# ax6.set_title("Win Percentage")
# plt.show()









































    