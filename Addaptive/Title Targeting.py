# -*- coding: utf-8 -*-
"""
Created on Fri May 31 08:21:42 2024

@author: Kevin Spang
"""

import pandas as pd
pd.options.mode.chained_assignment = None
import datetime as dt
from Login import snowflakeLogin
import snowflake.connector as sf
import string 
from nltk.corpus import stopwords
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.feature_extraction.text import TfidfTransformer
from sklearn.naive_bayes import MultinomialNB
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.metrics import classification_report,confusion_matrix
from sklearn.ensemble import RandomForestClassifier
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
               Select 1 had_meeting, 
               replace(replace(lower(title),'vice president','vp'),'senior','sr') title 
               from (
               Select distinct att.prospect_id, pro.title 
               from outreach.public.meeting_attendees att
               join outreach.public.prospects pro
               on pro.id = att.prospect_id
               where prospect_id is not null
               and invite_response = 'ACCEPTED'
               and title is not null)
               union all
               Select 0 had_meeting, 
               replace(replace(lower(title),'vice president','vp'),'senior','sr') title 
               from outreach.public.prospects pro
               where pro.touched_at is not null
               and pro.title is not null
               and pro.owner_id <> 21
               and pro.id not in (
                   Select distinct att.prospect_id
                   from outreach.public.meeting_attendees att
                   where prospect_id is not null
                   and invite_response = 'ACCEPTED'
               )
            ''')
                                
    titles = cs.fetch_pandas_all()
    cs.execute('''
                Select acc.name agency,
                concat(user.first_name,' ',user.last_name) current_sales_rep,
                pro.id prospect_id,
                concat(pro.first_name,' ',pro.last_name) prospect_name,
                replace(replace(lower(pro.title),'vice president','vp'),'senior','sr') title,
                acc.industry, 
                acc.size agency_size
                from outreach.public.prospects pro
                join outreach.public.accounts acc
                on acc.id = pro.account_id
                join outreach.public.users user
                on acc.owner_id = user.id
                where pro.title is not null 
                and pro.owner_id = 21
                and pro.opted_out_at is null
                and acc.id not in ('4155')
            ''')
    prospects = cs.fetch_pandas_all()
finally:
    cs.close()
ctx.close()    
    


#########################################################################################################################################

#titles = pd.read_csv('C:/Users/Kevin Spang/OneDrive - DataPoint Media/Documents/Python/Addaptive/Title Target NLM.csv')
#titles.head()

def text_process(mess):
    nopunc = [char for char in mess if char not in string.punctuation]
    nopunc = ''.join(nopunc)
    return [word for word in nopunc.split() if word.lower() not in stopwords.words('english')]

titles['TITLE'].apply(text_process)


#Create bag of words matrix 
bow_transformer = CountVectorizer(analyzer=text_process).fit(titles['TITLE'])
titles_bow = bow_transformer.transform(titles['TITLE'])
tfidf_transformer = TfidfTransformer().fit(titles_bow)
titles_tfidf = tfidf_transformer.transform(titles_bow)
title_cat = MultinomialNB().fit(titles_tfidf,titles['HAD_MEETING'])
all_pred = title_cat.predict(titles_tfidf)

################################################################################################################################################
#Need to deal with unbalanced data
#Train model
no_meetings = titles[titles['HAD_MEETING']==0]
yes_meetings = titles[titles['HAD_MEETING']==1]

count = int(round(no_meetings.shape[0]/yes_meetings.shape[0]/3,0))
yes_meetings_large = pd.DataFrame()

while count > 0:
    yes_meetings_large = pd.concat([yes_meetings_large,yes_meetings])
    count = count-1

titles = pd.concat([yes_meetings_large,no_meetings])

#Data Split
title_train, title_test, meeting_train, meeting_test = train_test_split(titles['TITLE'], titles['HAD_MEETING'], test_size=0.30, random_state=101)
################################################################################################################################################


#Multinomial model
# pipeline1 = Pipeline([
#         ('bow',CountVectorizer(analyzer=text_process)),
#         ('tfidf',TfidfTransformer()),
#         ('classifier',MultinomialNB())
    
#     ])

# pipeline1.fit(title_train,meeting_train)
# predictions = pipeline1.predict(title_test)
# print(classification_report(meeting_test, predictions))
# print(confusion_matrix(meeting_test, predictions))

#Random Forest Classifier
pipeline2 = Pipeline([
        ('bow',CountVectorizer(analyzer=text_process)),
        ('tfidf',TfidfTransformer()),
        ('classifier',RandomForestClassifier())
    ])

pipeline2.fit(title_train,meeting_train)
predictions = pipeline2.predict(title_test)
print(classification_report(meeting_test, predictions))
print(confusion_matrix(meeting_test, predictions))

#################################################################################################################################################
#prospects = pd.read_csv('C:/Users/Kevin Spang/OneDrive - DataPoint Media/Documents/Python/Addaptive/open_prospects.csv')
prospects['TITLE'].apply(text_process)

targeting = pipeline2.predict(prospects['TITLE'])
targeting = pd.DataFrame(targeting)

print_out = pd.concat([prospects,targeting.set_index(prospects.index)],axis=1)
print_out = print_out.rename(columns={0:'CONTACT_FLAG'})
contact_list = print_out[print_out['CONTACT_FLAG']>0]

#add date column
today_col = dt.datetime.today().strftime('%Y-%m-%d')
today_file = dt.datetime.today().strftime('%Y_%m_%d')
contact_list['DATE'] = today_col
contact_list.sort_values('AGENCY', inplace=True)
contact_list.to_csv(f'C:/Users/Kevin Spang/OneDrive - DataPoint Media/Documents/Python/Addaptive/Title Targeting/prospect_list_{today_file}.csv',index=False)


































