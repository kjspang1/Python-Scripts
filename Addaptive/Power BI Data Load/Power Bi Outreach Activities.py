# -*- coding: utf-8 -*-
"""
Created on Wed Oct 16 10:10:53 2024

@author: Kevin Spang
"""

from Login import snowflakeLogin
import pandas as pd
from snowflake.sqlalchemy import URL
from sqlalchemy import create_engine
from sqlalchemy.dialects import registry
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import serialization

sfLog = snowflakeLogin()
account = sfLog[0]
user = sfLog[1]
password = sfLog[2]
role = sfLog[3]
warehouse = sfLog[4]
database = sfLog[5]
schema = sfLog[6]
passphrase = sfLog[7]
private_key = sfLog[8]

p_key= serialization.load_pem_private_key(
    private_key,
    password=passphrase,
    backend=default_backend()
    )

pkb = p_key.private_bytes(
    encoding=serialization.Encoding.DER,
    format=serialization.PrivateFormat.PKCS8,
    encryption_algorithm=serialization.NoEncryption())

registry.register('snowflake', 'snowflake.sqlalchemy', 'dialect')

engine = create_engine(URL(
    account = account,
    user = user,
    password = password,
    database = database,
    schema = schema,
    warehouse = warehouse,
    role = role,
    ),
    connect_args={
        'private_key':pkb,
        }
    )

activities_query = '''
Select email.date,
so.platform_id salesperson_id,
ifnull(pro.prospect_added_count, 0) prospect_added_count,
ifnull(pro.unique_accounts, 0) unique_accounts,
ifnull(email.total_emails, 0) total_emails,
ifnull(email.manual_sequence_emails, 0) manual_sequence_emails,
ifnull(email.auto_sequence_emails, 0) auto_sequence_emails,
ifnull(email.other_emails, 0) other_emails,
ifnull(call.total_calls, 0) total_calls,
ifnull(call.sequence_calls, 0) sequence_calls,
ifnull(call.non_sequence_calls, 0) non_sequence_calls
from
(
    Select to_varchar(mail.delivered_at, 'yyyy-MM') date,
    concat(user.first_name,' ',user.last_name) sales_rep,
    user.id user_id,
    count(delivered_at) total_emails,
    sum(case when mail.sequence_id is not null and step.step_type = 'manual_email' then 1 else 0 end) manual_sequence_emails,
    sum(case when mail.sequence_id is not null and step.step_type = 'auto_email' then 1 else 0 end) auto_sequence_emails,
    sum(case when mail.sequence_id is null then 1 else 0 end) other_emails
    from outreach.public.mailings mail
    join outreach.public.users user
    on user.id = 
        (case 
            when mail.mailbox_id = 43 then 6 
            when mail.mailbox_id = 58 then 57 
            else mail.mailbox_id 
        end)
    left join outreach.public.sequence_steps step
    on mail.sequence_step_id = step.id
    where mail.prospect_id is not null
    and mail.delivered_at is not null
    and user.id in (select outreach_id from test_db.kevin_spang.sales_organization)
    group by to_varchar(mail.delivered_at, 'yyyy-MM'), sales_rep, user_id
) email
left join test_db.kevin_spang.sales_organization so
on email.user_id = so.outreach_id
left join
(
    Select  to_varchar(completed_at, 'yyyy-MM') date,
    concat(user.first_name,' ',user.last_name) sales_rep,
    count(call.user_id) total_calls,
    sum(case when call.sequence_id is not null then 1 else 0 end) sequence_calls,
    sum(case when call.sequence_id is null then 1 else 0 end) non_sequence_calls
    from outreach.public.calls call
    join outreach.public.users user
    on user.id = call.user_id
    where --user.id in (select outreach_id from test_db.kevin_spang.sales_organization)and 
    completed_at is not null
    group by date, sales_rep 
) call
on call.date = email.date
and call.sales_rep = email.sales_rep
left join 
(
    Select  to_varchar(state.activated_at, 'yyyy-MM') date,
    concat(user.first_name,' ',user.last_name) sales_rep,
    count(distinct state.prospect_id) prospect_added_count,
    count(distinct acc.id) unique_accounts
    from outreach.public.sequence_states state
    join outreach.public.users user
    on user.id  = state.user_id
    left join outreach.public.prospects pro
    on pro.id = state.prospect_id
    left join outreach.public.accounts acc
    on acc.id = pro.account_id
    where user.id in (select outreach_id from test_db.kevin_spang.sales_organization) 
    and state.activated_at is not null
    group by date, sales_rep
) pro
on pro.date = email.date
and pro.sales_rep = email.sales_rep
order by 1 DESC, 2 ASC, 3,4
'''

meetings_query = '''
Select to_varchar(meet.end_time, 'yyyy-MM') date, 
so.platform_id salesperson_id,
count(att.meeting_id) total_meetings, 
sum(case when att.organizer = TRUE then 1 else 0 end) meeting_organizer,
--Intro Meetings
sum(case when lower(meet.title) like '%intro%' then 1 else 0 end) intro_meetings,
sum(case when lower(meet.title) like '%intro%' and seq.meeting_type is not null then 1 else 0 end) intro_meetings_sequences,
sum(case when lower(meet.title) like '%intro%' and seq.meeting_type is null then 1 else 0 end) intro_meetings_non_sequences,
--Meetings from sequences
sum(case when seq.meeting_type is not null then 1 else 0 end) meetings_from_sequences,
sum(case when seq.meeting_type is not null and call.outreach_type is null then 1 else 0 end) meetings_from_emails_in_sequences,
sum(case when seq.meeting_type is not null and call.outreach_type is not null then 1 else 0 end) meetings_from_calls_in_sequences,
--Meetings not from sequences
sum(case when seq.meeting_type is null then 1 else 0 end) meetings_from_non_sequences,
sum(case when seq.meeting_type is null and call.outreach_type is null then 1 else 0 end) meetings_from_emails_not_in_sequences,
sum(case when seq.meeting_type is null and call.outreach_type is not null then 1 else 0 end) meetings_from_calls_not_in_sequences
from outreach.public.meeting_attendees att
left join outreach.public.users user
on user.id = att.user_id
left join outreach.public.meetings meet
on meet.id = att.meeting_id
left join (Select distinct
            att.meeting_id, 
            'sequence meeting' meeting_type
            from outreach.public.meeting_attendees att
            left join outreach.public.meetings meet
            on meet.id = att.meeting_id
            left join outreach.public.sequence_states state
            on state.prospect_id = att.prospect_id
            where att.prospect_id is not null
            and meet.canceled = 'False'
            and state.meeting_booked_at is not null 
        ) seq
on seq.meeting_id = att.meeting_id
left join (Select distinct att.meeting_id, 'call' outreach_type
        from outreach.public.meeting_attendees att
        left join outreach.public.meetings meet
        on meet.id = att.meeting_id
        where att.prospect_id is not null
        and meet.canceled = 'False'
        and att.prospect_id in (Select prospect_id 
                                from outreach.public.calls
                                where sequence_id is null
                                and answered_at is not null)
          ) call
on call.meeting_id = att.meeting_id
join test_db.kevin_spang.sales_organization so
on user.id = so.outreach_id
where user.id in (select outreach_id from test_db.kevin_spang.sales_organization) and 
att.meeting_id in (Select meeting_id from outreach.public.meeting_attendees where prospect_id is not null)
and meet.end_time <= sysdate() 
and meet.canceled = 'False'
group by date, salesperson_id
order by 3 DESC, 4 ASC
'''

try:
    con = engine.connect()
    activities = pd.read_sql(activities_query,con)
    meetings = pd.read_sql(meetings_query,con)
    con.close()
    engine.dispose()
except:
    con.close()
    engine.dispose()
    