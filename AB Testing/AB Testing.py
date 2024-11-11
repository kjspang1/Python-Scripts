# -*- coding: utf-8 -*-
"""
Created on Tue Jun  4 08:06:16 2024

@author: Kevin Spang
"""
#built in libraries
from datetime import datetime
import random
import math

#installed libraries
import seaborn as sns
import pandas as pd
import numpy as np
import scipy.stats as stats
import statsmodels.api as sm
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import MultipleLocator
from statsmodels.stats.power import TTestIndPower, tt_ind_solve_power
from statsmodels.stats.weightstats import ttest_ind
from statsmodels.stats.proportion import confint_proportions_2indep, proportions_chisquare

#disable warnings
from warnings import filterwarnings
filterwarnings('ignore')


#seed seed for np random
SEED = 123
np.random.seed(SEED)

#import data 
pretest = pd.read_csv('C://Users//Kevin Spang//OneDrive - DataPoint Media//Documents//Python//AB Testing//pretest.csv')
test = pd.read_csv('C://Users//Kevin Spang//OneDrive - DataPoint Media//Documents//Python//AB Testing//test.csv')

#change date fields to date time
pretest['date'] = pd.to_datetime(pretest['date'])
test['date'] = pd.to_datetime(test['date'])

###################################################################################################################################################
#table summary
pretest.head() 

#how many rows in the table
pretest.shape[0]

#date range of date field in table
print('Date Range:', pretest.date.min(), '-', pretest.date.max())

#null rate per column
pretest.isnull().mean()

#How many visitors were there?
#how many sign-ups were there?
#what was the sign-up rate

print('Total visitor count:', pretest.visitor_id.nunique())
print('Sign-up count:', pretest.submitted.sum())
print('Sign-up rate:', pretest.submitted.mean())

###################################################################################################################################################
#plot visitors per day
colors = sns.color_palette()
c1, c2 = colors[0],colors[1]

#count sign-ups by date
visits_per_day = pretest.groupby('date')['submitted'].count()
visits_mean = visits_per_day.mean()

#plot data
f, ax = plt.subplots(figsize=(12,5))
plt.plot(visits_per_day.index, visits_per_day, '-o', color=c1, linewidth=1, label='Visits')
plt.axhline(visits_mean, color=c1, linestyle='-', linewidth=3, alpha=0.3, label='Visits (mean)')

#format plot
ax.xaxis.set_major_locator(mdates.DayLocator(interval=7))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
ax.xaxis.set_minor_locator(mdates.DayLocator())
plt.title('Urban Wear Visitor Count', fontsize=10, weight='bold')
plt.ylabel('Visitors',fontsize=10)
plt.xlabel('Date',fontsize=10)
plt.legend()
plt.show()

#plot signup rate per day

#count sign-ups by date
signup_rate_per_day = pretest.groupby('date')['submitted'].mean()
signup_rate_mean = signup_rate_per_day.mean()

#plot data
f, ax = plt.subplots(figsize=(12,5))
plt.plot(signup_rate_per_day.index, signup_rate_per_day, '-o', color=c1, linewidth=1, label='Visits')
plt.axhline(signup_rate_mean, color=c1, linestyle='-', linewidth=3, alpha=0.3, label='Visits (mean)')

#format plot
ax.xaxis.set_major_locator(mdates.DayLocator(interval=7))
ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %d"))
ax.xaxis.set_minor_locator(mdates.DayLocator())
plt.title('Urban Wear Pretest Sign-Up Rate', fontsize=10, weight='bold')
plt.ylabel('Sign-up Rate',fontsize=10)
plt.xlabel('Date',fontsize=10)
plt.legend()
plt.show()

###################################################################################################################################################
#State the Hypothesis
# in this case we are testing if there is a difference in a green vs a blue sign up button
#Ho no difference in sign up rate
#Ha difference in the sign up rate

alpha = 0.05 # Set the probability threshold at 0.05. If the p-value is less than 0.05, we reject Ho
power = 0.80 # Ensure that there's 80% chance of detecting an effect with significance
mde = 0.10   # Detect a 10% improvement of the sign-up rate with statistical significance

#Proportions if the effect exists
#mean sign up rate was 10% so we are measuring distance from 10% sign up rate
p1 = 0.10      #Control group (Blue)
p2 = p1*(1+p1) #Treatment (Green)

###################################################################################################################################################
#Design the Experiment

#calculate sample size

#calculate the effect size using Cohen's D
cohen_D = sm.stats.proportion_effectsize(p1, p2)

#Estimate the sample size required per group
n = tt_ind_solve_power(effect_size=cohen_D, power=power, alpha=alpha)
n = int(round(n,-3)) #round to nearest thousand

print(f'To detect an effect of {100*(p2/p1-1):.1f}% lift from the pretest sign-up at {100*p1:.0f}%, ')
print(f'the sample size per group required is {n}.')
print(f'\nThe total sample required in the experiment is {2*n}.')
      
# Explore accross sample sizes
ttest_power = TTestIndPower()
ttest_power.plot_power(dep_var='nobs', nobs=np.arange(1000,30000,1000), effect_size=[cohen_D], title = 'Power Analysis')

# set plot parameters
plt.axhline(0.8, linestyle='--', label='Desired Power', alpha=0.5)
plt.axvline(n, linestyle='--', color='orange', label='Sample Size', alpha=0.5)
plt.ylabel('Statistical Power')
plt.grid(alpha=0.08)
plt.legend()
plt.show()

# Experiment duration
# what's the duration required to achieve the required sample size given the percentage
# of unique visitors allocated to the experiment?

alloc = np.arange(0.10,1.1,0.10)
size = round(visits_mean, -3) * alloc
days = np.ceil(2*n / size)

#generate plot
f, ax = plt.subplots(figsize=(6,4))
ax.plot(alloc, days, '-o')
ax.xaxis.set_major_locator(MultipleLocator(0.1))
ax.set_title('Days Required Given Traffic Allocation per Day')
ax.set_ylabel('Experiment Duration in Days')
ax.set_xlabel('% Traffic Allocated to the Experiment per Day')
plt.show()

# what's the duration required to achieve the required sample size given the number of
# unique visitors allocated to the experiment?

f, ax = plt.subplots(figsize=(6,4))
ax.plot(size, days, '-o')
ax.xaxis.set_major_locator(MultipleLocator(1000))
ax.set_title('Days Required Given Traffic Allocation per Day')
ax.set_ylabel('Experiment Duration in Days')
ax.set_xlabel('Traffic Allocated to the Experiment per Day')
plt.show()

# Display the number of users required per day in an experiment given the experiment duration.
print(f'For a 21-day experiment, {np.ceil(n*2/21)} users are required per day') # Too long to wait
print(f'For a 14-day experiment, {np.ceil(n*2/14)} users are required per day') # Sweet spot between risk and time
print(f'For a 7-day experiment, {np.ceil(n*2/7)} users are required per day') # Too Risky

###################################################################################################################################################
#Run the Experiment

# Get the subset tables of control and treatment results 
AB_test = test[test.experiment == 'email_test']
control_signups = AB_test[AB_test.group == 0]['submitted']
treatment_signups = AB_test[AB_test.group == 1]['submitted']

#get stats
AB_control_cnt = control_signups.sum()       # Control Sign-up Count
AB_treatment_cnt = treatment_signups.sum()   # Treatment Sign-up Count
AB_control_rate = control_signups.mean()     # Control Sign-up rate
AB_treatment_rate = treatment_signups.mean() # Treatment Sign-up rate
AB_control_size = control_signups.count()    # Control Sample Size
AB_treatment_size = treatment_signups.count()# Treatment Sample Size

#Show calculation
print(f'Control Sign-Up Rate: {AB_control_rate:.4}')
print(f'Treatment Sign-Up Rate: {AB_treatment_rate:.4}') 

# Calculate the sign-up rates per date
signups_per_day = AB_test.groupby(['group','date'])['submitted'].mean()
ctrl_props = signups_per_day.loc[0]
trt_props = signups_per_day.loc[1]

# Get the day range of experiment 
exp_days = range(1, test['date'].nunique() + 1)

# display the sign-up rate per experiment day
f, ax = plt.subplots(figsize=(10,6))
ax.plot(exp_days, ctrl_props, label='Control', color='b')
ax.plot(exp_days, trt_props, label='Treatment', color='g')
ax.axhline(AB_control_rate, label='Global Control Prop', linestyle='--', color='b')
ax.axhline(AB_treatment_rate, label='Global Treatment Prop', linestyle='--', color='g')

# Format plot
ax.set_xticks(exp_days)
ax.set_title('Email Sign-up Rates across a 14-day Experiment')
ax.set_ylabel('Sign-up Rate (Proportion)')
ax.set_xlabel('Days in the Experiment')
ax.legend()
plt.show()

###################################################################################################################################################
#Validity Threats

# In this step we will check for two of the checks for validity threats, which involve the AA test and the chi-square test for sample ratio mismatch (SRM).

# Conducting checks for the experiment ensures that the AB test result is trustworthy and reduces risk of committing type 1 or 2 errors.
# We run an AA test to ensure that there is no underlying difference between
# the control and treatment to begin with. Note that in an actual experiment,
# AA test would be conducted prior to the AB test. 
# We run a chi-square test on group sizes to check for sample-ratio mismatch (SRM). This test ensures that the randomization algorithm worked
# There are other potential checks that could be performed including segmentation analysis to perform novelty checks and such. 
# But, for this exercise, we will keep it simple to just two checks."

# ***Should technically run the AA test BEFORE the AB test

######################################################## Conduct AA test

# filter on visitors in the AA test
AA_test = pretest[pretest.experiment == 'AA_test']

# Grab the control and treatment groups in the AA test
AA_control = AA_test[AA_test.group == 0]['submitted']
AA_treatment = AA_test[AA_test.group == 1]['submitted']

# Get stats
AA_control_cnt = AA_control.sum()
AA_treatment_cnt = AA_treatment.sum()
AA_control_rate = AA_control.mean() 
AA_treatment_rate = AA_treatment.mean()
AA_control_size = AA_control.count()
AA_treatment_size = AA_treatment.count()

#show calculation
print('-----------------AA Test----------------')
print(f'Control Sign-up Rate: {AA_control_rate:.3}') 
print(f'Treatment Sign-up Rate: {AA_treatment_rate:.3}')
#0.101 vs 0.098 run chi square validiity to test hypothesis


#Sign-up rates per date
AA_signups_per_day = AA_test.groupby(['group','date'])['submitted'].mean()
AA_ctrl_props = AA_signups_per_day.loc[0]
AA_trt_props = AA_signups_per_day.loc[1]

#get the day range experiment
exp_days = range(1, AA_test['date'].nunique()+1)

# display the sign-up 
f, ax = plt.subplots(figsize=(10,6))

ax.plot(exp_days, AA_ctrl_props, label='Control', color='b')
ax.plot(exp_days, AA_trt_props, label='Treatment', color='g')
ax.axhline(AA_control_rate,label='Global Control Prop', linestyle='--', color='b')
ax.axhline(AA_treatment_rate, label='Global Treatment Prop', linestyle='--', color='g')

#format plot
ax.set_xticks(exp_days)
ax.set_title('AA Test')
ax.set_ylabel('Sign-up Rate (Proportion)')
ax.set_xlabel('Days in the Experiment')
ax.legend()
plt.show()

################################################ Run a chi square test

#Execute test
AA_chistats, AA_pvalue, AA_tab = proportions_chisquare([AA_control_cnt, AA_treatment_cnt], nobs=[AA_control_size, AA_treatment_size])

# grab dates
first_date = AA_test['date'].min().date()
last_date = AA_test['date'].max().date()

#Set the Alpha
AA_ALPHA = 0.05

print(f'--------AA Test ({first_date} - {last_date})---------------\n')
print('Ho: The sign-up rates between blue and green are the same.')
print('Ha: the sign-up rates between blue and green are different.\n')
print(f'Significance level: {AA_ALPHA}')

print(f'Significance level: {AA_ALPHA}')

print(f'Chi-Square = {AA_chistats:.3f} | P-Value = {AA_pvalue:.3f}')

print('\nConclusion:')
if AA_pvalue < AA_ALPHA:
    print('Reject Ho and conclude that there is a statistical significance in the difference between the two groups. Check for instrumentation errors.')
else:
    print('Fail to reject Ho. Therefore, proceed with AB test')


######################################################## Sample Ratio Mismatch

#set test param
SRM_ALPHA = 0.05

# get the observed and expected counts in the experiment
email_test = test[test.experiment == 'email_test']
observed = email_test.groupby('group')['experiment'].count().values
expected = [email_test.shape[0]*0.5]*2

# perform chi square goodness of fit test

chi_stats, pvalue = stats.chisquare(f_obs=observed, f_exp=expected)

print('-----------A Chi-Square Test for SRM ------------\n')
print('Ho: The ratio of samples is 1:1.')
print('Ha: The ratio of samples is not 1:1.\n')
print(f'Significance level: {SRM_ALPHA}')

print(f'Significance level: {AA_ALPHA}')

print(f'Chi-Square = {chi_stats:.3f} | P-Value = {pvalue:.3f}')

print('\nConclusion:')
if pvalue < SRM_ALPHA:
    print('Reject Ho and conclude that there is a statistical significance in the ratio of samples not being 1:1. Therefore, there is SRM.')
else:
    print('Fail to reject Ho. Therefore, there is no SRM')


#########################################################################################################################################################
# Conduct Statistical Inference
# In this step we will walk through the procedure of applying statistical tests on the email sign-up AB test. 
# We will take a look at Chi-Squared and T-Test to evaluate the results from the experiment. Though, in real life, 
# only one of the tests is sufficient, for learning, it's useful to compare and contrast the result from both.
# We will end this step by looking at the confidence interval.

#Set Alpha level for the testing
AB_ALPHA = 0.05

################################### Chi Square Test
AB_chistats, AB_pvalue, AB_tab = proportions_chisquare([AB_control_cnt, AB_treatment_cnt], nobs=[AB_control_size, AB_treatment_size])
    
# Grab dates
first_date = AB_test['date'].min().date()
last_date = AB_test['date'].max().date()

# Run results
print(f'-------- AB Test Email Sign-Ups ({first_date} - {last_date})---------\\n')
print('Ho: The sign-up rates between blue and green are the same.')
print('Ha: The sign-up rates between blue and green are different.\\n')
print(f'Significance level: {AB_ALPHA}')

print(f'Chi-Square = {AB_chistats:.3f} | P-value = {AB_pvalue:.3f}')

print('\nConclusion:')
if AB_pvalue < AB_ALPHA:
  print('Reject Ho and conclude that there is statistical significance in the difference of sign-up rates between blue and green buttons.')
else:
  print('Fail to reject Ho.') 

################################# T-Test for Proportions
AB_tstat, AB_pvalue, AB_df = ttest_ind(treatment_signups, control_signups)
        
# Grab dates
first_date = AB_test['date'].min().date()
last_date = AB_test['date'].max().date()

# Print results
print(f'-------- AB Test Email Sign-Ups ({first_date} - {last_date})---------\\n')
print('Ho: The sign-up rates between blue and green are the same.')
print('Ha: The sign-up rates between blue and green are different.\\n')
print(f'Significance level: {AB_ALPHA}')

print(f'T-Statistic = {AB_tstat:.3f} | P-value = {AB_pvalue:.3f}')

print('\\nConclusion:')
if AB_pvalue < AB_ALPHA:
  print('Reject Ho and conclude that there is statistical significance in the difference of sign-up rates between blue and green buttons.')
else:
  print('Fail to reject Ho.')


#In Both cases we reject Ho and conclude there is a difference in sign-up rates
################################# Confidence Interval

#Compute the Confidence Interval of the Test
ci = confint_proportions_2indep(AB_treatment_cnt, AB_treatment_size, AB_control_cnt, AB_control_size, method=None, compare='diff', alpha=0.05, correction=True)
lower = ci[0]
upper = ci[1]
lower_lift = ci[0] / AB_control_rate
upper_lift = ci[1] / AB_control_rate

#Print results
print('--------- Sample Sizes ----------')
print(f'Control: {AB_control_size}')
print(f'Treatment: {AB_treatment_size}')

print('\n--------- Sign-Up Counts (Rates) ----------')
print(f'Control: {AB_control_cnt} ({AB_control_rate*100:.1f}%)')
print(f'Treatment: {AB_treatment_cnt} ({AB_treatment_rate*100:.1f}%)')

print('\n--------- Differences ----------')
print(f'Absolute: {AB_treatment_rate - AB_control_rate:.4f}')
print(f'Relative (lift): {(AB_treatment_rate - AB_control_rate) / AB_control_rate*100:.1f}%')

print('\n--------- T-Stats ----------')
print(f'Test Statistic: {AB_tstat:3f}')
print(f'P-Value: {AB_pvalue:.5f}')

print('\n--------- Confidence Intervals ----------')
print(f'Absolute Difference CI: ({lower:.3f}, {upper:.3f})')
print(f'Relative Difference (lift) CI: ({lower_lift*100:.1f}%, {upper_lift*100:.1f}%)')


# Because we see that the lift was 12.5% against our usual 10% and we have tested for significance we colclude that the new user
# sign up button gets new users

#########################################################################################################################################################
# Decide whether to Launch

# In the email sign-up test for the Urban Wear pre-launch page, we aimed to improve the sign-up rate by 
# changing the submit button color from blue to green.
        
# We ran a two-week randomized controlled experiment (02/01/2022 - 02/14/2022) that enrolled a sample of users into the control (blue) 
# and treatment (green) groups.

# From the test, we observed an improvement of 12.8% lift from the benchmark (blue) at 9.6%. 
# The result was statistically significant with a 95% confidence interval between 5.7% and 19.9%. 

# Given that we observed practical and statistical significance, our recommendation is to launch the new submit button in green.




































