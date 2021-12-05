# -*- coding: utf-8 -*-
"""
Created on Mon Nov  1 18:42:50 2021

@author: rehat
"""
import pandas as pd
import numpy as np

d = pd.read_excel('C:\\Users\\rehat\\opl\\project1\\time_allocation_40.xlsx', sheet_name=None)

dfa = []
dfb = pd.DataFrame()

x=0
for x in range(200):
     dfa.append(d[str(x+1)])
     
x=0
for x in range(len(dfa)):
     dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep= 'first')  
     dfa[x] = dfa[x].rename(columns={'Ratio':'Ratio_'+str(x+1)})
     dfb = dfb.append(dfa[x])
     
dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 1)
dfb = dfb.reset_index(drop=True)    
dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
dfb = dfb.replace(0, np.nan)
dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)

with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation_40.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
    dfb.to_excel(writer, header=dfb.columns, index=False, sheet_name= 'Results Matrix')
    
d = pd.read_excel('C:\\Users\\rehat\\opl\\project1\\time_allocation_60.xlsx', sheet_name=None)

dfa = []
dfb = pd.DataFrame()

x=0
for x in range(200):
     dfa.append(d[str(x+1)])
     
x=0
for x in range(len(dfa)):
     dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep= 'first')  
     dfa[x] = dfa[x].rename(columns={'Ratio':'Ratio_'+str(x+1)})
     dfb = dfb.append(dfa[x])
     
dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 1)
dfb = dfb.reset_index(drop=True)    
dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
dfb = dfb.replace(0, np.nan)
dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)

with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation_60.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
    dfb.to_excel(writer, header=dfb.columns, index=False, sheet_name= 'Results Matrix')

d = pd.read_excel('C:\\Users\\rehat\\opl\\project1\\time_allocation_80.xlsx', sheet_name=None)

dfa = []
dfb = pd.DataFrame()

x=0
for x in range(200):
     dfa.append(d[str(x+1)])
     
x=0
for x in range(len(dfa)):
     dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep= 'first')  
     dfa[x] = dfa[x].rename(columns={'Ratio':'Ratio_'+str(x+1)})
     dfb = dfb.append(dfa[x])
     
dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 1)
dfb = dfb.reset_index(drop=True)    
dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
dfb = dfb.replace(0, np.nan)
dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)

with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation_80.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
    dfb.to_excel(writer, header=dfb.columns, index=False, sheet_name= 'Results Matrix')
    
    
d = pd.read_excel('C:\\Users\\rehat\\opl\\project1\\money_allocation_40.xlsx', sheet_name=None)

dfa = []
dfb = pd.DataFrame()

x=0
for x in range(200):
     dfa.append(d[str(x+1)])
     
x=0
for x in range(len(dfa)):
     dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep= 'first')  
     dfa[x] = dfa[x].rename(columns={'Ratio':'Ratio_'+str(x+1)})
     dfb = dfb.append(dfa[x])
     
dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 1)
dfb = dfb.reset_index(drop=True)    
dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
dfb = dfb.replace(0, np.nan)
dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)

with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation_40.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
    dfb.to_excel(writer, header=dfb.columns, index=False, sheet_name= 'Results Matrix')
    
d = pd.read_excel('C:\\Users\\rehat\\opl\\project1\\money_allocation_60.xlsx', sheet_name=None)

dfa = []
dfb = pd.DataFrame()

x=0
for x in range(200):
     dfa.append(d[str(x+1)])
     
x=0
for x in range(len(dfa)):
     dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep= 'first')  
     dfa[x] = dfa[x].rename(columns={'Ratio':'Ratio_'+str(x+1)})
     dfb = dfb.append(dfa[x])
     
dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 1)
dfb = dfb.reset_index(drop=True)    
dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
dfb = dfb.replace(0, np.nan)
dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)

with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation_60.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
    dfb.to_excel(writer, header=dfb.columns, index=False, sheet_name= 'Results Matrix')

d = pd.read_excel('C:\\Users\\rehat\\opl\\project1\\money_allocation_80.xlsx', sheet_name=None)

dfa = []
dfb = pd.DataFrame()

x=0
for x in range(200):
     dfa.append(d[str(x+1)])
     
x=0
for x in range(len(dfa)):
     dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep= 'first')  
     dfa[x] = dfa[x].rename(columns={'Ratio':'Ratio_'+str(x+1)})
     dfb = dfb.append(dfa[x])
     
dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 1)
dfb = dfb.reset_index(drop=True)    
dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
dfb = dfb.replace(0, np.nan)
dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)

with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation_80.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
    dfb.to_excel(writer, header=dfb.columns, index=False, sheet_name= 'Results Matrix')    