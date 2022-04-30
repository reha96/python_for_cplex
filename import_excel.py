# -*- coding: utf-8 -*-
"""
Created on Mon Nov  1 18:42:50 2021

@author: rehat
"""
import pandas as pd
import numpy as np

m60 = 'money_allocation_60.xlsx'
t60 = 'time_allocation_60.xlsx'
m80 = 'money_allocation_80.xlsx'
t80 = 'time_allocation_80.xlsx'

for file in [m60, t60, t80, m80]:
    d = pd.read_excel(file, sheet_name=None)
    
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
         
    dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 'columns')
    dfb = dfb.reset_index(drop=True)    
    dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
    dfb = dfb.replace(0, np.nan)
    dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
    dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)
    
    df_1 = dfb[dfb["Obs_Char"].str.contains("and")==False]
    df_2 = dfb[dfb["Obs_Char"].str.contains("and")==True]
    df_1 = df_1[df_1["Obs_Char"].str.contains("Allocation")==False]
    df_1 = df_1.sort_values("Mean")
    df_2 = df_2.sort_values("Mean")
    df_1 = df_1.reset_index(drop=True)
    df_2 = df_2.reset_index(drop=True)
    
    
    """
    Means and STDs (1-level and 2-level)
    """
    d = pd.read_excel(file, sheet_name=None)
    
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
         
    dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 'columns')
    dfb = dfb.reset_index(drop=True)    
    dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
    dfb = dfb.replace(0, np.nan)
    dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
    dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)
    
    df_1 = dfb[dfb["Obs_Char"].str.contains("and")==False]
    df_2 = dfb[dfb["Obs_Char"].str.contains("and")==True]
    df_1 = df_1[df_1["Obs_Char"].str.contains("Allocation")==False]
    df_1 = df_1.sort_values("Mean")
    df_2 = df_2.sort_values("Mean")
    df_1 = df_1.reset_index(drop=True)
    df_2 = df_2.reset_index(drop=True)
    
    with pd.ExcelWriter(file, engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
        df_1.to_excel(writer, header=df_1.columns, index=False, sheet_name= '1-level Branching')
    
    with pd.ExcelWriter(file, engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
        df_2.to_excel(writer, header=df_2.columns, index=False, sheet_name= '2-level Branching')
        
    """
    Dominance matrix
    df_1 for 1-level and df_2 for 2-level
    
    """
    x=0
    y=0
    p=0
    df_3 = pd.DataFrame()
    df_3["Obs_Char"] = df_1["Obs_Char"]
    
    for z in df_3["Obs_Char"]:
        df_3[z] = 0  
    
    i=0 
    for i in range(len(df_1)):
        for y in range(len(df_1)):
            for x in range(200):
                    if df_1['Ratio_'+str(x+1)][i] < df_1['Ratio_'+str(x+1)][y]:
                        p = p+1
            df_3.iloc[y,i+1] = (p/200)*100
            p=0
    
    df_3 = df_3.transpose()
    df_3.columns = df_3.iloc[0]
    df_3 = df_3[1:]
    with pd.ExcelWriter(file, engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
        df_3.to_excel(writer, header=df_3.columns, index=True, sheet_name= 'Strictly Better Matrix 1-level')
        
    x=0
    y=0
    p=0
    df_3 = pd.DataFrame()
    df_3["Obs_Char"] = df_2["Obs_Char"]
    
    for z in df_3["Obs_Char"]:
        df_3[z] = 0  
    
    i=0 
    for i in range(len(df_2)):
        for y in range(len(df_2)):
            for x in range(200):
                    if df_2['Ratio_'+str(x+1)][i] < df_2['Ratio_'+str(x+1)][y]:
                        p = p+1
            df_3.iloc[y,i+1] = (p/200)*100
            p=0
    
    df_3 = df_3.transpose()
    df_3.columns = df_3.iloc[0]
    df_3 = df_3[1:]
    with pd.ExcelWriter(file, engine="openpyxl", mode='a', if_sheet_exists='new') as writer:  
        df_3.to_excel(writer, header=df_3.columns, index=True, sheet_name= 'Strictly Better Matrix 2-level')
