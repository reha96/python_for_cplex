import pandas as pd
import numpy as np
import time
from subprocess import call
from openpyxl import Workbook
import logging
import psutil

# close Excel and ILOG Studio before running the code

logging.basicConfig(format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p',
                    filename="log_cplex.log", level=logging.INFO, force=True)

t = time.time()

configuration = 'Original'  # set configuration (sheet)

data_where = 'C:\\Users\\rehat\\opl\\project1\\AllExercices.xlsx'  # set data location

data = pd.read_excel(data_where, sheet_name=configuration)

tasks_which = 'time' # select 'all', 'money', 'time'


df3 = []  # to calculate ratios and store types for money allocation task
df3bis = []  # to calculate ratios and store types for time allocation task

char = "C"  # to increment where we write the results at each s
c = ord(char[0])


'''
step 0 - sampling from data

- set nb_draws=XXX for the number of draws
- set sample_size=XXX for the fraction of the original data we want in each draw

NOTE: Running the whole script from the beginning will overwrite the results file
'''
nb_draws = 1     # set the number of sampling
sample_size = 0.999    # set the sample size as a fraction of the original data. For example: 0.8, 0.6, and 0.999 for 80 pcent, 60 pcent and full sample

logging.info('Data Location: ' + data_where +
             ' Configuration: ' + configuration + ' Draws: ' + str(nb_draws) + ' Sample Size: ' + str(sample_size))


# Set s to the desired draw and run from this line to avoid overwriting results !
for s in range(nb_draws):
    looptime = time.time() - t
    logging.info('Loop ' + str(s+1) +
                 ' started, elapsed time: ' + str(looptime))
    df = data.sort_values("ID")
    df = df.sample(frac=sample_size)  # add random_state to set seed if need be
    df.to_excel(
        'C:\\Users\\rehat\\opl\\project1\\data_for_python.xlsx', index=False)
    if tasks_which == 'all' or tasks_which == 'money':
        df2 = Workbook()
        # prepare the results file for the money allocation task
        df2.save(filename='C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx')
    if tasks_which == 'all' or tasks_which ==  'time':
        df2bis = Workbook()
        # prepare the results file for the time allocation task
        df2bis.save(filename='C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx')
    '''
    1st step - no categories (finding tau_hat) for MONEY ALLOCATION TASK
    '''
    if tasks_which == 'all' or tasks_which == 'money':
        i = 1  # counter for excel cells (to properly store output in results file)
        with open("C:\\Users\\rehat\\opl\\project1\\part.dat", 'w') as f:
            f.write('NR_Goods=2;\n'+'SheetConnection comm("data_for_python.xlsx");\n')
            f.write('NR_Observations= ' + str(len(df)) + ';\n')
            f.write("""AllP from SheetRead(comm,"'Sheet1'!E2:F""" +
                    str(len(df)+1)+"""");\n""")
            f.write("""AllQ from SheetRead(comm,"'Sheet1'!C2:D""" +
                    str(len(df)+1)+"""");\n""")
            f.write("""Income from SheetRead(comm,"'Sheet1'!G2:G""" +
                    str(len(df)+1)+"""");\n""")
            f.write('SheetConnection comm2("money_allocation.xlsx");\n')
            f.write("""NR_Types to SheetWrite(comm2,"'Sheet'!"""+str(chr(c)) +
                    """"""+str(i)+""":"""+str(chr(c))+""""""+str(i)+"""");\n""")
            f.write("""NR_Observations to SheetWrite(comm2,"'Sheet'!D""" +
                    str(i)+""":D"""+str(i)+"""");\n""")
            f.write("""runtype = "Money_Allocation";\n""")
            f.write("""runtype to SheetWrite(comm2,"'Sheet'!A""" +
                    str(i)+""":A"""+str(i)+"""");\n""")
            f.write('runtype2 = " ";\n')
            f.write("""runtype2 to SheetWrite(comm2,"'Sheet'!B""" +
                    str(i)+""":B"""+str(i)+"""");\n""")
        call(["oplrun.exe", "C:\\Users\\rehat\\opl\\project1\\part.mod",
             "C:\\Users\\rehat\\opl\\project1\\part.dat"])
        i = i+1  # increment by 1 at the end of each run
        '''
        2nd step - loop over k and l (finding tau_k)
        '''
        for k in df[['Gender', 'Age', 'Education', 'Marital Status', 'Employment']]:
            df = df.sort_values(k)
            df.to_excel(
                'C:\\Users\\rehat\\opl\\project1\\data_for_python.xlsx', index=False)
            d = df.groupby([k]).size().reset_index(name='Count')
            for l in range(len(d)):  # counter for each state in an observable char
                with open("C:\\Users\\rehat\\opl\\project1\\part.dat", 'w') as f:
                    f.write('NR_Goods=2;\n' +
                            'SheetConnection comm("data_for_python.xlsx");\n')
                    f.write('NR_Observations= ' + str(d['Count'][l]) + ';\n')
                    if l == 0:
                        f.write("""AllP from SheetRead(comm,"'Sheet1'!E2:F""" +
                                str(d['Count'][l]+1)+"""");\n""")
                        f.write("""AllQ from SheetRead(comm,"'Sheet1'!C2:D""" +
                                str(d['Count'][l]+1)+"""");\n""")
                        f.write("""Income from SheetRead(comm,"'Sheet1'!G2:G""" +
                                str(d['Count'][l]+1)+"""");\n""")
                        n = d['Count'][l]
                    elif l > 0:
                        f.write("""AllP from SheetRead(comm,"'Sheet1'!E""" +
                                str(n+2)+""":F""" + str(n+d['Count'][l]+1)+"""");\n""")
                        f.write("""AllQ from SheetRead(comm,"'Sheet1'!C""" +
                                str(n+2)+""":D""" + str(n+d['Count'][l]+1)+"""");\n""")
                        f.write("""Income from SheetRead(comm,"'Sheet1'!G""" +
                                str(n+2)+""":G""" + str(n+d['Count'][l]+1)+"""");\n""")
                        n = n+d['Count'][l]
                    f.write('SheetConnection comm2("money_allocation.xlsx");\n')
                    f.write("""NR_Types to SheetWrite(comm2,"'Sheet'!"""+str(chr(c)) +
                            """"""+str(i)+""":"""+str(chr(c))+""""""+str(i)+"""");\n""")
                    f.write("""NR_Observations to SheetWrite(comm2,"'Sheet'!D""" +
                            str(i)+""":D"""+str(i)+"""");\n""")
                    f.write('runtype = "'+str(k)+'";\n')
                    f.write("""runtype to SheetWrite(comm2,"'Sheet'!A""" +
                            str(i)+""":A"""+str(i)+"""");\n""")
                    f.write('runtype2 = "'+str(d[k][l])+'";\n')
                    f.write("""runtype2 to SheetWrite(comm2,"'Sheet'!B""" +
                            str(i)+""":B"""+str(i)+"""");\n""")
                call(["oplrun.exe", "C:\\Users\\rehat\\opl\\project1\\part.mod",
                     "C:\\Users\\rehat\\opl\\project1\\part.dat"])
                i = i+1  # increment by 1 at the end of each run
            n = 0  # reset to 0 at the end of k
            l = 0  # reset to 0 at the end of k
        '''
        3rd step - loop over k and l given j (find tau_k,j) 
        '''
        for j in df[['Gender', 'Age', 'Education', 'Marital Status', 'Employment']]:
            for k in df[['Gender', 'Age', 'Education', 'Marital Status', 'Employment']]:
                # there should be an elegant way to do this but let's do it manually since it's just 5 choose 2
                if j == 'Age' and k == 'Gender':
                    continue
                if j == 'Education' and (k == 'Gender' or k == 'Age'):
                    continue
                if j == 'Marital Status' and (k == 'Gender' or k == 'Age' or k == 'Education'):
                    continue
                if j == 'Employment' and (k == 'Gender' or k == 'Age' or k == 'Education' or k == 'Marital Status'):
                    continue
                if j == k:
                    continue
                df = df.sort_values([j, k])
                df.to_excel(
                    'C:\\Users\\rehat\\opl\\project1\\data_for_python.xlsx', index=False)
                d = df.groupby([j, k]).size().reset_index(name='Count')
                for l in range(len(d)):  # counter for each state in an observable char
                    with open("C:\\Users\\rehat\\opl\\project1\\part.dat", 'w') as f:
                        f.write('NR_Goods=2;\n' +
                                'SheetConnection comm("data_for_python.xlsx");\n')
                        f.write('NR_Observations= ' + str(d['Count'][l]) + ';\n')
                        if l == 0:
                            f.write(
                                """AllP from SheetRead(comm,"'Sheet1'!E2:F""" + str(d['Count'][l]+1)+"""");\n""")
                            f.write(
                                """AllQ from SheetRead(comm,"'Sheet1'!C2:D""" + str(d['Count'][l]+1)+"""");\n""")
                            f.write(
                                """Income from SheetRead(comm,"'Sheet1'!G2:G""" + str(d['Count'][l]+1)+"""");\n""")
                            n = d['Count'][l]
                        elif l > 0:
                            f.write("""AllP from SheetRead(comm,"'Sheet1'!E""" +
                                    str(n+2)+""":F""" + str(n+d['Count'][l]+1)+"""");\n""")
                            f.write("""AllQ from SheetRead(comm,"'Sheet1'!C""" +
                                    str(n+2)+""":D""" + str(n+d['Count'][l]+1)+"""");\n""")
                            f.write("""Income from SheetRead(comm,"'Sheet1'!G""" +
                                    str(n+2)+""":G""" + str(n+d['Count'][l]+1)+"""");\n""")
                            n = n+d['Count'][l]
                        f.write('SheetConnection comm2("money_allocation.xlsx");\n')
                        f.write("""NR_Types to SheetWrite(comm2,"'Sheet'!"""+str(chr(c)) +
                                """"""+str(i)+""":"""+str(chr(c))+""""""+str(i)+"""");\n""")
                        f.write("""NR_Observations to SheetWrite(comm2,"'Sheet'!D""" +
                                str(i)+""":D"""+str(i)+"""");\n""")
                        f.write('runtype2 = "' +
                                str(d[j][l])+' and ' + str(d[k][l]) + '";\n')
                        f.write("""runtype2 to SheetWrite(comm2,"'Sheet'!B""" +
                                str(i)+""":B"""+str(i)+"""");\n""")
                        f.write('runtype = "'+str(j)+' and ' + str(k)+'";\n')
                        f.write("""runtype to SheetWrite(comm2,"'Sheet'!A""" +
                                str(i)+""":A"""+str(i)+"""");\n""")
                    call(["oplrun.exe", "C:\\Users\\rehat\\opl\\project1\\part.mod",
                         "C:\\Users\\rehat\\opl\\project1\\part.dat"])
                    i = i+1  # increment by 1 at the end of each run
                n = 0  # reset to 0 at the end of k
                l = 0  # reset to 0 at the end of k
        df3.append(pd.read_excel(
            'C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx', header=None))
        logging.info('Money Allocation loop ended successfully. CPU usage: ' +
                     str(psutil.cpu_percent(1)) + ' RAM usage: ' + str(psutil.virtual_memory()[2]))

    '''
    1bis - no categories (finding tau_hat) for TIME ALLOCATION TASK
    '''
    if tasks_which == 'all' or tasks_which == 'time':
        i = 1  # counter for excel cells (to properly store output in results file)
        with open("C:\\Users\\rehat\\opl\\project1\\part.dat", 'w') as f:
            f.write('NR_Goods=2;\n'+'SheetConnection comm("data_for_python.xlsx");\n')
            f.write('NR_Observations= ' + str(len(df)) + ';\n')
            f.write("""AllP from SheetRead(comm,"'Sheet1'!K2:L""" +
                    str(len(df)+1)+"""");\n""")
            f.write("""AllQ from SheetRead(comm,"'Sheet1'!I2:J""" +
                    str(len(df)+1)+"""");\n""")
            f.write("""Income from SheetRead(comm,"'Sheet1'!M2:M""" +
                    str(len(df)+1)+"""");\n""")
            f.write('SheetConnection comm2("time_allocation.xlsx");\n')
            f.write("""NR_Types to SheetWrite(comm2,"'Sheet'!"""+str(chr(c)) +
                    """"""+str(i)+""":"""+str(chr(c))+""""""+str(i)+"""");\n""")
            f.write("""NR_Observations to SheetWrite(comm2,"'Sheet'!D""" +
                    str(i)+""":D"""+str(i)+"""");\n""")
            f.write("""runtype = "Time_Allocation";\n""")
            f.write("""runtype to SheetWrite(comm2,"'Sheet'!A""" +
                    str(i)+""":A"""+str(i)+"""");\n""")
            f.write('runtype2 = " ";\n')
            f.write("""runtype2 to SheetWrite(comm2,"'Sheet'!B""" +
                    str(i)+""":B"""+str(i)+"""");\n""")
        call(["oplrun.exe", "C:\\Users\\rehat\\opl\\project1\\part.mod",
             "C:\\Users\\rehat\\opl\\project1\\part.dat"])
        i = i+1  # increment by 1 at the end of each run
        '''
        2bis - loop over k and l (finding tau_k)
        '''
        for k in df[['Gender', 'Age', 'Education', 'Marital Status', 'Employment']]:
            df = df.sort_values(k)
            df.to_excel(
                'C:\\Users\\rehat\\opl\\project1\\data_for_python.xlsx', index=False)
            d = df.groupby([k]).size().reset_index(name='Count')
            for l in range(len(d)):  # counter for each state in an observable char
                with open("C:\\Users\\rehat\\opl\\project1\\part.dat", 'w') as f:
                    f.write('NR_Goods=2;\n' +
                            'SheetConnection comm("data_for_python.xlsx");\n')
                    f.write('NR_Observations= ' + str(d['Count'][l]) + ';\n')
                    if l == 0:
                        f.write("""AllP from SheetRead(comm,"'Sheet1'!K2:L""" +
                                str(d['Count'][l]+1)+"""");\n""")
                        f.write("""AllQ from SheetRead(comm,"'Sheet1'!I2:J""" +
                                str(d['Count'][l]+1)+"""");\n""")
                        f.write("""Income from SheetRead(comm,"'Sheet1'!M2:M""" +
                                str(d['Count'][l]+1)+"""");\n""")
                        n = d['Count'][l]
                    elif l > 0:
                        f.write("""AllP from SheetRead(comm,"'Sheet1'!K""" +
                                str(n+2)+""":L""" + str(n+d['Count'][l]+1)+"""");\n""")
                        f.write("""AllQ from SheetRead(comm,"'Sheet1'!I""" +
                                str(n+2)+""":J""" + str(n+d['Count'][l]+1)+"""");\n""")
                        f.write("""Income from SheetRead(comm,"'Sheet1'!M""" +
                                str(n+2)+""":M""" + str(n+d['Count'][l]+1)+"""");\n""")
                        n = n+d['Count'][l]
                    f.write('SheetConnection comm2("time_allocation.xlsx");\n')
                    f.write("""NR_Types to SheetWrite(comm2,"'Sheet'!"""+str(chr(c)) +
                            """"""+str(i)+""":"""+str(chr(c))+""""""+str(i)+"""");\n""")
                    f.write("""NR_Observations to SheetWrite(comm2,"'Sheet'!D""" +
                            str(i)+""":D"""+str(i)+"""");\n""")
                    f.write('runtype = "'+str(k)+'";\n')
                    f.write("""runtype to SheetWrite(comm2,"'Sheet'!A""" +
                            str(i)+""":A"""+str(i)+"""");\n""")
                    f.write('runtype2 = "'+str(d[k][l])+'";\n')
                    f.write("""runtype2 to SheetWrite(comm2,"'Sheet'!B""" +
                            str(i)+""":B"""+str(i)+"""");\n""")
                call(["oplrun.exe", "C:\\Users\\rehat\\opl\\project1\\part.mod",
                     "C:\\Users\\rehat\\opl\\project1\\part.dat"])
                i = i+1  # increment by 1 at the end of each run
            n = 0  # reset to 0 at the end of k
            l = 0  # reset to 0 at the end of k
        '''
        3bis - loop over k and l given j (find tau_k,j) 
        '''
        for j in df[['Gender', 'Age', 'Education', 'Marital Status', 'Employment']]:
            for k in df[['Gender', 'Age', 'Education', 'Marital Status', 'Employment']]:
                # there should be an elegant way to do this but let's do it manually since it's just 5 choose 2
                if j == 'Age' and k == 'Gender':
                    continue
                if j == 'Education' and (k == 'Gender' or k == 'Age'):
                    continue
                if j == 'Marital Status' and (k == 'Gender' or k == 'Age' or k == 'Education'):
                    continue
                if j == 'Employment' and (k == 'Gender' or k == 'Age' or k == 'Education' or k == 'Marital Status'):
                    continue
                if j == k:
                    continue
                df = df.sort_values([j, k])
                df.to_excel(
                    'C:\\Users\\rehat\\opl\\project1\\data_for_python.xlsx', index=False)
                d = df.groupby([j, k]).size().reset_index(name='Count')
                for l in range(len(d)):  # counter for each state in an observable char
                    with open("C:\\Users\\rehat\\opl\\project1\\part.dat", 'w') as f:
                        f.write('NR_Goods=2;\n' +
                                'SheetConnection comm("data_for_python.xlsx");\n')
                        f.write('NR_Observations= ' + str(d['Count'][l]) + ';\n')
                        if l == 0:
                            f.write(
                                """AllP from SheetRead(comm,"'Sheet1'!K2:L""" + str(d['Count'][l]+1)+"""");\n""")
                            f.write(
                                """AllQ from SheetRead(comm,"'Sheet1'!I2:J""" + str(d['Count'][l]+1)+"""");\n""")
                            f.write(
                                """Income from SheetRead(comm,"'Sheet1'!M2:M""" + str(d['Count'][l]+1)+"""");\n""")
                            n = d['Count'][l]
                        elif l > 0:
                            f.write("""AllP from SheetRead(comm,"'Sheet1'!K""" +
                                    str(n+2)+""":L""" + str(n+d['Count'][l]+1)+"""");\n""")
                            f.write("""AllQ from SheetRead(comm,"'Sheet1'!I""" +
                                    str(n+2)+""":J""" + str(n+d['Count'][l]+1)+"""");\n""")
                            f.write("""Income from SheetRead(comm,"'Sheet1'!M""" +
                                    str(n+2)+""":M""" + str(n+d['Count'][l]+1)+"""");\n""")
                            n = n+d['Count'][l]
                        f.write('SheetConnection comm2("time_allocation.xlsx");\n')
                        f.write("""NR_Types to SheetWrite(comm2,"'Sheet'!"""+str(chr(c)) +
                                """"""+str(i)+""":"""+str(chr(c))+""""""+str(i)+"""");\n""")
                        f.write("""NR_Observations to SheetWrite(comm2,"'Sheet'!D""" +
                                str(i)+""":D"""+str(i)+"""");\n""")
                        f.write('runtype2 = "' +
                                str(d[j][l])+' and ' + str(d[k][l]) + '";\n')
                        f.write("""runtype2 to SheetWrite(comm2,"'Sheet'!B""" +
                                str(i)+""":B"""+str(i)+"""");\n""")
                        f.write('runtype = "'+str(j)+' and ' + str(k)+'";\n')
                        f.write("""runtype to SheetWrite(comm2,"'Sheet'!A""" +
                                str(i)+""":A"""+str(i)+"""");\n""")
                    call(["oplrun.exe", "C:\\Users\\rehat\\opl\\project1\\part.mod",
                         "C:\\Users\\rehat\\opl\\project1\\part.dat"])
                    i = i+1  # increment by 1 at the end of each run
                n = 0  # reset to 0 at the end of k
                l = 0  # reset to 0 at the end of k
        df3bis.append(pd.read_excel(
            'C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx', header=None))
        logging.info('Time Allocation loop ended successfully. CPU usage: ' +
                     str(psutil.cpu_percent(1)) + 'RAM usage: ' + str(psutil.virtual_memory()[2]))


'''
4th step - calculate kappa ratios and report the results for money allocation task 
'''
if tasks_which == 'all' or tasks_which == 'money':
    resultcopy = df3[:]  # copy of results so things do not get lost
    
    df5 = pd.DataFrame()  # intermediate dataframe object combining best 2-level values
    
    x = 0
    for x in range(len(df3)):
        df3[x] = df3[x].rename(columns={3: 'n'})
        df3[x] = df3[x].rename(columns={2: 'Types'})
        df3[x] = df3[x].rename(columns={1: 'State'})
        df3[x] = df3[x].rename(columns={0: 'Obs_Char'})
    
    x = 0
    for x in range(len(df3)):
        df3[x]['Sum_Types'] = df3[x].groupby(
            ['Obs_Char'])['Types'].transform('sum')
        df3[x]['Ratio_' + str(x+1)] = df3[x]['Sum_Types']/df3[x]['Sum_Types'][0]
        df3[x] = df3[x].sort_values('Obs_Char')
        with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
            df3[x].to_excel(writer, header=["Obs_Char", "State", "Types",
                            "n", "Sum_Types", "Ratio"], index=False, sheet_name=str(x+1))
    
    d = pd.read_excel(
        'C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx', sheet_name=None)
    dfa = []
    dfb = pd.DataFrame()
    
    x = 0
    for x in range(nb_draws):
        dfa.append(d[str(x+1)])
    
    x = 0
    for x in range(len(dfa)):
        dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep='first')
        dfa[x] = dfa[x].rename(columns={'Ratio': 'Ratio_'+str(x+1)})
        dfb = dfb.append(dfa[x])
    
    dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 'columns')
    dfb = dfb.reset_index(drop=True)
    dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
    dfb = dfb.replace(0, np.nan)
    dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
    dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)
    
    df_1 = dfb[dfb["Obs_Char"].str.contains("and") == False]
    df_2 = dfb[dfb["Obs_Char"].str.contains("and") == True]
    df_1 = df_1[df_1["Obs_Char"].str.contains("Allocation") == False]
    df_1 = df_1.sort_values("Mean")
    df_2 = df_2.sort_values("Mean")
    df_1 = df_1.reset_index(drop=True)
    df_2 = df_2.reset_index(drop=True)
    
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_1.to_excel(writer, header=df_1.columns, index=False,
                      sheet_name='1-level Branching')
    
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_2.to_excel(writer, header=df_2.columns, index=False,
                      sheet_name='2-level Branching')
    
    
    """
    Stricly better ratios matrix
    """
    
    x = 0
    y = 0
    p = 0
    df_3 = pd.DataFrame()
    df_3["Obs_Char"] = df_1["Obs_Char"]
    
    for z in df_3["Obs_Char"]:
        df_3[z] = 0
    
    i = 0
    for i in range(len(df_1)):
        for y in range(len(df_1)):
            for x in range(nb_draws):
                if df_1['Ratio_'+str(x+1)][i] < df_1['Ratio_'+str(x+1)][y]:
                    p = p+1
            df_3.iloc[y, i+1] = (p/nb_draws)*100
            p = 0
    
    df_3 = df_3.transpose()
    df_3.columns = df_3.iloc[0]
    df_3 = df_3[1:]
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_3.to_excel(writer, header=df_3.columns, index=True,
                      sheet_name='Strictly Better Matrix 1-level')
    
    x = 0
    y = 0
    p = 0
    df_3 = pd.DataFrame()
    df_3["Obs_Char"] = df_2["Obs_Char"]
    
    for z in df_3["Obs_Char"]:
        df_3[z] = 0
    
    i = 0
    for i in range(len(df_2)):
        for y in range(len(df_2)):
            for x in range(nb_draws):
                if df_2['Ratio_'+str(x+1)][i] < df_2['Ratio_'+str(x+1)][y]:
                    p = p+1
            df_3.iloc[y, i+1] = (p/nb_draws)*100
            p = 0
    
    df_3 = df_3.transpose()
    df_3.columns = df_3.iloc[0]
    df_3 = df_3[1:]
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\money_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_3.to_excel(writer, header=df_3.columns, index=True,
                      sheet_name='Strictly Better Matrix 2-level')


'''
4bis - calculate kappa ratios and report the results for time allocation task 
'''
if tasks_which == 'all' or tasks_which == 'time':
    resultcopybis = df3bis[:]  # copy of results so things do not get lost
    
    df5bis = pd.DataFrame()  # intermediate dataframe object combining best 2-level values
    
    x = 0
    for x in range(len(df3bis)):
        df3bis[x] = df3bis[x].rename(columns={3: 'n'})
        df3bis[x] = df3bis[x].rename(columns={2: 'Types'})
        df3bis[x] = df3bis[x].rename(columns={1: 'State'})
        df3bis[x] = df3bis[x].rename(columns={0: 'Obs_Char'})
    
    x = 0
    for x in range(len(df3bis)):
        df3bis[x]['Sum_Types'] = df3bis[x].groupby(
            ['Obs_Char'])['Types'].transform('sum')
        df3bis[x]['Ratio_' + str(x+1)] = df3bis[x]['Sum_Types'] / \
            df3bis[x]['Sum_Types'][0]
        df3bis[x] = df3bis[x].sort_values('Obs_Char')
        with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
            df3bis[x].to_excel(writer, header=["Obs_Char", "State", "Types",
                               "n", "Sum_Types", "Ratio"], index=False, sheet_name=str(x+1))
    
    d = pd.read_excel(
        'C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx', sheet_name=None)
    
    dfa = []
    dfb = pd.DataFrame()
    
    x = 0
    for x in range(nb_draws):
        dfa.append(d[str(x+1)])
    
    x = 0
    for x in range(len(dfa)):
        dfa[x] = dfa[x].drop_duplicates(subset=['Obs_Char', 'Ratio'], keep='first')
        dfa[x] = dfa[x].rename(columns={'Ratio': 'Ratio_'+str(x+1)})
        dfb = dfb.append(dfa[x])
    
    dfb = dfb.drop(['Types', 'n', 'State', 'Sum_Types'], 'columns')
    dfb = dfb.reset_index(drop=True)
    dfb = dfb.groupby(['Obs_Char'], as_index=False).sum()
    dfb = dfb.replace(0, np.nan)
    dfb['Mean'] = dfb.mean(axis=1, numeric_only=True)
    dfb['Sample std'] = dfb.std(axis=1, numeric_only=True)
    
    df_1 = dfb[dfb["Obs_Char"].str.contains("and") == False]
    df_2 = dfb[dfb["Obs_Char"].str.contains("and") == True]
    df_1 = df_1[df_1["Obs_Char"].str.contains("Allocation") == False]
    df_1 = df_1.sort_values("Mean")
    df_2 = df_2.sort_values("Mean")
    df_1 = df_1.reset_index(drop=True)
    df_2 = df_2.reset_index(drop=True)
    
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_1.to_excel(writer, header=df_1.columns, index=False,
                      sheet_name='1-level Branching')
    
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_2.to_excel(writer, header=df_2.columns, index=False,
                      sheet_name='2-level Branching')
    
    """
    Stricly better ratios matrix
    """
    x = 0
    y = 0
    p = 0
    df_3 = pd.DataFrame()
    df_3["Obs_Char"] = df_1["Obs_Char"]
    
    for z in df_3["Obs_Char"]:
        df_3[z] = 0
    
    i = 0
    for i in range(len(df_1)):
        for y in range(len(df_1)):
            for x in range(nb_draws):
                if df_1['Ratio_'+str(x+1)][i] < df_1['Ratio_'+str(x+1)][y]:
                    p = p+1
            df_3.iloc[y, i+1] = (p/nb_draws)*100
            p = 0
    
    df_3 = df_3.transpose()
    df_3.columns = df_3.iloc[0]
    df_3 = df_3[1:]
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_3.to_excel(writer, header=df_3.columns, index=True,
                      sheet_name='Strictly Better Matrix 1-level')
    
    x = 0
    y = 0
    p = 0
    df_3 = pd.DataFrame()
    df_3["Obs_Char"] = df_2["Obs_Char"]
    
    for z in df_3["Obs_Char"]:
        df_3[z] = 0
    
    i = 0
    for i in range(len(df_2)):
        for y in range(len(df_2)):
            for x in range(nb_draws):
                if df_2['Ratio_'+str(x+1)][i] < df_2['Ratio_'+str(x+1)][y]:
                    p = p+1
            df_3.iloc[y, i+1] = (p/nb_draws)*100
            p = 0
    
    df_3 = df_3.transpose()
    df_3.columns = df_3.iloc[0]
    df_3 = df_3[1:]
    with pd.ExcelWriter('C:\\Users\\rehat\\opl\\project1\\time_allocation.xlsx', engine="openpyxl", mode='a', if_sheet_exists='new') as writer:
        df_3.to_excel(writer, header=df_3.columns, index=True,
                      sheet_name='Strictly Better Matrix 2-level')


'''
CONCLUDING INFORMATION
'''
elapsed = time.time() - t
logging.info('Computation done with ' + str(nb_draws) + ' subsamples whose size equals ' + str(sample_size*100) + ' percent of the original data. \nTotal elapsed time (in seconds): '
             + str(elapsed) + '\nAverage loop length (in seconds): ' + str(elapsed/nb_draws))
