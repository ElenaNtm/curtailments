# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 16:11:20 2024

@author: Eleni
"""

"""
Libraries
"""
import pandas as pd
from datetime import datetime, timedelta
import re
import os
import numpy as np
"""
Data-RES DA forecast
"""
base_directory = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\2. DA RES forecasts ipto"
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 7, 25)
date_format = "%Y%m%d"
current_date = start_date
 
da = pd.DataFrame()
while current_date <= end_date:
    date_str = current_date.strftime(date_format)
    print(date_str)
    file_name = f"{date_str} ISP2 IPTO RES FORECAST.xlsx"
    file_path = os.path.join(base_directory, file_name)
    file_path = re.sub(r'\\', '/', file_path)
    if os.path.exists(file_path):
        try:
            # Read the XLSX file into a pandas DataFrame
            df1 = pd.read_excel(file_path, skiprows=2)
            df1.drop(['Unnamed: 0','Unnamed: 1','Unnamed: 36'], axis = 1, inplace = True)
            df1.rename(index = {0:current_date},inplace=True)
            da = pd.concat([da, df1], axis = 0)
        except Exception as e:
                print(f"Error reading file: {file_name}, {e}")
                

    current_date += timedelta(days=1)


da.sort_index(inplace=True)
path = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA RES forecasts ipto.xlsx"
da.to_excel(path, index=True)


"""
Data-RES Intraday forecast
"""
base_directory = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\3. Intraday RES forecasts ipto"
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 7, 25)
date_format = "%Y%m%d"
current_date = start_date
 
da = pd.DataFrame()
while current_date <= end_date:
    date_str = current_date.strftime(date_format)
    print(date_str)
    file_name = f"{date_str} ISP3 IPTO RES FORECAST.xlsx"
    file_path = os.path.join(base_directory, file_name)
    file_path = re.sub(r'\\', '/', file_path)
    if os.path.exists(file_path):
        try:
            # Read the XLSX file into a pandas DataFrame
            df1 = pd.read_excel(file_path, skiprows=2)
            df1.drop(['Unnamed: 0','Unnamed: 1','Unnamed: 36'], axis = 1, inplace = True)
            df1.rename(index = {0:current_date},inplace=True)
            da = pd.concat([da, df1], axis = 0)
        except Exception as e:
                print(f"Error reading file: {file_name}, {e}")
                

    current_date += timedelta(days=1)


da.sort_index(inplace=True)
path = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\Intraday RES forecasts ipto.xlsx"
da.to_excel(path, index=True)


"""
Data-Load DA forecast
"""
base_directory = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\1. load forecast ipto\DA"
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 7, 25)
date_format = "%Y%m%d"
current_date = start_date
 
da = pd.DataFrame()
while current_date <= end_date:
    date_str = current_date.strftime(date_format)
    print(date_str)
    file_name = f"{date_str} ISP2 IPTO LOAD FORECAST"
    file_path = os.path.join(base_directory, file_name)
    file_path = re.sub(r'\\', '/', file_path)
    if os.path.exists(file_path):
        try:
            # Read the XLSX file into a pandas DataFrame
            df1 = pd.read_excel(file_path, skiprows=2)
            df1.drop(['Unnamed: 0','Unnamed: 1','Unnamed: 36'], axis = 1, inplace = True)
            df1.rename(index = {0:current_date},inplace=True)
            da = pd.concat([da, df1], axis = 0)
        except Exception as e:
                print(f"Error reading file: {file_name}, {e}")
                

    current_date += timedelta(days=1)


da.sort_index(inplace=True)
path = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Load forecasts ipto.xlsx"
da.to_excel(path, index=True)


"""
Data-Load DA forecast
"""
base_directory = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\1. load forecast ipto\DA"
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 7, 25)
date_format = "%Y%m%d"
current_date = start_date
 
da = pd.DataFrame()
while current_date <= end_date:
    date_str = current_date.strftime(date_format)
    print(date_str)
    file_name = f"{date_str} ISP2 IPTO LOAD FORECAST.xlsx"
    file_path = os.path.join(base_directory, file_name)
    file_path = re.sub(r'\\', '/', file_path)
    if os.path.exists(file_path):
        try:
            # Read the XLSX file into a pandas DataFrame
            df1 = pd.read_excel(file_path, skiprows=2)
            df1.drop(['Unnamed: 0','Unnamed: 1','Unnamed: 36'], axis = 1, inplace = True)
            df1.rename(index = {0:current_date},inplace=True)
            #da.append(df1)
            da = pd.concat([da, df1], axis = 0)
        except Exception as e:
                print(f"Error reading file: {file_name}, {e}")
                

    current_date += timedelta(days=1)

df1.head()
da.sort_index(inplace=True)
path = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Load forecasts ipto.xlsx"
da.to_excel(path, index=True)



"""
Data-Load Intrday forecast
"""
base_directory = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\1. load forecast ipto\Intraday"
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 7, 25)
date_format = "%Y%m%d"
current_date = start_date
 
da = pd.DataFrame()
while current_date <= end_date:
    date_str = current_date.strftime(date_format)
    print(date_str)
    file_name = f"{date_str} ISP3 IPTO LOAD FORECAST.xlsx"
    file_path = os.path.join(base_directory, file_name)
    file_path = re.sub(r'\\', '/', file_path)
    if os.path.exists(file_path):
        try:
            # Read the XLSX file into a pandas DataFrame
            df1 = pd.read_excel(file_path, skiprows=2)
            df1.drop(['Unnamed: 0','Unnamed: 1','Unnamed: 36'], axis = 1, inplace = True)
            df1.rename(index = {0:current_date},inplace=True)
            da = pd.concat([da, df1], axis = 0)
        except Exception as e:
                print(f"Error reading file: {file_name}, {e}")
                

    current_date += timedelta(days=1)


da.sort_index(inplace=True)
path = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\Intraday Load forecasts ipto.xlsx"
da.to_excel(path, index=True)

"""
Data- Real Production Consumption
"""
base_directory = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\4. production-consumption ipto"
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 7, 25)
date_format = "%Y%m%d"
current_date = start_date
 
da = pd.DataFrame()
while current_date <= end_date:
    date_str = current_date.strftime(date_format)
    print(date_str)
    file_name = f"Unit Production and System Facts {date_str}.xls"
    #file_name = f"{date_str} InterconnectedSystemReport.xls"
    file_path = os.path.join(base_directory, file_name)
    file_path = re.sub(r'\\', '/', file_path)
    #print(file_path)
    if os.path.exists(file_path):
        try:
            # Read the XLSX file into a pandas DataFrame
            df1 = pd.read_excel(file_path, sheet_name="System_Production", converters={'(ΣΤΟΙΧΕΙΑ SCADA ΜΗ-ΠΙΣΤΟΠΟΙΗΜΕΝΑ)':str})
        
                       
            #Net Load
            substring = 'NET LOAD'
            filtered = df1[df1.apply(lambda row: row.astype(str).str.contains(substring, case=False).any(), axis=1)]
            wanted_index = filtered.index.tolist()
            if not wanted_index:
                substring = 'ΚΑΘΑΡΟ ΦΟΡΤΙΟ'
                filtered = df1[df1.apply(lambda row: row.astype(str).str.contains(substring, case=False).any(), axis=1)]
                wanted_index = filtered.index.tolist()
            values = df1.iloc[wanted_index[len(wanted_index) - 1], 2:26]  # Extract columns 2 to 25
            row = pd.DataFrame([values.values], columns=[i for i in range(0, 24)], index=[current_date])
            da = pd.concat([da, row])
            da = da.rename(index={wanted_index[len(wanted_index) - 1]: current_date})
            
           
        except Exception as e:
            print(f"Error reading file: {file_name}, {e}")

    current_date += timedelta(days=1)


da.sort_index(inplace=True)
path = r"C:\users\eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\Load ipto.xlsx"
da.to_excel(path, index=True)




"""
Curtailments based on the DA forecasts
"""
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Load forecasts ipto.xlsx"
daload = pd.read_excel(path)
#load.info()
daload.rename(columns = { 'Unnamed: 0': 'Date'}, inplace = True)
daload.set_index('Date', inplace = True)

path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA RES forecasts ipto.xlsx"
dares = pd.read_excel(path)

dares.rename(columns = { 'Unnamed: 0': 'Date'}, inplace = True)
dares.set_index('Date', inplace = True)
#dares.info()


curtailments = pd.DataFrame()

curtailments = pd.DataFrame(np.where(dares - daload <0, 0, dares-daload), index=dares.index)

#curtailments.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA curtailments.xlsx"
curtailments.to_excel(path, index = True)

path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\Intraday RES forecasts ipto.xlsx"
inres = pd.read_excel(path)

inres.rename(columns = { 'Unnamed: 0': 'Date'}, inplace = True)
inres.set_index('Date', inplace = True)

curtailments = pd.DataFrame()

curtailments = pd.DataFrame(np.where(inres - daload <0, 0, inres-daload), index=inres.index)


path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA-Intradat curtailments.xlsx"
curtailments.to_excel(path, index = True)

path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\Intraday Load forecasts ipto.xlsx"
inload = pd.read_excel(path)
#load.info()
inload.rename(columns = { 'Unnamed: 0': 'Date'}, inplace = True)
inload.set_index('Date', inplace = True)
curtailments = pd.DataFrame()

curtailments = pd.DataFrame(np.where(inres - inload <0, 0, inres-inload), index=inres.index)

path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\Intradat curtailments.xlsx"
curtailments.to_excel(path, index = True)

##################################################################################################################################
"""
Import DAM data
"""
base_directory = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\GEARS TASKS\Market Prices Data\EnexGroup&Historical\2024 DAM"
start_date = datetime(2024, 1, 1)
end_date = datetime(2024, 7, 29)
date_format = "%Y%m%d"
current_date = start_date

da = pd.DataFrame()
while current_date <= end_date:
    date_str = current_date.strftime(date_format)
    print(date_str)
    file_name = f"{date_str}_EL-DAM_PrelimResults_EN_v01.xls"
    #file_name = f"{date_str} InterconnectedSystemReport.xls"
    file_path = os.path.join(base_directory, file_name)
    file_path = re.sub(r'\\', '/', file_path)
    #print(file_path)
    if os.path.exists(file_path):
        df1 = pd.read_excel(file_path, sheet_name="EL-DAM_Results")
        df1.drop(['TARGET', 'BIDDING_ZONE_DESCR', 'DDAY', 'ASSET_DESCR', 'SORT', 'DELIVERY_DURATION', 'PUB_TIME', 'VER'], axis = 1, inplace= True)
        da = pd.concat([da,df1], axis = 0)


    current_date += timedelta(days=1)
    
#Construct the dataframe with MCP
da.info()
da['DELIVERY_MTU'] = pd.to_datetime(da['DELIVERY_MTU'])
mcp = da.drop_duplicates(subset=['DELIVERY_MTU'], keep='first')
da.set_index('DELIVERY_MTU', inplace = True)
mcp.info()
#mcp = pd.DataFrame()
mcp.drop(['SIDE_DESCR','CLASSIFICATION','TOTAL_TRADES'], axis = 1, inplace = True)
mcp.set_index('DELIVERY_MTU', inplace = True)


mcp['date'] = mcp.index.date
mcp['hour'] = mcp.index.hour

# Pivot the DataFrame so that rows are dates and columns are hours
mcp_df = mcp.pivot(index='date', columns='hour')
mcp_df.head()

path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA MCP.xlsx"
mcp_df.to_excel(path, index = True)

#Construct dataframes with traded volumes based on the CLASSIFICATION
classif = da.groupby(['CLASSIFICATION', 'DELIVERY_MTU'])['TOTAL_TRADES'].sum()
print(classif)


path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA TRADED VOLUMES BASED ON CLASSIFICATION.xlsx"
grouped = da.groupby('CLASSIFICATION')

with pd.ExcelWriter(path) as writer:  
    for name, group in grouped:
        group.to_excel(writer, sheet_name=name, index = False)
############################################################################################################################################
"""
Keeping only the Sell
"""
core_path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA TRADED VOLUMES BASED ON CLASSIFICATION.xlsx"
        
#Big Hydro        
big_hydro = pd.read_excel(core_path, sheet_name = 'Big Hydro') 
#big_hydro.head()
mask = 'Sell'
big_hydro = big_hydro[big_hydro['SIDE_DESCR'] == mask]
group = big_hydro.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA big hydro vols.xlsx"
vols_df.to_excel(path,index = True)

#CRETE CONVENTIONAL        
cr_conv = pd.read_excel(core_path, sheet_name = 'CRETE CONVENTIONAL') 
mask = 'Sell'
cr_conv = cr_conv[cr_conv['SIDE_DESCR'] == mask]
group = cr_conv.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Crete conventional vols.xlsx"
vols_df.to_excel(path,index = True)


#CRETE LOAD       
cr_load = pd.read_excel(core_path, sheet_name = 'CRETE LOAD') 
mask = 'Sell'
cr_load = cr_load[cr_load['SIDE_DESCR'] == mask]
group = cr_load.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Crete Load vols.xlsx"
vols_df.to_excel(path,index = True)


#CRETE RENEWABLES     
cr_ren = pd.read_excel(core_path, sheet_name = 'CRETE RENEWABLES') 
mask = 'Sell'
cr_ren = cr_ren[cr_ren['SIDE_DESCR'] == mask]
group = cr_ren.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Crete renewables vols.xlsx"
vols_df.to_excel(path,index = True)


#Exoorts
exp = pd.read_excel(core_path, sheet_name = 'Exports') 
group = exp.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA exports vols.xlsx"
vols_df.to_excel(path,index = True)


#HV  
hv = pd.read_excel(core_path, sheet_name = 'HV') 
mask = 'Sell'
hv = hv[hv['SIDE_DESCR'] == mask]
group = hv.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA HV vols.xlsx"
vols_df.to_excel(path,index = True)

#Imports    
imp = pd.read_excel(core_path, sheet_name = 'Imports') 
group = imp.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Imports vols.xlsx"
vols_df.to_excel(path,index = True)


#LOSSES    
los = pd.read_excel(core_path, sheet_name = 'LOSSES') 
group = los.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA losses vols.xlsx"
vols_df.to_excel(path,index = True)


#LV
lv = pd.read_excel(core_path, sheet_name = 'LV') 
mask = 'Sell'
lv = lv[lv['SIDE_DESCR'] == mask]
group = lv.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA LV vols.xlsx"
vols_df.to_excel(path,index = True)


#Lignite
lig = pd.read_excel(core_path, sheet_name = 'Lignite') 
mask = 'Sell'
lig = lig[lig['SIDE_DESCR'] == mask]
group = lig.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA Lignite vols.xlsx"
vols_df.to_excel(path,index = True)

#MV
mv = pd.read_excel(core_path, sheet_name = 'MV') 
mask = 'Sell'
mv = mv[mv['SIDE_DESCR'] == mask]
group = mv.groupby('DELIVERY_MTU')['TOTAL_TRADES'].sum()
#vols.info()
vols = group.to_frame()
vols['date'] = vols.index.date
vols['hour'] = vols.index.hour
# Pivot the DataFrame so that rows are dates and columns are hours
vols_df = vols.pivot(index='date', columns='hour')
vols_df.head()
path = r"C:\Users\Eleni\OneDrive - Hellenic Association for Energy Economics (1)\25. Reports\5. Greek Energy Market Report\Greek Energy Market Report 2024\Electricity\electricity production\DA MV vols.xlsx"
vols_df.to_excel(path,index = True)