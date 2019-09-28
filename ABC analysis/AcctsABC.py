# -*- coding: utf-8 -*-
"""
Created on Tue Apr 16 09:50:51 2019

@author: Elvis Ma
"""

import pandas as pd
import numpy as np
import os
import datetime as dt

#LAST week column 7/28/2019
sams_col=83+10
chewy_col=32+10
costco_col=83+10
HD_col=80+10
macys_col=32+10
overstock_col=82+10
target_col=81+10
wayfair_col=81+10
wmt_col=81+10
zinus_col=54+10
eBay_col=32+10

# prepare for getting the first sales dates
os.chdir("C:/Users/Elvis Ma/Desktop/Weekly Work/ABC analysis/First Sales Date")
firstDate_tracker=pd.read_excel("FIRST_SALES_DATA_ACCTS.xlsx",sheet_name="Final")

# get the first sales date for all accounts
sams_first=firstDate_tracker.loc[:,['Collection','item','Samsclub']]
costco_first=firstDate_tracker.loc[:,['Collection','item','COSTCO.COM']]
HD_first=firstDate_tracker.loc[:,['Collection','item','Home Depot']]
macys_first=firstDate_tracker.loc[:,['Collection','item','Macys']]
overstock_first=firstDate_tracker.loc[:,['Collection','item','Overstock']]
target_first=firstDate_tracker.loc[:,['Collection','item','Target.com']]
wayfair_first=firstDate_tracker.loc[:,['Collection','item','Wayfair.com']]
wmt_first=firstDate_tracker.loc[:,['Collection','item','Walmart.COM']]
zinus_first=firstDate_tracker.loc[:,['Collection','item','Zinus.com']]   
eBay_first=firstDate_tracker.loc[:,['Collection','item','EBAY INC']]  



# get the SISR data from all accounts
os.chdir("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")

sams_sisr = pd.read_excel("samsclub_testing.xlsm",sheet_name="Master", header=2)
chewy_sisr = pd.read_excel("CHEWY_testing.xlsm",sheet_name="Demand", header=2)
costco_sisr = pd.read_excel("costcocom_testing.xlsm",sheet_name="Master", header=2)
HD_sisr = pd.read_excel("HomeDepot_testing.xlsm",sheet_name="Demand", header=2)
macys_sisr = pd.read_excel("MACYS_testing.xlsm",sheet_name="Demand", header=2)
overstock_sisr = pd.read_excel("OVERSTOCK_testing.xlsm",sheet_name="Demand", header=2)
target_sisr = pd.read_excel("target_testing.xlsm",sheet_name="Demand", header=2)
wayfair_sisr = pd.read_excel("USWayfair_testing.xlsm",sheet_name="Demand", header=1)
wmt_sisr = pd.read_excel("WMT_testing.xlsm",sheet_name="Demand", header=3)
zinus_sisr = pd.read_excel("ZINUSCOM_testing.xlsm",sheet_name="Demand", header=1)
eBay_sisr = pd.read_excel("eBay_testing.xlsm",sheet_name="Demand", header=2)

# get actual sales 
sams_sales=sams_sisr[sams_sisr["Account"]=="Actual Sales"]
chewy_sales=chewy_sisr[chewy_sisr["Account"]=="CHEWY.COM"] # calculate all weeks, no new item ABC classification
costco_sales=costco_sisr[costco_sisr["Account"]=="Actual"]
HD_sales=HD_sisr[HD_sisr["Account"]=="Home Depot"]
macys_sales=macys_sisr[macys_sisr["Account"]=="Macys"]   # calculate all weeks, no new item ABC classification
overstock_sales=overstock_sisr[overstock_sisr["Account"]=="Overstock"]
target_sales=target_sisr[target_sisr["Account"]=="Target.com"]
wayfair_sales=wayfair_sisr[wayfair_sisr["Account"]=="Actual Sales Total"]
wmt_sales=wmt_sisr[wmt_sisr["Account"]=="Walmart.COM"]
zinus_sales=zinus_sisr[zinus_sisr["Account"]=="Actual Sales Total"]
eBay_sales=eBay_sisr[eBay_sisr["Account"]=="EBAY INC"]

# data manipulation
sams_uptodate_sales=sams_sales.iloc[:,:sams_col]
sams_uptodate_sales.rename(columns={'SystemSKU':'Zinus SKU#'},inplace=True) 
chewy_uptodate_sales=chewy_sales.iloc[:,:chewy_col]
costco_uptodate_sales=costco_sales.iloc[:,:costco_col]
HD_uptodate_sales=HD_sales.iloc[:,:HD_col]
macys_uptodate_sales=macys_sales.iloc[:,:macys_col]
overstock_uptodate_sales=overstock_sales.iloc[:,:overstock_col]
target_uptodate_sales=target_sales.iloc[:,:target_col]
wayfair_uptodate_sales=wayfair_sales.iloc[:,:wayfair_col]
wayfair_uptodate_sales.rename(columns={'ZINUSSKU':'Zinus SKU#'},inplace=True)
wmt_uptodate_sales=wmt_sales.iloc[:,:wmt_col]
zinus_uptodate_sales=zinus_sales.iloc[:,:zinus_col]
zinus_uptodate_sales.rename(columns={'ZINUS SKU':'Zinus SKU#'},inplace=True)
eBay_uptodate_sales=eBay_sales.iloc[:,:eBay_col]                                    

# half Year Sales for all channels
# SamsClub good to go
sams_halfyear=pd.concat([sams_uptodate_sales.loc[:,["Dropship Cost","Zinus SKU#","Collection"]],sams_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
sams_halfyear[sams_halfyear['Dropship Cost']=='Discontinued']=0
sams_halfyear['Dropship Cost']=sams_halfyear['Dropship Cost'].fillna(0)
sams_halfyear['1/2 Year Amount']=sams_halfyear["Dropship Cost"].astype(float)*(sams_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
sams_clean=sams_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]

# Costco good to go
costco_halfyear=pd.concat([costco_uptodate_sales.loc[:,["Unit Price","Zinus SKU#","Collection"]],costco_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
costco_halfyear['Unit Price']=costco_halfyear['Unit Price'].fillna(0)
costco_halfyear['1/2 Year Amount']=costco_halfyear["Unit Price"].astype(float)*(costco_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
costco_clean=costco_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                                    
# HD good to go
HD_halfyear=pd.concat([HD_uptodate_sales.loc[:,["Unit Price","Zinus SKU#","Collection"]],HD_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
HD_halfyear['Unit Price']=HD_halfyear['Unit Price'].fillna(0)
HD_halfyear['1/2 Year Amount']=HD_halfyear["Unit Price"].astype(float)*(HD_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
HD_clean=HD_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                              
# Overstock good to go
overstock_halfyear=pd.concat([overstock_uptodate_sales.loc[:,["Unit Price","Zinus SKU#","Collection"]],overstock_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
overstock_halfyear['Unit Price']=overstock_halfyear['Unit Price'].fillna(0)
overstock_halfyear['1/2 Year Amount']=overstock_halfyear["Unit Price"].astype(float)*(overstock_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
overstock_clean=overstock_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                                          
# Target good to go
target_halfyear=pd.concat([target_uptodate_sales.loc[:,["Unit Price","Zinus SKU#","Collection"]],target_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
target_halfyear['Unit Price']=target_halfyear['Unit Price'].fillna(0)
target_halfyear['1/2 Year Amount']=target_halfyear["Unit Price"].astype(float)*(target_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
target_clean=target_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                                
# Wayfair good to go
wayfair_halfyear=pd.concat([wayfair_uptodate_sales.loc[:,["Zinus SKU#","Collection"]],wayfair_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
wayfair_halfyear['1/2 Year Amount']=wayfair_halfyear.iloc[:,2:].sum(axis=1,numeric_only=True)
wayfair_clean=wayfair_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                                
# WMT good to go
wmt_halfyear=pd.concat([wmt_uptodate_sales.loc[:,["Dropship Cost","Zinus SKU#","Collection"]],wmt_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
wmt_halfyear['Dropship Cost']=wmt_halfyear['Dropship Cost'].fillna(0)
wmt_halfyear['1/2 Year Amount']=wmt_halfyear["Dropship Cost"].astype(float)*(wmt_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
wmt_clean=wmt_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]

# Zinus.com good to go
zinus_halfyear=pd.concat([zinus_uptodate_sales.loc[:,["Zinus SKU#","Collection"]],zinus_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
zinus_halfyear['1/2 Year Amount']=zinus_halfyear.iloc[:,2:].sum(axis=1,numeric_only=True)
zinus_clean=zinus_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                                  
# Macy's good to go
macys_halfyear=pd.concat([macys_uptodate_sales.loc[:,["Zinus SKU#","Collection"]],macys_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
macys_halfyear['1/2 Year Amount']=macys_halfyear.iloc[:,2:].sum(axis=1,numeric_only=True)
macys_clean=macys_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                                 
# Chewy.com good to go
chewy_halfyear=pd.concat([chewy_uptodate_sales.loc[:,["Zinus SKU#","Collection"]],chewy_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
chewy_halfyear['1/2 Year Amount']=chewy_halfyear.iloc[:,2:].sum(axis=1,numeric_only=True)
chewy_clean=chewy_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]
                                  
# eBay good to go
eBay_halfyear=pd.concat([eBay_uptodate_sales.loc[:,["Zinus SKU#","Collection"]],eBay_uptodate_sales.iloc[:,-26:]],axis=1).reset_index(drop=True)
eBay_halfyear['1/2 Year Amount']=eBay_halfyear.iloc[:,2:].sum(axis=1,numeric_only=True)
eBay_clean=eBay_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]

# Macy's   different (not hit 26 weeks yet account)
##macys_halfyear=macys_uptodate_sales
##macys_halfyear['Unit Price']=macys_halfyear['Unit Price'].fillna(0)
##macys_halfyear['1/2 Year Amount']=macys_halfyear["Unit Price"].astype(float)*(macys_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
##macys_clean=macys_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]

# Chewy.com   different (not hit 26 weeks yet account)
##chewy_halfyear=chewy_uptodate_sales
##chewy_halfyear['Unit Price']=chewy_halfyear['Unit Price'].fillna(0)
##chewy_halfyear['1/2 Year Amount']=chewy_halfyear["Unit Price"].astype(float)*(chewy_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
##chewy_clean=chewy_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]

# EBAY INC   different (not hit 26 weeks yet account)
##eBay_halfyear=eBay_uptodate_sales
##eBay_halfyear['Unit Price']=eBay_halfyear['Unit Price'].fillna(0)
##eBay_halfyear['1/2 Year Amount']=eBay_halfyear["Unit Price"].astype(float)*(eBay_halfyear.iloc[:,3:].sum(axis=1,numeric_only=True))
##eBay_clean=eBay_halfyear.loc[:,['Zinus SKU#','Collection','1/2 Year Amount']]

                               
                                  
# data cleaning
################################################################
# By Account --> Sam's Club
sams_merge=sams_clean.merge(sams_first,left_on='Zinus SKU#', right_on='item',how='left')
sams_merge.rename(columns={'Collection_x':'Collection','Samsclub':'firstDate'},inplace=True)
sams_merged=sams_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
sams_merged['firstDate']=sams_merged['firstDate'].fillna(dt.datetime.today())
sams_merged['today']=dt.datetime.today()
sams_merged['daysDelta']=sams_merged['today']-sams_merged['firstDate']
sams_merged['daysDelta']=(sams_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
sams_ready=sams_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
sams_ready['status']=np.where(sams_ready['daysDelta']<=180,'New', 'Regular')

sams_new=sams_ready[sams_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
sams_regular=sams_ready[sams_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
sams_new['cum_sum']=sams_new['1/2 Year Amount'].cumsum()
sams_new['cum_perc']=100*sams_new.cum_sum/sams_new['1/2 Year Amount'].sum()
sams_new['class']=np.where(sams_new['cum_perc']<=80,'New A', np.where(sams_new['cum_perc']<=95,'New B','New C'))

### Regular Item
sams_regular['cum_sum']=sams_regular['1/2 Year Amount'].cumsum()
sams_regular['cum_perc']=100*sams_regular.cum_sum/sams_regular['1/2 Year Amount'].sum()
sams_regular['class']=np.where(sams_regular['cum_perc']<=80,'A', np.where(sams_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
sams_ABC = pd.concat([sams_regular,sams_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Costco
costco_merge=costco_clean.merge(costco_first,left_on='Zinus SKU#', right_on='item',how='left')
costco_merge.rename(columns={'Collection_x':'Collection','COSTCO.COM':'firstDate'},inplace=True)
costco_merged=costco_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
costco_merged['firstDate']=costco_merged['firstDate'].fillna(dt.datetime.today())
costco_merged['today']=dt.datetime.today()
costco_merged['daysDelta']=costco_merged['today']-costco_merged['firstDate']
costco_merged['daysDelta']=(costco_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
costco_ready=costco_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
costco_ready['status']=np.where(costco_ready['daysDelta']<=180,'New', 'Regular')

costco_new=costco_ready[costco_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
costco_regular=costco_ready[costco_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
costco_new['cum_sum']=costco_new['1/2 Year Amount'].cumsum()
costco_new['cum_perc']=100*costco_new.cum_sum/costco_new['1/2 Year Amount'].sum()
costco_new['class']=np.where(costco_new['cum_perc']<=80,'New A', np.where(costco_new['cum_perc']<=95,'New B','New C'))

### Regular Item
costco_regular['cum_sum']=costco_regular['1/2 Year Amount'].cumsum()
costco_regular['cum_perc']=100*costco_regular.cum_sum/costco_regular['1/2 Year Amount'].sum()
costco_regular['class']=np.where(costco_regular['cum_perc']<=80,'A', np.where(costco_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
costco_ABC = pd.concat([costco_regular,costco_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Home Depot
HD_merge=HD_clean.merge(HD_first,left_on='Zinus SKU#', right_on='item',how='left')
HD_merge.rename(columns={'Collection_x':'Collection','Home Depot':'firstDate'},inplace=True)
HD_merged=HD_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
HD_merged['firstDate']=HD_merged['firstDate'].fillna(dt.datetime.today())
HD_merged['today']=dt.datetime.today()
HD_merged['daysDelta']=HD_merged['today']-HD_merged['firstDate']
HD_merged['daysDelta']=(HD_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
HD_ready=HD_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
HD_ready['status']=np.where(HD_ready['daysDelta']<=180,'New', 'Regular')

HD_new=HD_ready[HD_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
HD_regular=HD_ready[HD_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
HD_new['cum_sum']=HD_new['1/2 Year Amount'].cumsum()
HD_new['cum_perc']=100*HD_new.cum_sum/HD_new['1/2 Year Amount'].sum()
HD_new['class']=np.where(HD_new['cum_perc']<=80,'New A', np.where(HD_new['cum_perc']<=95,'New B','New C'))

### Regular Item
HD_regular['cum_sum']=HD_regular['1/2 Year Amount'].cumsum()
HD_regular['cum_perc']=100*HD_regular.cum_sum/HD_regular['1/2 Year Amount'].sum()
HD_regular['class']=np.where(HD_regular['cum_perc']<=80,'A', np.where(HD_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
HD_ABC = pd.concat([HD_regular,HD_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Macys
macys_merge=macys_clean.merge(macys_first,left_on='Zinus SKU#', right_on='item',how='left')
macys_merge.rename(columns={'Collection_x':'Collection','Macys':'firstDate'},inplace=True)
macys_merged=macys_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
macys_merged['firstDate']=macys_merged['firstDate'].fillna(dt.datetime.today())
macys_merged['today']=dt.datetime.today()
macys_merged['daysDelta']=macys_merged['today']-macys_merged['firstDate']
macys_merged['daysDelta']=(macys_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
macys_ready=macys_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
macys_ready['status']=np.where(macys_ready['daysDelta']<=180,'New', 'Regular')

macys_new=macys_ready[macys_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
macys_regular=macys_ready[macys_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
macys_new['cum_sum']=macys_new['1/2 Year Amount'].cumsum()
macys_new['cum_perc']=100*macys_new.cum_sum/macys_new['1/2 Year Amount'].sum()
macys_new['class']=np.where(macys_new['cum_perc']<=80,'New A', np.where(macys_new['cum_perc']<=95,'New B','New C'))

### Regular Item
macys_regular['cum_sum']=macys_regular['1/2 Year Amount'].cumsum()
macys_regular['cum_perc']=100*macys_regular.cum_sum/macys_regular['1/2 Year Amount'].sum()
macys_regular['class']=np.where(macys_regular['cum_perc']<=80,'A', np.where(macys_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
macys_ABC = pd.concat([macys_regular,macys_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Overstock
overstock_merge=overstock_clean.merge(overstock_first,left_on='Zinus SKU#', right_on='item',how='left')
overstock_merge.rename(columns={'Collection_x':'Collection','Overstock':'firstDate'},inplace=True)
overstock_merged=overstock_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
overstock_merged['firstDate']=overstock_merged['firstDate'].fillna(dt.datetime.today())
overstock_merged['today']=dt.datetime.today()
overstock_merged['daysDelta']=overstock_merged['today']-overstock_merged['firstDate']
overstock_merged['daysDelta']=(overstock_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
overstock_ready=overstock_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
overstock_ready['status']=np.where(overstock_ready['daysDelta']<=180,'New', 'Regular')

overstock_new=overstock_ready[overstock_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
overstock_regular=overstock_ready[overstock_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
overstock_new['cum_sum']=overstock_new['1/2 Year Amount'].cumsum()
overstock_new['cum_perc']=100*overstock_new.cum_sum/overstock_new['1/2 Year Amount'].sum()
overstock_new['class']=np.where(overstock_new['cum_perc']<=80,'New A', np.where(overstock_new['cum_perc']<=95,'New B','New C'))

### Regular Item
overstock_regular['cum_sum']=overstock_regular['1/2 Year Amount'].cumsum()
overstock_regular['cum_perc']=100*overstock_regular.cum_sum/overstock_regular['1/2 Year Amount'].sum()
overstock_regular['class']=np.where(overstock_regular['cum_perc']<=80,'A', np.where(overstock_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
overstock_ABC = pd.concat([overstock_regular,overstock_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Target
target_merge=target_clean.merge(target_first,left_on='Zinus SKU#', right_on='item',how='left')
target_merge.rename(columns={'Collection_x':'Collection','Target.com':'firstDate'},inplace=True)
target_merged=target_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
target_merged['firstDate']=target_merged['firstDate'].fillna(dt.datetime.today())
target_merged['today']=dt.datetime.today()
target_merged['daysDelta']=target_merged['today']-target_merged['firstDate']
target_merged['daysDelta']=(target_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
target_ready=target_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
target_ready['status']=np.where(target_ready['daysDelta']<=180,'New', 'Regular')

target_new=target_ready[target_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
target_regular=target_ready[target_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
target_new['cum_sum']=target_new['1/2 Year Amount'].cumsum()
target_new['cum_perc']=100*target_new.cum_sum/target_new['1/2 Year Amount'].sum()
target_new['class']=np.where(target_new['cum_perc']<=80,'New A', np.where(target_new['cum_perc']<=95,'New B','New C'))

### Regular Item
target_regular['cum_sum']=target_regular['1/2 Year Amount'].cumsum()
target_regular['cum_perc']=100*target_regular.cum_sum/target_regular['1/2 Year Amount'].sum()
target_regular['class']=np.where(target_regular['cum_perc']<=80,'A', np.where(target_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
target_ABC = pd.concat([target_regular,target_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Wayfair
wayfair_merge=wayfair_clean.merge(wayfair_first,left_on='Zinus SKU#', right_on='item',how='left')
wayfair_merge.rename(columns={'Collection_x':'Collection','Wayfair.com':'firstDate'},inplace=True)
wayfair_merged=wayfair_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
wayfair_merged['firstDate']=wayfair_merged['firstDate'].fillna(dt.datetime.today())
wayfair_merged['today']=dt.datetime.today()
wayfair_merged['daysDelta']=wayfair_merged['today']-wayfair_merged['firstDate']
wayfair_merged['daysDelta']=(wayfair_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
wayfair_ready=wayfair_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
wayfair_ready['status']=np.where(wayfair_ready['daysDelta']<=180,'New', 'Regular')

wayfair_new=wayfair_ready[wayfair_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
wayfair_regular=wayfair_ready[wayfair_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
wayfair_new['cum_sum']=wayfair_new['1/2 Year Amount'].cumsum()
wayfair_new['cum_perc']=100*wayfair_new.cum_sum/wayfair_new['1/2 Year Amount'].sum()
wayfair_new['class']=np.where(wayfair_new['cum_perc']<=80,'New A', np.where(wayfair_new['cum_perc']<=95,'New B','New C'))

### Regular Item
wayfair_regular['cum_sum']=wayfair_regular['1/2 Year Amount'].cumsum()
wayfair_regular['cum_perc']=100*wayfair_regular.cum_sum/wayfair_regular['1/2 Year Amount'].sum()
wayfair_regular['class']=np.where(wayfair_regular['cum_perc']<=80,'A', np.where(wayfair_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
wayfair_ABC = pd.concat([wayfair_regular,wayfair_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Walmart
wmt_merge=wmt_clean.merge(wmt_first,left_on='Zinus SKU#', right_on='item',how='left')
wmt_merge.rename(columns={'Collection_x':'Collection','Walmart.COM':'firstDate'},inplace=True)
wmt_merged=wmt_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
wmt_merged['firstDate']=wmt_merged['firstDate'].fillna(dt.datetime.today())
wmt_merged['today']=dt.datetime.today()
wmt_merged['daysDelta']=wmt_merged['today']-wmt_merged['firstDate']
wmt_merged['daysDelta']=(wmt_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
wmt_ready=wmt_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
wmt_ready['status']=np.where(wmt_ready['daysDelta']<=180,'New', 'Regular')

wmt_new=wmt_ready[wmt_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
wmt_regular=wmt_ready[wmt_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
wmt_new['cum_sum']=wmt_new['1/2 Year Amount'].cumsum()
wmt_new['cum_perc']=100*wmt_new.cum_sum/wmt_new['1/2 Year Amount'].sum()
wmt_new['class']=np.where(wmt_new['cum_perc']<=80,'New A', np.where(wmt_new['cum_perc']<=95,'New B','New C'))

### Regular Item
wmt_regular['cum_sum']=wmt_regular['1/2 Year Amount'].cumsum()
wmt_regular['cum_perc']=100*wmt_regular.cum_sum/wmt_regular['1/2 Year Amount'].sum()
wmt_regular['class']=np.where(wmt_regular['cum_perc']<=80,'A', np.where(wmt_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
wmt_ABC = pd.concat([wmt_regular,wmt_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> Zinus
zinus_merge=zinus_clean.merge(zinus_first,left_on='Zinus SKU#', right_on='item',how='left')
zinus_merge.rename(columns={'Collection_x':'Collection','Zinus.com':'firstDate'},inplace=True)
zinus_merged=zinus_merge.loc[:,['Zinus SKU#','Collection','1/2 Year Amount','firstDate']]
zinus_merged['firstDate']=zinus_merged['firstDate'].fillna(dt.datetime.today())
zinus_merged['today']=dt.datetime.today()
zinus_merged['daysDelta']=zinus_merged['today']-zinus_merged['firstDate']
zinus_merged['daysDelta']=(zinus_merged['daysDelta']/np.timedelta64(1,'D')).astype(int)
zinus_ready=zinus_merged.groupby('Collection',as_index=False).agg({'1/2 Year Amount':np.sum,'daysDelta':np.max})
zinus_ready['status']=np.where(zinus_ready['daysDelta']<=180,'New', 'Regular')

zinus_new=zinus_ready[zinus_ready['status']=='New'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)
zinus_regular=zinus_ready[zinus_ready['status']=='Regular'].loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### New Item
zinus_new['cum_sum']=zinus_new['1/2 Year Amount'].cumsum()
zinus_new['cum_perc']=100*zinus_new.cum_sum/zinus_new['1/2 Year Amount'].sum()
zinus_new['class']=np.where(zinus_new['cum_perc']<=80,'New A', np.where(zinus_new['cum_perc']<=95,'New B','New C'))

### Regular Item
zinus_regular['cum_sum']=zinus_regular['1/2 Year Amount'].cumsum()
zinus_regular['cum_perc']=100*zinus_regular.cum_sum/zinus_regular['1/2 Year Amount'].sum()
zinus_regular['class']=np.where(zinus_regular['cum_perc']<=80,'A', np.where(zinus_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
zinus_ABC = pd.concat([zinus_regular,zinus_new],axis=0).loc[:,['Collection','1/2 Year Amount','cum_perc','class']] 

################################################################
# By Account --> chewy (new account)
chewy_merge=chewy_clean
chewy_ready=chewy_merge
chewy_regular=chewy_ready.loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### Regular Item
chewy_regular['cum_sum']=chewy_regular['1/2 Year Amount'].cumsum()
chewy_regular['cum_perc']=100*chewy_regular.cum_sum/chewy_regular['1/2 Year Amount'].sum()
chewy_regular['class']=np.where(chewy_regular['cum_perc']<=80,'A', np.where(chewy_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
chewy_ABC = chewy_regular.loc[:,['Collection','1/2 Year Amount','cum_perc','class']]

################################################################
# By Account --> eBay (new account)
eBay_merge=eBay_clean
eBay_ready=eBay_merge
eBay_regular=eBay_ready.loc[:,['Collection','1/2 Year Amount','status']].sort_values('1/2 Year Amount', ascending=False).reset_index(drop=True)

### Regular Item
eBay_regular['cum_sum']=eBay_regular['1/2 Year Amount'].cumsum()
eBay_regular['cum_perc']=100*eBay_regular.cum_sum/eBay_regular['1/2 Year Amount'].sum()
eBay_regular['class']=np.where(eBay_regular['cum_perc']<=80,'A', np.where(eBay_regular['cum_perc']<=95,'B','C'))

### Combine New & Regular
eBay_ABC = eBay_regular.loc[:,['Collection','1/2 Year Amount','cum_perc','class']]


#################################################################
#### Write files ######

os.chdir("C:/Users/Elvis Ma/Desktop/Weekly Work")
writer=pd.ExcelWriter('ABC_all.xlsx',engine='xlsxwriter')

sams_ABC.to_excel(writer,sheet_name='Sams',index=False)
costco_ABC.to_excel(writer,sheet_name='Costco',index=False)
HD_ABC.to_excel(writer,sheet_name='HD',index=False)
macys_ABC.to_excel(writer,sheet_name='Macys',index=False)
overstock_ABC.to_excel(writer,sheet_name='Overstock',index=False)
target_ABC.to_excel(writer,sheet_name='Target',index=False)
wayfair_ABC.to_excel(writer,sheet_name='USWF',index=False)
wmt_ABC.to_excel(writer,sheet_name='WMT',index=False)
zinus_ABC.to_excel(writer,sheet_name='Zinus',index=False)
chewy_ABC.to_excel(writer,sheet_name='Chewy',index=False)
eBay_ABC.to_excel(writer,sheet_name='eBay',index=False)
writer.save()


######################## gather the collection information #####################
#collection_list=[sams_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],chewy_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],
    #                costco_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],HD_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]], 
    #               macys_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],
    #              overstock_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],target_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],
    #             wayfair_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],
    #            wmt_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]],zinus_uptodate_sales.loc[:,['Zinus SKU#',"Collection"]]]
#collection_combine=pd.concat(collection_list,axis=0).reset_index(drop=True)
#collection_combine.to_excel("C:/Users/Elvis Ma/Desktop/Daily Work/ABC report/collection_combine.xlsx")
