# -*- coding: utf-8 -*-
"""
Created on Mon May 20 22:07:33 2019

@author: Elvis Ma
"""

import pandas as pd
import numpy as np
import os
import datetime as dt

### Get the first sales date for all sales items (for collection level)
os.chdir("C:/Users/Elvis Ma/Desktop/Weekly Work/ABC analysis/First Sales Date")
domestic_first=pd.read_excel("FIRST_SALES_DATA_COMPONENT.xlsx",sheet_name="Final").loc[:,['ZINUS SKU','Final First Sales Date']]

### Get the collection data for at company level
os.chdir("C:/Users/Elvis Ma/Desktop/Weekly Work/ABC analysis/Collection Definition")
company_collection=pd.read_excel("collection_companySalesItem.xlsx", sheet_name="Company SalesItem")

### Get the 180-day sales data for all sales Item
source="C:/Users/Elvis Ma/Desktop/Weekly Work/ABC analysis/SAP sales history"
os.chdir(source)
SH_combine=pd.DataFrame()
file_list=os.listdir(source)
for i in range(len(file_list)):
    filename=file_list[i]
    book=pd.ExcelFile(filename)
    if filename!="SH_combine.xlsx":
        for j in range(len(book.sheet_names)):
            df=pd.read_excel(book,sheet_name=book.sheet_names[j])
            SH_combine=SH_combine.append(df)
        
SH_combine.rename(columns={"Net VAlue":"Sales Amount","QTY.":"Sales QTY"}, inplace=True)



## all accounts
SH_all=SH_combine.loc[:,["ZINUS SKU","Sales Amount"]]
SH_all=SH_combine.groupby("ZINUS SKU", as_index=False).agg({"Sales Amount":np.sum, "Sales QTY":np.sum}).reset_index(drop=True)

SH_all.to_excel("SH_combine.xlsx",sheet_name='all', index=False)


## Costco
SH_Costco=SH_combine[SH_combine['Customer name']=='COSTCO.COM']
SH_Costco=SH_Costco.loc[:,["ZINUS SKU","Sales QTY","Sales Amount"]]
SH_Costco=SH_Costco.groupby("ZINUS SKU", as_index=False).agg({"Sales Amount":np.sum, "Sales QTY": np.sum}).reset_index(drop=True)
## Target
SH_Target=SH_combine[SH_combine['Customer name']=='Target.com']
SH_Target=SH_Target.loc[:,["ZINUS SKU","Sales QTY","Sales Amount"]]
SH_Target=SH_Target.groupby("ZINUS SKU", as_index=False).agg({"Sales Amount":np.sum,"Sales QTY":np.sum}).reset_index(drop=True)

SH_Costco.to_excel("SH_Costco.xlsx", sheet_name='Costco', index=False)
SH_Target.to_excel("SH_Target.xlsx", sheet_name='Target', index=False)
