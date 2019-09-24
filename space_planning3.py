# -*- coding: utf-8 -*-
"""
Created on Fri May  3 16:04:46 2019

@author: Elvis Ma
"""

import pandas as pd
import numpy as np
import os
import datetime as dt

os.chdir("C:/Users/Elvis Ma/Desktop/Weekly Work/space_planning")

tracy_other_dds = pd.read_excel("space_planning.xlsx",sheet_name='Tracy_Other_DDS',header=0)
chs_other_dds=pd.read_excel("space_planning.xlsx",sheet_name='CHS_Other_DDS',header=0)

tracy_amz_dds = pd.read_excel("space_planning.xlsx",sheet_name='Tracy_AMZ_DDS',header=0)
chs_amz_dds=pd.read_excel("space_planning.xlsx",sheet_name='CHS_AMZ_DDS',header=0)

tracy_amz_direct=pd.read_excel("space_planning.xlsx",sheet_name='Tracy_AMZ_DIRECT',header=0)
chs_amz_direct=pd.read_excel("space_planning.xlsx",sheet_name='CHS_AMZ_DIRECT',header=0)

### Amazon Direct : 1
tracy_amz_direct_melt=pd.melt(tracy_amz_direct,id_vars=['Zinus SKU#'],value_vars=list(tracy_amz_direct.columns.values[1:]))
tracy_amz_direct_melt.rename(columns={'Zinus SKU#':'Material','variable':'Week of','value':'Demand Forecast'},inplace=True)
tracy_amz_direct_melt['Plant']=2000
tracy_amz_direct_melt['Category']=1

chs_amz_direct_melt=pd.melt(chs_amz_direct,id_vars=['Zinus SKU#'],value_vars=list(chs_amz_direct.columns.values[1:]))
chs_amz_direct_melt.rename(columns={'Zinus SKU#':'Material','variable':'Week of','value':'Demand Forecast'},inplace=True)
chs_amz_direct_melt['Plant']=2100
chs_amz_direct_melt['Category']=1

### Amazon DDS : 2
tracy_amz_dds_melt=pd.melt(tracy_amz_dds,id_vars=['Zinus SKU#'],value_vars=list(tracy_amz_dds.columns.values[1:]))
tracy_amz_dds_melt.rename(columns={'Zinus SKU#':'Material','variable':'Week of','value':'Demand Forecast'},inplace=True)
tracy_amz_dds_melt['Plant']=2000
tracy_amz_dds_melt['Category']=2

chs_amz_dds_melt=pd.melt(chs_amz_dds,id_vars=['Zinus SKU#'],value_vars=list(chs_amz_dds.columns.values[1:]))
chs_amz_dds_melt.rename(columns={'Zinus SKU#':'Material','variable':'Week of','value':'Demand Forecast'},inplace=True)
chs_amz_dds_melt['Plant']=2100
chs_amz_dds_melt['Category']=2

### Other Channel DDS : 3
tracy_other_dds_melt=pd.melt(tracy_other_dds,id_vars=['Zinus SKU#'],value_vars=list(tracy_other_dds.columns.values[1:]))
tracy_other_dds_melt.rename(columns={'Zinus SKU#':'Material','variable':'Week of','value':'Demand Forecast'},inplace=True)
tracy_other_dds_melt['Plant']=2000
tracy_other_dds_melt['Category']=3

chs_other_dds_melt=pd.melt(chs_other_dds,id_vars=['Zinus SKU#'],value_vars=list(chs_other_dds.columns.values[1:]))
chs_other_dds_melt.rename(columns={'Zinus SKU#':'Material','variable':'Week of','value':'Demand Forecast'},inplace=True)
chs_other_dds_melt['Plant']=2100
chs_other_dds_melt['Category']=3

combine_list=[tracy_amz_dds_melt,tracy_other_dds_melt,tracy_amz_direct_melt,chs_other_dds_melt, chs_amz_direct_melt, chs_amz_dds_melt]

space_plan =pd.concat(combine_list,axis=0).reset_index(drop=True)
space_plan=space_plan.drop_duplicates()
space_plan.to_excel('space_plan_melted.xlsx',sheet_name='both location', index=False)
