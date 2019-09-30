# -*- coding: utf-8 -*-
"""
Created on Tue Mar 19 23:39:35 2019

@author: Elvis Ma
"""

import pandas as pd
SH_report =pd.read_excel("C:/Users/Elvis Ma/Desktop/Weekly Work/SH.xlsx")

sams_sisr = pd.read_excel("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales/samsclub_testing.xlsm",header=2)

fcst_num =pd.ExcelFile("C:/Users/Elvis Ma/Desktop/Weekly Work/Forecast Numbers/FcstNum.xlsx")

sams_xlsm = pd.ExcelFile("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales/samsclub_testing.xlsm")



import pandas as pd
sap_sku = pd.read_excel("C:/Users/Elvis Ma/Desktop/Daily Work/ABC report/Master SKU list.xlsx")
sap_sku_list = pd.DataFrame(sap_sku['MATNR'])
sap_sku_list.to_excel("C:/Users/Elvis Ma/Desktop/Daily Work/ABC report/SAP_SKUs.xlsx")

SH_list=pd.ExcelFile("C:/Users/Elvis Ma/Desktop/Daily Work/ABC report/DATA_TO_COMBINE.xlsx")
#for sheet in SH_list.sheet_names:
 #   SH_list.parse(sheet)
 
pre_sap = SH_list.parse('pre SAP')
post_sap =SH_list.parse('SAP')

sap_both=pre_sap.merge(post_sap,on='item', how='left')
sap_combine=sap_sku_list.merge(sap_both,left_on="MATNR", right_on='item', how='left')
sap_both.to_excel("C:/Users/Elvis Ma/Desktop/Daily Work/ABC report/SAP_BOTH.xlsx")
sap_combine.to_excel("C:/Users/Elvis Ma/Desktop/Daily Work/ABC report/SAP_COMBINE.xlsx")


import pandas as pd
import os
os.chdir("C:/Users/Elvis Ma/Desktop")
file_list =pd.ExcelFile("fileTexting.xlsx")

df=pd.DataFrame()
for sheetName in file_list.sheet_names:
    df=df.append(file_list.parse(sheetName), ignore_index=True)