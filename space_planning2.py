# -*- coding: utf-8 -*-
"""
Created on Wed May  8 23:30:03 2019

@author: Elvis Ma
"""

import pandas as pd
import os
os.chdir("C:/Users/Elvis Ma/Desktop/Weekly Work/space_planning")

startCol=50 # 9/22/2019 weekly change 

#### reading Amazon Direct Data in SISR

SB_direct=pd.read_excel("SISR_SB_AMZ.xlsm", sheet_name='SB', header=4)
SB_direct=pd.concat([SB_direct[SB_direct['Account']=='Amazon Direct'].iloc[:,6],SB_direct[SB_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
SB_direct=SB_direct.fillna(0)

BX_direct=pd.read_excel("SISR_BX_AMZ.xlsm", sheet_name='BX', header=4)
BX_direct=pd.concat([BX_direct[BX_direct['Account']=='Amazon Direct'].iloc[:,6],BX_direct[BX_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
BX_direct=BX_direct.fillna(0)

FR_direct=pd.read_excel("SISR_FR_AMZ.xlsm", sheet_name='FR', header=4)
FR_direct=pd.concat([FR_direct[FR_direct['Account']=='Amazon Direct'].iloc[:,6],FR_direct[FR_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
FR_direct=FR_direct.fillna(0)

SM_direct=pd.read_excel("SISR_SM_AMZ.xlsm", sheet_name='SM', header=4)
SM_direct=pd.concat([SM_direct[SM_direct['Account']=='Amazon Direct'].iloc[:,6],SM_direct[SM_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
SM_direct=SM_direct.fillna(0)

FM_direct=pd.read_excel("SISR_FM_AMZ.xlsm", sheet_name='FM', header=4)
FM_direct=pd.concat([FM_direct[FM_direct['Account']=='Amazon Direct'].iloc[:,6],FM_direct[FM_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
FM_direct=FM_direct.fillna(0)

PB_direct=pd.read_excel("SISR_PB_AMZ.xlsm", sheet_name='PB', header=4)
PB_direct=pd.concat([PB_direct[PB_direct['Account']=='Amazon Direct'].iloc[:,6],PB_direct[PB_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
PB_direct=PB_direct.fillna(0)

OT_direct=pd.read_excel("SISR_OT_AMZ.xlsm", sheet_name='OT', header=4)
OT_direct=pd.concat([OT_direct[OT_direct['Account']=='Amazon Direct'].iloc[:,6],OT_direct[OT_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
OT_direct=OT_direct.fillna(0)

FN_direct=pd.read_excel("SISR_FN_AMZ.xlsm", sheet_name='FN', header=4)
FN_direct=pd.concat([FN_direct[FN_direct['Account']=='Amazon Direct'].iloc[:,6],FN_direct[FN_direct['Account']=='Amazon Direct'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
FN_direct=FN_direct.fillna(0)

#### reading Amazon DDS Data in SISR

SB_amzdds=pd.read_excel("SISR_SB_AMZ.xlsm", sheet_name='SB', header=4)
SB_amzdds=pd.concat([SB_amzdds[SB_amzdds['Account']=='Amazon DDS'].iloc[:,6],SB_amzdds[SB_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
SB_amzdds=SB_amzdds.fillna(0)

BX_amzdds=pd.read_excel("SISR_BX_AMZ.xlsm", sheet_name='BX', header=4)
BX_amzdds=pd.concat([BX_amzdds[BX_amzdds['Account']=='Amazon DDS'].iloc[:,6],BX_amzdds[BX_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
BX_amzdds=BX_amzdds.fillna(0)

FR_amzdds=pd.read_excel("SISR_FR_AMZ.xlsm", sheet_name='FR', header=4)
FR_amzdds=pd.concat([FR_amzdds[FR_amzdds['Account']=='Amazon DDS'].iloc[:,6],FR_amzdds[FR_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
FR_amzdds=FR_amzdds.fillna(0)

SM_amzdds=pd.read_excel("SISR_SM_AMZ.xlsm", sheet_name='SM', header=4)
SM_amzdds=pd.concat([SM_amzdds[SM_amzdds['Account']=='Amazon DDS'].iloc[:,6],SM_amzdds[SM_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
SM_amzdds=SM_amzdds.fillna(0)

FM_amzdds=pd.read_excel("SISR_FM_AMZ.xlsm", sheet_name='FM', header=4)
FM_amzdds=pd.concat([FM_amzdds[FM_amzdds['Account']=='Amazon DDS'].iloc[:,6],FM_amzdds[FM_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
FM_amzdds=FM_amzdds.fillna(0)

PB_amzdds=pd.read_excel("SISR_PB_AMZ.xlsm", sheet_name='PB', header=4)
PB_amzdds=pd.concat([PB_amzdds[PB_amzdds['Account']=='Amazon DDS'].iloc[:,6],PB_amzdds[PB_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
PB_amzdds=PB_amzdds.fillna(0)

OT_amzdds=pd.read_excel("SISR_OT_AMZ.xlsm", sheet_name='OT', header=4)
OT_amzdds=pd.concat([OT_amzdds[OT_amzdds['Account']=='Amazon DDS'].iloc[:,6],OT_amzdds[OT_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
OT_amzdds=OT_amzdds.fillna(0)

FN_amzdds=pd.read_excel("SISR_FN_AMZ.xlsm", sheet_name='FN', header=4)
FN_amzdds=pd.concat([FN_amzdds[FN_amzdds['Account']=='Amazon DDS'].iloc[:,6],FN_amzdds[FN_amzdds['Account']=='Amazon DDS'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True)
FN_amzdds=FN_amzdds.fillna(0)

### function to open and save system generated files from space_planning1
import xlwings as xl

def df_from_excel(path):
    app = xl.App(visible=False)
    book = app.books.open(path)
    book.save()
    app.kill()
    return pd.read_excel(path,sheet_name=path[5:7],header=4)

SB_dds = df_from_excel("SISR_SB_DDS.xlsx")
BX_dds = df_from_excel("SISR_BX_DDS.xlsx")
FR_dds = df_from_excel("SISR_FR_DDS.xlsx")
SM_dds = df_from_excel("SISR_SM_DDS.xlsx")
FM_dds = df_from_excel("SISR_FM_DDS.xlsx")
PB_dds = df_from_excel("SISR_PB_DDS.xlsx")
OT_dds = df_from_excel("SISR_OT_DDS.xlsx")
FN_dds = df_from_excel("SISR_FN_DDS.xlsx")

#### reading West east DF dds data in SISR

SB_west=pd.concat([SB_dds[SB_dds['Account']=='West DF'].iloc[:,6],SB_dds[SB_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
SB_east=pd.concat([SB_dds[SB_dds['Account']=='East DF'].iloc[:,6],SB_dds[SB_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)


BX_west=pd.concat([BX_dds[BX_dds['Account']=='West DF'].iloc[:,6],BX_dds[BX_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
BX_east=pd.concat([BX_dds[BX_dds['Account']=='East DF'].iloc[:,6],BX_dds[BX_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)



FR_west=pd.concat([FR_dds[FR_dds['Account']=='West DF'].iloc[:,6],FR_dds[FR_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
FR_east=pd.concat([FR_dds[FR_dds['Account']=='East DF'].iloc[:,6],FR_dds[FR_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)



SM_west=pd.concat([SM_dds[SM_dds['Account']=='West DF'].iloc[:,6],SM_dds[SM_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
SM_east=pd.concat([SM_dds[SM_dds['Account']=='East DF'].iloc[:,6],SM_dds[SM_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)


FM_west=pd.concat([FM_dds[FM_dds['Account']=='West DF'].iloc[:,6],FM_dds[FM_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
FM_east=pd.concat([FM_dds[FM_dds['Account']=='East DF'].iloc[:,6],FM_dds[FM_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)



PB_west=pd.concat([PB_dds[PB_dds['Account']=='West DF'].iloc[:,6],PB_dds[PB_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
PB_east=pd.concat([PB_dds[PB_dds['Account']=='East DF'].iloc[:,6],PB_dds[PB_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)



OT_west=pd.concat([OT_dds[OT_dds['Account']=='West DF'].iloc[:,6],OT_dds[OT_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
OT_east=pd.concat([OT_dds[OT_dds['Account']=='East DF'].iloc[:,6],OT_dds[OT_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)



FN_west=pd.concat([FN_dds[FN_dds['Account']=='West DF'].iloc[:,6],FN_dds[FN_dds['Account']=='West DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)
FN_east=pd.concat([FN_dds[FN_dds['Account']=='East DF'].iloc[:,6],FN_dds[FN_dds['Account']=='East DF'].iloc[:,startCol:startCol+24]],axis=1).reset_index(drop=True).fillna(0)


####### Mattress SKU list
mattress_sku=pd.concat([FM_direct['Zinus SKU#'],SM_direct['Zinus SKU#']],axis=0).reset_index(drop=True)
mattress_df=pd.DataFrame(mattress_sku)
mattress_df['Check']='mattress'                                 
##stacking data
direct_dflist=[SB_direct,BX_direct,FR_direct,SM_direct,FM_direct,PB_direct,FN_direct,OT_direct]
amzdds_dflist=[SB_amzdds,BX_amzdds,FR_amzdds,SM_amzdds,FM_amzdds,PB_amzdds,FN_amzdds,OT_amzdds]
west_dflist=[SB_west,BX_west,FR_west,SM_west,FM_west,PB_west,FN_west,OT_west]
east_dflist=[SB_east,BX_east,FR_east,SM_east,FM_east,PB_east,FN_east,OT_east]

direct_all=pd.concat(direct_dflist,axis=0).reset_index(drop=True)
amzdds_all=pd.concat(amzdds_dflist,axis=0).reset_index(drop=True)

west_all=pd.concat(west_dflist,axis=0).reset_index(drop=True)
east_all=pd.concat(east_dflist,axis=0).reset_index(drop=True)

###### revise Amazon Direct and DDS
direct_all=mattress_df.merge(direct_all,on='Zinus SKU#', how='right')
#### Tracy warehouse get 90% of mattress direct demand
tracy_mattress_direct=direct_all[direct_all['Check']=='mattress'].drop(['Check'],axis=1)
tracy_mattress_direct.iloc[:,1:]=tracy_mattress_direct.iloc[:,1:]*0.9

#### CHS warehouse get 10% of mattress direct demand
chs_mattress_direct=direct_all[direct_all['Check']=='mattress'].drop(['Check'],axis=1)
chs_mattress_direct.iloc[:,1:]=chs_mattress_direct.iloc[:,1:]*0.1
#### 50-50 for ohter categories
half_nonmattress_direct=direct_all[direct_all['Check']!='mattress'].drop(['Check'],axis=1)
half_nonmattress_direct.iloc[:,1:]=half_nonmattress_direct.iloc[:,1:]/2
#### combine direct demand togher                             
tracy_direct_all=pd.concat([tracy_mattress_direct,half_nonmattress_direct],axis=0).reset_index(drop=True)
chs_direct_all=pd.concat([chs_mattress_direct,half_nonmattress_direct],axis=0).reset_index(drop=True)

#### 50-50 for Amazon DDS demand
AMZDDS_all=pd.concat([amzdds_all['Zinus SKU#'],amzdds_all.iloc[:,1:]/2],axis=1).reset_index(drop=True)
                                 
#### writing file
file=pd.ExcelWriter('space_planning.xlsx',engine='xlsxwriter')

tracy_direct_all.to_excel(file,sheet_name='Tracy_AMZ_DIRECT',index=False)
chs_direct_all.to_excel(file,sheet_name='CHS_AMZ_DIRECT',index=False)
AMZDDS_all.to_excel(file,sheet_name='Tracy_AMZ_DDS', index=False)
AMZDDS_all.to_excel(file,sheet_name='CHS_AMZ_DDS', index=False)
west_all.to_excel(file,sheet_name='Tracy_Other_DDS',index=False)
east_all.to_excel(file,sheet_name='CHS_Other_DDS',index=False)

file.save()

