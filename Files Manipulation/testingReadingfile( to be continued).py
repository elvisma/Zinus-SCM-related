# -*- coding: utf-8 -*-
"""
Created on Thu Jul 25 13:27:11 2019

@author: Elvis Ma
"""


import os
import re
import glob
import pandas as pd

inventory_df=pd.DataFrame()
source="//192.168.1.41/03.Demand Planning/SCM Team"
os.chdir(source)
dir_list=os.listdir(source)
lookup_list=[folder for folder in dir_list if re.match("SISR", folder)]   ### import re to filter list
lookup_directory_list=[source+"/"+item+"/" for item in lookup_list]
inventory_directory_list=[item+'Inventory Feed' for item in lookup_directory_list]

for inventory_directory in inventory_directory_list:
    filename_list=os.listdir(inventory_directory)### list the excel file names
    if filename_list:  ### cannot find some folder exception
        
        
        
        
        

filename_list=[]
for i in range(len(inventory_directory_list)):
    single_directory=inventory_directory_list[i]
    for file in os.listdir(single_directory):
        book=pd.ExcelFile(file)


### testing below ####
for SISR_folder in glob.glob("SISR 11*/"):
    for inventory_folder in glob.glob("Inventory*"):
        for inventory_file in inventory_folder:
            df_piece=pd.read_excel(inventory_file)
            inventory_df=inventory_df.append(df_piece)

for inventory_file in glob.glob("\\192.168.1.41\03.Demand Planning\SCM Team\SISR 11.13.2018\Inventory Feed"):
    df_piece=pd.read_excel(inventory_file)
    inventory_df=inventory_df.append(df_piece)

for file in inventory_directory_list[0]:
    df_piece=pd.read_excel(file)
    inventory_df.append(df_piece)


filename_list=os.listdir(inventory_directory_list[0])


