# -*- coding: utf-8 -*-
"""
Created on Sat May 25 17:57:47 2019

@author: Elvis Ma
"""

import pandas as pd
import numpy as np
import os

wd="C:/Users/Elvis Ma/Desktop/Weekly Work/Forecast Numbers"
os.chdir(wd)
fcst_combine=pd.DataFrame()
fcst_book=pd.ExcelFile("FcstNum.xlsx")
for sheet in fcst_book.sheet_names:
    df=pd.read_excel(fcst_book, sheet_name=sheet)
    fcst_combine=fcst_combine.append(df)
    
fcst_combine.to_excel("fcst_combine.xlsx", sheet_name="all", index=False)