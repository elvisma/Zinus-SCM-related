library(xlsx)
library(openxlsx)
library(dplyr)
library(XLConnect)
options(scipen = 999)

setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Forecast Numbers")

FcstNum_workbook <- loadWorkbook('FcstNum.xlsx')
sheetNames_list<-getSheets(FcstNum_workbook)
#put sheets into a list of data frames
sheet_list <-lapply(sheetNames_list, function(.sheet){readWorksheet(object=FcstNum_workbook,.sheet)})

df <-data.frame()
for (i in 1: length(sheet_list)){
  data_append <-sheet_list[[i]]
  df <-rbind(df,data_append)
}

