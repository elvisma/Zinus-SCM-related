library(xlsx)
library(openxlsx)
library(readr)
options(scipen = 999)

###############Elvis Created ##########################

###### working with Dates ###########
setwd("C:/Users/Elvis Ma/Desktop/Elvis Local/Young/2) Wednesday/AMZ Dashobard/Update AMZ_DDS_Direct (every 2 weeks)")
SH_data <-readr::read_csv("SH.csv")
col_list <- colnames(SH_data)
str(col_list)
column_date <- as.Date(col_list[4],format = "%m/%d/%Y")
