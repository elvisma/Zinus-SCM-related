
library(xlsx)
library(openxlsx)
library(plyr)
options(scipen = 999)

setwd("C:/Users/Elvis Ma/Desktop")
Date = "022219"

File_Name<- paste('Forecast by channel',Date, sep = " ")

Channel_Forecast <- read.xlsx(paste(File_Name,'.xlsx', sep = ""))



setwd("//zinussvr02/Shared/Accpacadv/Elvis/Inventory Feeding")
Default_File_Name <- paste('default_master', sep= "")
Default_File <- read.xlsx(paste(Default_File_Name,'.xlsx', sep = ""))


#Create a new variable (column) & assign each element an "ID"
Channel_Forecast$ID <-1:nrow(Channel_Forecast)


# Merge the two data frames using the "SKU" column without sorting
File_Merge <- merge(Channel_Forecast, Default_File, by = "SKU", all.x =TRUE,sort = FALSE, suffixes = c(".x",""))

# Order the merged data set using "ID" 
File_Merged <-File_Merge[order(File_Merge$ID),]
File_Export <- File_Merged[,c(1,12:ncol(File_Merge))]



setwd("C:/Users/Elvis Ma/Desktop")
write.csv(File_Export,"channel forecast default.csv",row.names = FALSE)


