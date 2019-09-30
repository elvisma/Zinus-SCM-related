
library(xlsx)
library(openxlsx)
library(plyr)
options(scipen = 999)

setwd("C:/Users/Elvis Ma/Dropbox (Zinus Inc.)/HomeDepot WKLY/Homedepot Inventory Feed/HD Inv Feed")
#Inv_Feed_Date = "122018"

#Tracy_File_Name <- paste('Feed_Tracy_',Inv_Feed_Date, sep = "")
#CHS_File_Name <-paste('Feed_CHS_',Inv_Feed_Date, sep = "")

Tracy_File_Name <-'Feed_Tracy'
CHS_File_Name <-'Feed_CHS'

CHS_File <- read.xlsx(paste(CHS_File_Name,'.xlsx', sep = ""))
Tracy_File <- read.xlsx(paste(Tracy_File_Name,'.xlsx', sep = ""))


setwd("//192.168.1.41/03.Demand Planning/Elvis/Inventory Feeding")
Default_File_Name <- paste('default_master_website', sep= "")
Default_File <- read.xlsx(paste(Default_File_Name,'.xlsx', sep = ""))


Tracy_File_toMerge <- as.data.frame(Tracy_File[,c(2)])
CHS_File_toMerge <- as.data.frame(CHS_File[,c(2)])
names(Tracy_File_toMerge)="SKU"
names(CHS_File_toMerge)='SKU'

#Create a new variable (column) & assign each element an "ID"
Tracy_File_toMerge$ID <-1:nrow(Tracy_File_toMerge)
CHS_File_toMerge$ID <-1:nrow(CHS_File_toMerge)

# Merge the two data frames using the "SKU" column without sorting
Tracy_File_Merge <- merge(Tracy_File_toMerge, Default_File, by = "SKU", all.x =TRUE,sort = FALSE, suffixes = c(".x",""))

# Order the merged data set using "ID" 
Tracy_File_Merged <- Tracy_File_Merge[order(Tracy_File_Merge$ID),]
Tracy_File_Export <- Tracy_File_Merged[,c(1,3:ncol(Tracy_File_Merge))]

# Merge the two data frames using the "SKU" column without sorting
CHS_File_Merge <- merge(CHS_File_toMerge, Default_File, by = "SKU", all.x =TRUE,sort = FALSE, suffixes = c(".x",""))

# Order the merged data set using "ID" 
CHS_File_Merged <- CHS_File_Merge[order(CHS_File_Merge$ID),]
CHS_File_Export <- CHS_File_Merged[,c(1,3:ncol(CHS_File_Merge))]


setwd("C:/Users/Elvis Ma/Desktop")
write.csv(Tracy_File_Export,"Tracy Inv Feed Check.csv")
write.csv(CHS_File_Export,"CHS Inv Feed Check.csv")


