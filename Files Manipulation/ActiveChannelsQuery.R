
library(xlsx)
library(openxlsx)
library(plyr)
options(scipen = 999)

###### ONLY THING TO CHANGE IS DATE ######
Inv_Feed_Date = "122818"

setwd("//zinussvr02/Shared/Accpacadv/Elvis/R code Active channels")
Tracy_File_Name <- paste('Feed_Tracy_',Inv_Feed_Date, sep = "")
CHS_File_Name <-paste('Feed_CHS_',Inv_Feed_Date, sep = "")

CHS_File <- read.xlsx(paste(CHS_File_Name,'.xlsx', sep = ""))
Tracy_File <- read.xlsx(paste(Tracy_File_Name,'.xlsx', sep = ""))


setwd("//zinussvr02/Shared/Accpacadv/Elvis/Inventory Feeding")
Default_File_Name <- paste('default_master', sep= "")
Default_File <- read.xlsx(paste(Default_File_Name,'.xlsx', sep = ""))


Tracy_File_toMerge <- Tracy_File[,c(2,20:ncol(Tracy_File)-1)]
CHS_File_toMerge <- CHS_File[,c(2,20:ncol(CHS_File)-1)]

#Create a new variable (column) & assign each element an "ID"
Tracy_File_toMerge$ID <-1:nrow(Tracy_File_toMerge)
CHS_File_toMerge$ID <-1:nrow(CHS_File_toMerge)

# Merge the two data frames using the "SKU" column without sorting
Tracy_File_Merge <- merge(Tracy_File_toMerge, Default_File, by = "SKU", all.x =TRUE,sort = FALSE, suffixes = c(".x",""))

# Order the merged data set using "ID" 
Tracy_File_Merged <- Tracy_File_Merge[order(Tracy_File_Merge$ID),]
Tracy_File_Export <- Tracy_File_Merged[,c(1,14:ncol(Tracy_File_Merge))]

# Merge the two data frames using the "SKU" column without sorting
CHS_File_Merge <- merge(CHS_File_toMerge, Default_File, by = "SKU", all.x =TRUE,sort = FALSE, suffixes = c(".x",""))

# Order the merged data set using "ID" 
CHS_File_Merged <- CHS_File_Merge[order(CHS_File_Merge$ID),]
CHS_File_Export <- CHS_File_Merged[,c(1,14:ncol(CHS_File_Merge))]


setwd("//zinussvr02/Shared/Accpacadv/Elvis/R code Active channels")
xlsx::write.xlsx(Tracy_File_Export,"Active Channels by Location.xlsx", sheetName = "Tracy",row.names = FALSE)
xlsx::write.xlsx(CHS_File_Export,"Active Channels by Location.xlsx", sheetName = "CHS", append = TRUE, row.names = FALSE)


