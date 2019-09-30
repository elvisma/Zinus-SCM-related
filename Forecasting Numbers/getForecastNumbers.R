#date of this week 12/18/2018
#install.packages('xlsx')
library(xlsx)
#install.packages('openxlsx')
library(openxlsx)
#install.packages('RODBC')

library(excel.link)
#install.packages('dplyr')
library(dplyr)
#install.packages('scales')
library('scales')
library(purrr)
options(scipen = 999)

Date<-"testing"


#Column,need to update (+25)everytime you run

### 9/22
Col_costco = 78+9+2+2+2+2    # Demand tab COSTCO.com 
Col_WF=79+9+2+2+2+2          # Master tab Zinus DDS
Col_SC=77+9+2+2+2+2
Col_WMT = 82+9+2+2+2+2
Col_Zinus=55+9+2+2+2+2
Col_Overstock=83+9+2+2+2+2
Col_target=82+9+2+2+2+2
Col_HD=81+9+2+2+2+2
Col_Macys=33+9+2+2+2+2
Col_Chewy=33+9+2+2+2+2
Col_eBay=33+9+2+2+2+2

# Can be changed regularly
Update_weeks=25


#Costco
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")
File_Name <- paste('costcocom_', Date, sep = "")
CSC_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3)
CSC_SISR <- CSC_SISR_Original[,c(2,4,Col_costco:(Col_costco+Update_weeks))]
CSC_SISR<-filter(CSC_SISR,Account=='COSTCO.COM')[,-2]
colnames(CSC_SISR)[1]<-"SKU"

#WMT

File_Name <- paste('WMT_', Date, sep = "")
WMT_SISR_Original<- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 4, cols=c(4,7,Col_WMT:(Col_WMT+Update_weeks)))
WMT_SISR<-filter(WMT_SISR_Original,Account=='Forecast')[,-2]
colnames(WMT_SISR)[1]<-"SKU"

#Wayfair

File_Name <- paste('USWayfair_', Date, sep = "")
WF_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 4, cols=c(4,6,Col_WF:(Col_WF+Update_weeks)))
WF_SISR<-filter(WF_SISR_Original,Account=='ZINUS DDS')[,-2]
colnames(WF_SISR)[1]<-"SKU"

#Zinus

File_Name <- paste('ZINUSCOM_', Date, sep = "")
ZINUS_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 2, cols=c(4,6,Col_Zinus:(Col_Zinus+Update_weeks)))
ZINUS_SISR<-filter(ZINUS_SISR_Original,Account=='Forecast Number')[,-2]
colnames(ZINUS_SISR)[1]<-"SKU"

#Home Depot

File_Name <- paste('HomeDepot_', Date, sep = "")
HD_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_HD:(Col_HD+Update_weeks)))
HD_SISR<-filter(HD_SISR_Original,Account=='Forecast')[,-2]
colnames(HD_SISR)[1]<-"SKU"

#Overstock

File_Name <- paste('OVERSTOCK_', Date, sep = "")
Overstock_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,9,Col_Overstock:(Col_Overstock+Update_weeks)))
Overstock_SISR<-filter(Overstock_SISR_Original,Account=="Forecast")[,-2]
colnames(Overstock_SISR)[1]<-"SKU"

#Sam's Club

File_Name <- paste('samsclub_',Date, sep = "")
SamsClub_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(2,4,Col_SC:(Col_SC+Update_weeks)))
SamsClub_SISR<-filter(SamsClub_SISR_Original,Account=="Samsclub")[,-2]
colnames(SamsClub_SISR)[1]<-"SKU"

#Target

File_Name <- paste('target_', Date, sep = "")
Target_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_target:(Col_target+Update_weeks)))
Target_SISR<-filter(Target_SISR_Original,Account=="Forecast")[,-2]
colnames(Target_SISR)[1]<-"SKU"

#Macys

File_Name <- paste('MACYS_', Date, sep = "")
Macys_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_Macys:(Col_Macys+Update_weeks)))
Macys_SISR<-filter(Macys_SISR_Original,Account=="Forecast")[,-2]
colnames(Macys_SISR)[1]<-"SKU"


#Chewy

File_Name <- paste('CHEWY_', Date, sep = "")
Chewy_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_Chewy:(Col_Chewy+Update_weeks)))
Chewy_SISR<-filter(Chewy_SISR_Original,Account=="Forecast")[,-2]
colnames(Chewy_SISR)[1]<-"SKU"

#eBay

File_Name <- paste('eBay_', Date, sep = "")
eBay_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_eBay:(Col_eBay+Update_weeks)))
eBay_SISR<-filter(eBay_SISR_Original,Account=="Forecast")[,-2]
colnames(eBay_SISR)[1]<-"SKU"


# data cleaning
CSC_unique <-CSC_SISR[which(!duplicated(CSC_SISR$SKU)),]
WMT_unique <-WMT_SISR[which(!duplicated(WMT_SISR$SKU)),]
WF_unique <-WF_SISR[which(!duplicated(WF_SISR$SKU)),]
ZINUS_unique <-ZINUS_SISR[which(!duplicated(ZINUS_SISR$SKU)),]
HD_unique <-HD_SISR[which(!duplicated(HD_SISR$SKU)),]
Overstock_unique <-Overstock_SISR[which(!duplicated(Overstock_SISR$SKU)),]
SamsClub_unique <-SamsClub_SISR[which(!duplicated(SamsClub_SISR$SKU)),]
Target_unique <-Target_SISR[which(!duplicated(Target_SISR$SKU)),]
Macys_unique <-Macys_SISR[which(!duplicated(Macys_SISR$SKU)),]
Chewy_unique <-Chewy_SISR[which(!duplicated(Chewy_SISR$SKU)),]
eBay_unique <-eBay_SISR[which(!duplicated(eBay_SISR$SKU)),]

#remove NA values in the SISRs
remove_na <-function(df){
  for(i in seq_along(df)){
    df[is.na(df[,i]),i]<-0
  }
  return(df)
}

CSC_clean <- remove_na(CSC_unique)
WMT_clean <- remove_na(WMT_unique)
HD_clean <- remove_na(HD_unique)
Overstock_clean <- remove_na(Overstock_unique)
SamsClub_clean <- remove_na(SamsClub_unique)
Target_clean <- remove_na(Target_unique)
WF_clean <- remove_na(WF_unique)
ZINUS_clean <- remove_na(ZINUS_unique)
Macys_clean <- remove_na(Macys_unique)
Chewy_clean <- remove_na(Chewy_unique)
eBay_clean <- remove_na(eBay_unique)

setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Forecast Numbers")

#Combine files together
xlsx::write.xlsx(CSC_clean, file="FcstNum.xlsx", sheetName="CSC", row.names=FALSE)
xlsx::write.xlsx(WMT_clean, file="FcstNum.xlsx", sheetName="WMT", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(HD_clean, file="FcstNum.xlsx", sheetName="HD", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(Overstock_clean, file="FcstNum.xlsx", sheetName="OS", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(SamsClub_clean, file="FcstNum.xlsx", sheetName="SC", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(Target_clean, file="FcstNum.xlsx", sheetName="Target", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(WF_clean, file="FcstNum.xlsx", sheetName="WFDDS", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(ZINUS_clean, file="FcstNum.xlsx", sheetName="Zinus", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(Macys_clean, file="FcstNum.xlsx", sheetName="Macys", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(Chewy_clean, file="FcstNum.xlsx", sheetName="Chewy", append=TRUE, row.names=FALSE)
xlsx::write.xlsx(eBay_clean, file="FcstNum.xlsx", sheetName="eBay", append=TRUE, row.names=FALSE)

