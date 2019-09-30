#date of this week 12/18/2018
#install.packages('xlsx')
library(xlsx)
#install.packages('openxlsx')
library(openxlsx)
#install.packages('RODBC')
library(RODBC)
#install.packages('excel.link')
library(excel.link)
#install.packages('dplyr')
library(dplyr)
#install.packages('scales')
library('scales')
options(scipen = 999)

Date<-"testing"

#Read all the SKUs
setwd("//192.168.1.41/03.Demand Planning/Elvis/Inventory Feeding")
SKUs <- read.xlsx("default_master_SAP.xlsx",cols=1)  ### SAP inventory status list
#Column,need to update (+1)everytime you run
Col_costco = 87+2+1+1+4 #9/22 demand tab
Col_WF=88+2+1+1+4
Col_SC=86+2+1+1+4
Col_WMT = 91+2+1+1+4
Col_Zinus=64+2+1+1+4
Col_Overstock=92+2+1+1+4
Col_target=91+2+1+1+4
Col_HD=90+2+1+1+4
Col_Macys=42+2+1+1+4
Col_eBay=42+2+1+1+4


#Costco
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")
File_Name <- paste('costcocom_', Date, sep = "")
CSC_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3)
CSC_SISR <- CSC_SISR_Original[,c(2,4,Col_costco)]
CSC_SISR<-filter(CSC_SISR,Account=='COSTCO.COM')[,c(1,3)]
names(CSC_SISR) <- c("SKU", "Costco")


#Wayfair

File_Name <- paste('USWayfair_', Date, sep = "")
WF_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 4, cols=c(4,6,Col_WF))
WF_SISR<-filter(WF_SISR_Original,Account=='ZINUS DDS')[,c(1,3)]
names(WF_SISR) <- c("SKU", "Wayfair")

#WMT

File_Name <- paste('WMT_', Date, sep = "")
WMT_SISR_Original<- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 4, cols=c(4,7,Col_WMT))
WMT_SISR<-filter(WMT_SISR_Original,Account=='Forecast')[,c(1,3)]
names(WMT_SISR) <- c("SKU", "Walmart")



#Zinus

File_Name <- paste('ZINUSCOM_', Date, sep = "")
ZINUS_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 2, cols=c(4,6,Col_Zinus))
ZINUS_SISR<-filter(ZINUS_SISR_Original,Account=='Forecast Number')[,c(1,3)]
names(ZINUS_SISR) <- c("SKU", "Zinus")

#Home Depot

File_Name <- paste('HomeDepot_', Date, sep = "")
HD_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_HD))
HD_SISR<-filter(HD_SISR_Original,Account=='Forecast')[,c(1,3)]
names(HD_SISR) <- c("SKU", "HomeDepot")

#Overstock

File_Name <- paste('OVERSTOCK_', Date, sep = "")
Overstock_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,9,Col_Overstock))
Overstock_SISR<-filter(Overstock_SISR_Original,Account=="Forecast")[,c(1,3)]
names(Overstock_SISR) <- c("SKU", "Overstock")

#Sam's Club

File_Name <- paste('samsclub_',Date, sep = "")
SamsClub_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(2,4,Col_SC))
SamsClub_SISR<-filter(SamsClub_SISR_Original,Account=="Samsclub")[,c(1,3)]
names(SamsClub_SISR) <- c("SKU", "SamsClub")

#Target

File_Name <- paste('target_', Date, sep = "")
Target_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_target))
Target_SISR<-filter(Target_SISR_Original,Account=="Forecast")[,c(1,3)]
names(Target_SISR) <- c("SKU", "Target")

#Macys

File_Name <- paste('MACYS_', Date, sep = "")
Macys_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_Macys))
Macys_SISR<-filter(Macys_SISR_Original,Account=="Forecast")[,c(1,3)]
names(Macys_SISR) <- c("SKU", "Macys")

#eBay

File_Name <- paste('eBay_', Date, sep = "")
eBay_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(4,7,Col_eBay))
eBay_SISR<-filter(eBay_SISR_Original,Account=="Forecast")[,c(1,3)]
names(eBay_SISR) <- c("SKU", "eBay")

# data cleaning
CSC_unique <-CSC_SISR[which(!duplicated(CSC_SISR$SKU)),]
WF_unique <-WF_SISR[which(!duplicated(WF_SISR$SKU)),]
WMT_unique <-WMT_SISR[which(!duplicated(WMT_SISR$SKU)),]
ZINUS_unique <-ZINUS_SISR[which(!duplicated(ZINUS_SISR$SKU)),]
HD_unique <-HD_SISR[which(!duplicated(HD_SISR$SKU)),]
Overstock_unique <-Overstock_SISR[which(!duplicated(Overstock_SISR$SKU)),]
SamsClub_unique <-SamsClub_SISR[which(!duplicated(SamsClub_SISR$SKU)),]
Target_unique <-Target_SISR[which(!duplicated(Target_SISR$SKU)),]
Macys_unique <-Macys_SISR[which(!duplicated(Macys_SISR$SKU)),]
eBay_unique <-eBay_SISR[which(!duplicated(eBay_SISR$SKU)),]


#write the report
report<-Reduce(function(...) merge(..., all.x = TRUE, by = "SKU"),list(unique(SKUs),CSC_unique,WF_unique,WMT_unique,ZINUS_unique,HD_unique,Overstock_unique,SamsClub_unique,Target_unique,Macys_unique,eBay_unique))
report[is.na(report)] <- 0
report<-data.frame(report)

#You may change your wd
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Forecast Numbers")
xlsx::write.xlsx(report, file="report_sap.xlsx", sheetName="report QTY", row.names=FALSE)

#####################
sum= rowSums(report[,-1])
round2 = function(x, n) {
  posneg = sign(x)
  z = abs(x)*10^n
  z = ceiling(z)
  z = z/10^n
  z*posneg}
Final_report1 <- cbind(report, Total = rowSums(report[,-1]))

PercentReport <- report
for (i in 1:nrow(PercentReport ))
{ 
  if ( sum[i]==0) PercentReport[i,-1]<-percent(0)
  else
  {
    for (j in 2:ncol(PercentReport))
      PercentReport[i,j]=percent(round2(as.numeric(PercentReport[i,j])/(sum[i]),8))
  }
}
Final_report2 <-PercentReport

xlsx::write.xlsx(Final_report2, file="report_sap_percentage.xlsx", sheetName="report %", row.names=FALSE)
