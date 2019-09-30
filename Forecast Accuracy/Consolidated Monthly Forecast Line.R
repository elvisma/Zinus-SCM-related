# All channels Monthly Forecast Line

library(xlsx)
library(openxlsx)
options(scipen = 999)
## Read Files


## 2. Costco

Costco_Date = "testing"
Costco_Col1 = 80+5  # Costco June
Costco_Col2 = 84+5

## 8. Wayfair

WF_Date = "testing"
WF_Col1 = 78+5   # Wayfair June
WF_Col2 = 82+5

## 7. Sam's

Sams_Date = "testing"
Sams_Col1 = 80+5 # Sam's June
Sams_Col2 = 84+5

## 4. Walmart

WMT_Date = "testing"
WMT_Col1 = 78+5    # WMT June
WMT_Col2 = 82+5

## 6. Zinus

Zinus_Date = "testing"
Zinus_Col1 = 51+5  # Zinus.com June
Zinus_Col2 = 55+5

## 1. Overstock 

Overstock_Date = "testing"
Overstock_Col1 = 79+5  #Overstock June
Overstock_Col2 = 83+5


## 3. Target

Target_Date = "testing"
Target_Col1 = 78+5  # Target June
Target_Col2 = 82+5



## 5. Homedepot

HD_Date = "testing"
HD_Col1 = 77+5   # HD June
HD_Col2 = 81+5

## 9. Macys
Macys_Date = "testing"
Macys_Col1 = 29+5   # Macys June
Macys_Col2 = 33+5

## 10. CHEWY
CHEWY_Date = "testing"
CHEWY_Col1 = 29+5   # Chewy June
CHEWY_Col2 = 33+5


####### Costco Files ############
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")
File_Name <- paste('costcocom_', Costco_Date, sep = "")
CSC_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 3)
CSC_SISR <- CSC_SISR_Original[,c(1,2,3,7,9,Costco_Col1:Costco_Col2)]

for(i in 1:ncol(CSC_SISR)){
  
  CSC_SISR[is.na(CSC_SISR[,i]),i] <- 0
}
for (i in 1:nrow(CSC_SISR)){
  CSC_SISR$Units[i] = 0
  CSC_SISR$Units[i]<-sum(as.numeric(CSC_SISR[i,c(6:(ncol(CSC_SISR)))]))
}

CSC_SISR$Amount <- as.numeric(CSC_SISR$Unit.Price) * as.numeric(CSC_SISR$Units)

####### Wayfair File ##############

File_Name <- paste('USWayfair_', WF_Date, sep = "")
WF_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 2, cols=c(3,5,6,8,WF_Col1:WF_Col2))
for (i in 1:ncol(WF_SISR)){
  WF_SISR[is.na(WF_SISR[,i]),i] <- 0
}
for (i in 1:nrow(WF_SISR)){
  WF_SISR$Value[i] = 0
  WF_SISR$Value[i]<-sum(as.numeric(WF_SISR[i,5:ncol(WF_SISR)]))
}

########## Sam's Club File######################

File_Name <- paste('SAMSCLUB_', Sams_Date, sep = "")
Sams_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 3, cols=c(2,3,5,10,Sams_Col1:Sams_Col2))

for (i in 1:ncol(Sams_SISR)){
  Sams_SISR[is.na(Sams_SISR[,i]),i] <- 0
}


for (i in 1:nrow(Sams_SISR)){
  Sams_SISR$Value[i] = 0
  Sams_SISR$Value[i]<-sum(as.numeric(Sams_SISR[i,5:ncol(Sams_SISR)]))
}

Sams_Forecast <- Sams_SISR[which(Sams_SISR$Account == "Forecast"),]
Sams_Retail <- Sams_SISR[which(Sams_SISR$Account == "Retail $"), ]

Sams_Forecast$ID <-1:nrow(Sams_Forecast)
Sams_Retail$ID <- 1:nrow(Sams_Retail)
Sams_Forecast <- Sams_Forecast[,c(1,2,3,ncol(Sams_Forecast)-1,ncol(Sams_Forecast))]
Sams_Retail <-Sams_Retail[,c(ncol(Sams_Retail)-2,ncol(Sams_Retail))]
Sams_Forecast <- merge(Sams_Forecast, Sams_Retail, by = 'ID', all.x = TRUE)
colnames(Sams_Forecast)[ncol(Sams_Forecast)] <- "Retail$"
colnames(Sams_Retail)[1] <- "Retail$"
Sams_Forecast$Amount <-as.numeric(Sams_Forecast$Value)*as.numeric(Sams_Forecast$`Retail$`)
for (i in 1:ncol(Sams_Forecast)){
  Sams_Forecast[is.na(Sams_Forecast[,i]),i] <- 0
}

################ WMT file#############

File_Name <- paste('WMT_', WMT_Date, sep = "")
WMT_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 4, cols=c(2,3,4,6,7,WMT_Col1:WMT_Col2))
for (i in 1:ncol(WMT_SISR)){
  WMT_SISR[is.na(WMT_SISR[,i]),i] <- 0
}

for (i in 1:nrow(WMT_SISR)){
  WMT_SISR$Units[i] = 0
  WMT_SISR$Units[i]<-sum(as.numeric(WMT_SISR[i,6:ncol(WMT_SISR)]))
}

WMT_SISR$Amount <- as.numeric(WMT_SISR$Dropship.Cost) * as.numeric(WMT_SISR$Units)


############### Zinus File ################

File_Name <- paste('ZINUSCOM_', Zinus_Date, sep = "")
ZINUS_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 2, cols=c(2,3,4,6,Zinus_Col1:Zinus_Col2))
#Overstock_Sales <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Sales Budget', startRow = 32, cols=c(1:18))

ZINUS_SISR <- ZINUS_SISR[-c(1:11),]

for (i in 1:ncol(ZINUS_SISR)){
  ZINUS_SISR[is.na(ZINUS_SISR[,i]),i] <- 0
}

for (i in 1:nrow(ZINUS_SISR)){
  ZINUS_SISR$Amount[i] = 0
  ZINUS_SISR$Amount[i]<-sum(as.numeric(ZINUS_SISR[i,5:ncol(ZINUS_SISR)]))
}

############ Overstock file#################

File_Name <- paste('OVERSTOCK_', Overstock_Date, sep = "")
Overstock_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(1,3,4,6,9,Overstock_Col1:Overstock_Col2))
for (i in 1:ncol(Overstock_SISR)){
  Overstock_SISR[is.na(Overstock_SISR[,i]),i] <- 0
}
for (i in 1:nrow(Overstock_SISR)){
  Overstock_SISR$Units[i]=0
  Overstock_SISR$Units[i]<-sum(as.numeric(Overstock_SISR[i,6:(ncol(Overstock_SISR))]))
}
colnames(Overstock_SISR)[4]<-"Unit.Price"
Overstock_SISR$Amount <- as.numeric(Overstock_SISR$Unit.Price) * as.numeric(Overstock_SISR$Units)

######################## Target File ##########################

File_Name <- paste('target_', Target_Date, sep = "")
Target_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(2,3,4,6,7,Target_Col1:Target_Col2))

for(i in 1:ncol(Target_SISR)){
  
  Target_SISR[is.na(Target_SISR[,i]),i] <- 0
}

for (i in 1:nrow(Target_SISR)){
  Target_SISR$Units[i] = 0
  Target_SISR$Units[i]<-sum(as.numeric(Target_SISR[i,c(6:(ncol(Target_SISR)))]))
}

Target_SISR[which(Target_SISR$Unit.Price=="Missing"),"Unit Price"] <-0
Target_SISR$Amount <- as.numeric(Target_SISR$Unit.Price) * as.numeric(Target_SISR$Units)

##################### Homedepot File ##########################

File_Name <- paste('HomeDepot_', HD_Date, sep = "")
HD_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(2,3,4,6,7,HD_Col1:HD_Col2))

for (i in 1:ncol(HD_SISR)){
  HD_SISR[is.na(HD_SISR[,i]),i] <- 0
}

for (i in 1:nrow(HD_SISR)){
  HD_SISR$Units[i] = 0
  HD_SISR$Units[i]<-sum(as.numeric(HD_SISR[i,6:ncol(HD_SISR)]))
}
HD_SISR$Amount <- as.numeric(HD_SISR$Unit.Price) * as.numeric(HD_SISR$Units)

##################### Macys File ##########################
File_Name <- paste('MACYS_', Macys_Date, sep = "")
Macys_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(2,3,4,6,7,Macys_Col1:Macys_Col2))

for (i in 1:ncol(Macys_SISR)){
  Macys_SISR[is.na(Macys_SISR[,i]),i] <- 0
}
for (i in 1:nrow(Macys_SISR)){
  Macys_SISR$Units[i]=0
  Macys_SISR$Units[i]<-sum(as.numeric(Macys_SISR[i,6:(ncol(Macys_SISR))]))
}
Macys_SISR$Amount <- as.numeric(Macys_SISR$Unit.Price) * as.numeric(Macys_SISR$Units)
###################

##################### Chewy File ##########################
File_Name <- paste('CHEWY_', CHEWY_Date, sep = "")
CHEWY_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 3, cols=c(2,3,4,6,7,CHEWY_Col1:CHEWY_Col2))

for (i in 1:ncol(CHEWY_SISR)){
  CHEWY_SISR[is.na(CHEWY_SISR[,i]),i] <- 0
}
for (i in 1:nrow(CHEWY_SISR)){
  CHEWY_SISR$Units[i]=0
  CHEWY_SISR$Units[i]<-sum(as.numeric(CHEWY_SISR[i,6:(ncol(CHEWY_SISR))]))
}
CHEWY_SISR$Amount <- as.numeric(CHEWY_SISR$Unit.Price) * as.numeric(CHEWY_SISR$Units)
###################
####### Matrix########
Consolidated_Forecast_Line <- data.frame(matrix(nrow=59,ncol=2))
Consolidated_Forecast_Line[c(2,8,14,20,26,32,38,44,50,56),1] <- "A"
Consolidated_Forecast_Line[c(3,9,15,21,27,33,39,45,51,57),1] <- "B"
Consolidated_Forecast_Line[c(4,10,16,22,28,34,40,46,52,58),1] <- "C"
Consolidated_Forecast_Line[c(5,11,17,23,29,35,41,47,53,59),1] <- "Total"
Consolidated_Forecast_Line[c(1,7,13,19,25,31,37,43,49,55),1] <- c("Costco","Wayfair","Sam's Clus","Walmart","Zinus.com","Overstock","Target","Homedepot","Macys","CHEWY.COM")
Consolidated_Forecast_Line[c(1,7,13,19,25,31,37,43,49,55),2] <-"Monthly Forecast Line"

########## Costco ###########
Consolidated_Forecast_Line[2,2] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[3,2] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[4,2] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[5,2] <- sum(as.numeric(Consolidated_Forecast_Line[2:4,2]))

######## Wayfair ###########
Consolidated_Forecast_Line[8,2] <- sum(as.numeric(WF_SISR[which(WF_SISR$Account == 'Forecast Sales Total' & WF_SISR$ABC == 'A'), "Value"]))
Consolidated_Forecast_Line[9,2] <- sum(as.numeric(WF_SISR[which(WF_SISR$Account == 'Forecast Sales Total' & WF_SISR$ABC == 'B'), "Value"]))
Consolidated_Forecast_Line[10,2] <- sum(as.numeric(WF_SISR[which(WF_SISR$Account == 'Forecast Sales Total' & WF_SISR$ABC == 'C'), "Value"]))
Consolidated_Forecast_Line[11,2] <- sum(as.numeric(Consolidated_Forecast_Line[8:10,2]))

######## Sam's Club##########
Consolidated_Forecast_Line[14,2] <- sum(as.numeric(Sams_Forecast[which(Sams_Forecast$Class == 'A'), "Amount"]))
Consolidated_Forecast_Line[15,2] <- sum(as.numeric(Sams_Forecast[which(Sams_Forecast$Class == 'B'), "Amount"]))
Consolidated_Forecast_Line[16,2] <- sum(as.numeric(Sams_Forecast[which(Sams_Forecast$Class == 'C'), "Amount"]))
Consolidated_Forecast_Line[17,2] <- sum(as.numeric(Consolidated_Forecast_Line[14:16,2]))

########## Walmart #########
Consolidated_Forecast_Line[20,2] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[21,2] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[22,2] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[23,2] <- sum(as.numeric(Consolidated_Forecast_Line[20:22,2]))

########## Zinus.com ########
Consolidated_Forecast_Line[26,2] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Sales Total' & ZINUS_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[27,2] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Sales Total' & ZINUS_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[28,2] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Sales Total' & ZINUS_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[29,2] <- sum(as.numeric(Consolidated_Forecast_Line[26:28,2]))

######### Overstock #########
Consolidated_Forecast_Line[32,2] <- sum(as.numeric(Overstock_SISR[which(Overstock_SISR$Account == 'Forecast' & Overstock_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[33,2] <- sum(as.numeric(Overstock_SISR[which(Overstock_SISR$Account == 'Forecast' & Overstock_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[34,2] <- sum(as.numeric(Overstock_SISR[which(Overstock_SISR$Account == 'Forecast' & Overstock_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[35,2] <- sum(as.numeric(Consolidated_Forecast_Line[32:34,2]))

########## Target ###########
Consolidated_Forecast_Line[38,2] <- sum(as.numeric(Target_SISR[which(Target_SISR$Account == 'Forecast' & Target_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[39,2] <- sum(as.numeric(Target_SISR[which(Target_SISR$Account == 'Forecast' & Target_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[40,2] <- sum(as.numeric(Target_SISR[which(Target_SISR$Account == 'Forecast' & Target_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[41,2] <- sum(as.numeric(Consolidated_Forecast_Line[38:40,2]))

########## Homedepot ##########
Consolidated_Forecast_Line[44,2] <- sum(as.numeric(HD_SISR[which(HD_SISR$Account == 'Forecast' & HD_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[45,2] <- sum(as.numeric(HD_SISR[which(HD_SISR$Account == 'Forecast' & HD_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[46,2] <- sum(as.numeric(HD_SISR[which(HD_SISR$Account == 'Forecast' & HD_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[47,2] <- sum(as.numeric(Consolidated_Forecast_Line[44:46,2]))

########## Macys ##########
Consolidated_Forecast_Line[50,2] <- sum(as.numeric(Macys_SISR[which(Macys_SISR$Account == 'Forecast' & Macys_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[51,2] <- sum(as.numeric(Macys_SISR[which(Macys_SISR$Account == 'Forecast' & Macys_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[52,2] <- sum(as.numeric(Macys_SISR[which(Macys_SISR$Account == 'Forecast' & Macys_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[53,2] <- sum(as.numeric(Consolidated_Forecast_Line[50:52,2]))

########## Chewy ##########
Consolidated_Forecast_Line[56,2] <- sum(as.numeric(CHEWY_SISR[which(CHEWY_SISR$Account == 'Forecast' & CHEWY_SISR$ABC == 'A'), "Amount"]))
Consolidated_Forecast_Line[57,2] <- sum(as.numeric(CHEWY_SISR[which(CHEWY_SISR$Account == 'Forecast' & CHEWY_SISR$ABC == 'B'), "Amount"]))
Consolidated_Forecast_Line[58,2] <- sum(as.numeric(CHEWY_SISR[which(CHEWY_SISR$Account == 'Forecast' & CHEWY_SISR$ABC == 'C'), "Amount"]))
Consolidated_Forecast_Line[59,2] <- sum(as.numeric(Consolidated_Forecast_Line[56:58,2]))

########### Write File#################
setwd("//192.168.1.41/03.Demand Planning/Elvis/Forecast Accuracy Raw Data")
write.csv(Consolidated_Forecast_Line,"Monthly Forecast Line.csv")
