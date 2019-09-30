
## Sam's Club Demand Forecast Summary Report

library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read and prepare AMZ Price Data
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")

Date = "testing"
Col1 = 85 # Half June of 2019
Col2 = 86


File_Name <- paste('samsclub_', Date, sep = "")
Sams_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 3, cols=c(2,3,5,10,Col1:Col2))
#Overstock_Sales <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Sales Budget', startRow = 32, cols=c(1:18))

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
##########################################################################
Actual <-Sams_SISR[which(Sams_SISR$Account=="Actual Sales"),]

Actual$ID <- 1:nrow(Actual)
Actual<-cbind(Actual, Sams_Retail)
Actual$Amount <-as.numeric(Actual$Value)*as.numeric(Actual$`Retail$`)

##########################################################################
Summary <- data.frame(matrix(nrow=71,ncol=12))

Summary[c(2:6,8:22,24,35,37,49,61),1] <- c("Sam's Club", "A","B","C","Total","Category","Foam Mattresses", "Spring Mattresses", 
                                           "SmartBases", "Frames for Mattress", "Bunk Beds", "Guest Beds", "Frames for Box Spring",
                                           "Box Springs", "Platform Beds", "Furnitures", "Toppers", "Pillows", "Others", "Total", "Top 10", 
                                           "Total","A inaccurate 10","B inaccurate 10","C inaccurate 10")

Summary[c(2,8,24,37,49,61),2] <- "Actual"
Summary[c(2,8,24,37,49,61),4] <- "Forecast"
Summary[c(2,8,24,37,49,61),6] <- "Deviation"
Summary[c(2,8),8] <- "SKU"
Summary[c(37,49,61),7] <-"Inaccuracy Rate"
Summary[c(24,37,49,61),8] <-"Actual #"
Summary[c(24,37,49,61),10] <-"Forecast #"
Summary[c(24,37,49,61),12] <-"Deviation #"



################ 2. ABC Summary ################

######## 1) Actual Sales ######## 
Summary[3,2] <- sum(as.numeric(Actual[which(Actual$Class == 'A'), "Amount"]))
Summary[4,2] <- sum(as.numeric(Actual[which(Actual$Class == 'B'), "Amount"]))
Summary[5,2] <- sum(as.numeric(Actual[which(Actual$Class == 'C'), "Amount"]))
Summary[6,2] <- sum(as.numeric(Summary[3:5,2]))

for(i in 3:5){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[6,2]) 
}

######## 2) Forecast ########
Summary[3,4] <- sum(as.numeric(Sams_Forecast[which(Sams_Forecast$Class == 'A'), 'Amount']))
Summary[4,4] <- sum(as.numeric(Sams_Forecast[which(Sams_Forecast$Class == 'B'), 'Amount']))
Summary[5,4] <- sum(as.numeric(Sams_Forecast[which(Sams_Forecast$Class == 'C'), 'Amount']))
Summary[6,4] <- sum(as.numeric(Summary[3:5,4]))

for(i in 3:6){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
}

######## 3) Deviation ########
for(i in 3:6){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

######## 4) SKU ########
Summary[3,8] <- length(Sams_SISR[which(Sams_SISR$Account == 'Forecast' & Sams_SISR$Class == 'A'), "Value"])
Summary[4,8] <- length(Sams_SISR[which(Sams_SISR$Account == 'Forecast' & Sams_SISR$Class == 'B'), "Value"])
Summary[5,8] <- length(Sams_SISR[which(Sams_SISR$Account == 'Forecast' & Sams_SISR$Class == 'C'), "Value"])
Summary[6,8] <- sum(as.numeric(Summary[3:5,8]))

for(i in 3:5){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[6,8])
}

################ 2. Cateogry Summary ################
Category <- c('Foam Mattresses', 'Spring Mattresses', 'SmartBases', 'Frames for Mattress', 'Bunk Beds', 'Guest Beds', 
              'Frames for Box Spring', 'Box Springs', 'Platform Beds', 'Furnitures', 'Toppers', 'Pillows', 'Others')

######## 1) Actual Sales ######## 
for(i in 9:21){
  Summary[i,2] <- sum(as.numeric(Actual[which(Actual$Sub.Category == Category[i-8]), "Amount"]))
}
Summary[22,2] <- sum(as.numeric(Summary[9:21,2]))

for(i in 9:21){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[22,2])
}

######## 2) Forecast ######## 
for(i in 9:21){
  Summary[i,4] <- sum(as.numeric(Sams_Forecast[which(Sams_Forecast$Sub.Category == Category[i-8]), "Amount"]))
}
Summary[22,4] <- sum(as.numeric(Summary[9:21,4]))

for(i in 9:22){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
  if(!is.finite(Summary[i,5])){
    Summary[i,5] <- NaN
  }
}

######## 3) Deviation ########
for(i in 9:22){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

######## 4) SKU ######## 
for(i in 9:21){
  Summary[i,8] <- length(Sams_SISR[which(Sams_SISR$Account == 'Forecast' & Sams_SISR$Sub.Category == Category[i-8]), "Value"])
}
Summary[22,8] <- sum(as.numeric(Summary[9:21,8]))

for(i in 9:21){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[22,8])
}

################ 2. Top 10 Summary ################


Actual <- Actual[,c("Class","AccountSKU","ID","Value","Amount")]

Actual <- Actual[order(-as.numeric(Actual[,"Amount"])),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)

Top10 <- c(Actual[1:10,"AccountSKU"])

Summary[25:34,1] <- Top10
############################### prepare data for ABC class inaccurate SKUs #################

Forecast_toMerge <- Sams_Forecast[,c("ID","Value","Amount")]

Sams_Comparison <- merge(Actual, Forecast_toMerge, by ="ID", all.x = TRUE)
Sams_Comparison <- Sams_Comparison[,-c(1)]
colnames(Sams_Comparison)[3] <- "Actual Units"
colnames(Sams_Comparison)[4] <- "Actual Amount"
colnames(Sams_Comparison)[5] <- "Forecast Units"
colnames(Sams_Comparison)[6] <- "Forecast Amount"
Sams_Comparison<-Sams_Comparison[order(-as.numeric(Sams_Comparison$`Actual Amount`)),]
#################################dirty code ends #################################

######## 1-1) Actual Sales $ ######## 
Summary[25:34,2] <- Sams_Comparison[1:10,"Actual Amount"]
Summary[35,2] <- sum(as.numeric(Summary[25:34,2]))

Summary[25:34,4] <- Sams_Comparison[1:10,"Forecast Amount"]
Summary[35,4] <- sum(as.numeric(Summary[25:34,4]))

Summary[25:34,8] <- Sams_Comparison[1:10,"Actual Units"]
Summary[35,8] <- sum(as.numeric(Summary[25:34,8]))

Summary[25:34,10] <- Sams_Comparison[1:10,"Forecast Units"]
Summary[35,10] <- sum(as.numeric(Summary[25:34,10]))


for(i in 25:35){
  Summary[i,3] <- as.numeric(Summary[i,2]) / sum(as.numeric(Sams_Comparison[, "Actual Amount"]))
}

for(i in 25:35){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
  if(!is.finite(Summary[i,5])){
    Summary[i,5] <- NaN
  }
}

######## 1-3) Deviation $ ########
for(i in 25:35){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

######## 2-1) Actual Sales units ######## 

for(i in 25:35){
  Summary[i,9] <- as.numeric(Summary[i,8]) / sum(as.numeric(Sams_Comparison[,"Actual Units"]))
}

######## 2-2) Forecast Units ########


for(i in 25:35){
  Summary[i,11] <- (as.numeric(Summary[i,10]) - as.numeric(Summary[i,8])) / as.numeric(Summary[i,10])
  if(!is.finite(Summary[i,11])){
    Summary[i,11] <- NaN
  }
}

######## 2-3) Deviation Units ########
for(i in 25:35){
  Summary[i,12] <- as.numeric(Summary[i,8]) - as.numeric(Summary[i,10])
}


######################  TO   COPY  ################################
################ 2. Top 10 Inaccurate SKUs By Class ################



Sams_Comparison_A <-Sams_Comparison[which(Sams_Comparison$Class=="A"),]
Sams_Comparison_B <-Sams_Comparison[which(Sams_Comparison$Class=="B"),]
Sams_Comparison_C <-Sams_Comparison[which(Sams_Comparison$Class=="C"),]

################### A item ######################################
for(i in 1:nrow(Sams_Comparison_A)){
  Sams_Comparison_A$Rate[i] = 0
  Sams_Comparison_A$Rate[i] <-abs(as.numeric(Sams_Comparison_A[i,"Forecast Amount"])-as.numeric(Sams_Comparison_A[i,"Actual Amount"]))/as.numeric(Sams_Comparison_A[i,"Forecast Amount"])
}

for(i in 1:ncol(Sams_Comparison_A)){
  Sams_Comparison_A[is.na(Sams_Comparison_A[,i]),i]<-0
}

Sams_Comparison_A <- Sams_Comparison_A[order(-as.numeric(Sams_Comparison_A[,"Rate"])),]
Sams_Comparison_A <- data.frame(Sams_Comparison_A)
rownames(Sams_Comparison_A) <- 1:nrow(Sams_Comparison_A)

A10 <- c(Sams_Comparison_A[1:10,"AccountSKU"])
Summary[38:47,1] <- A10

################### B item ######################################
for(i in 1:nrow(Sams_Comparison_B)){
  Sams_Comparison_B$Rate[i] = 0
  Sams_Comparison_B$Rate[i] <-abs(as.numeric(Sams_Comparison_B[i,"Forecast Amount"])-as.numeric(Sams_Comparison_B[i,"Actual Amount"]))/as.numeric(Sams_Comparison_B[i,"Forecast Amount"])
}

for(i in 1:ncol(Sams_Comparison_B)){
  Sams_Comparison_B[is.na(Sams_Comparison_B[,i]),i]<-0
}

Sams_Comparison_B <- Sams_Comparison_B[order(-as.numeric(Sams_Comparison_B[,"Rate"])),]
Sams_Comparison_B <- data.frame(Sams_Comparison_B)
rownames(Sams_Comparison_B) <- 1:nrow(Sams_Comparison_B)

B10 <- c(Sams_Comparison_B[1:10,"AccountSKU"])
Summary[50:59,1] <- B10

################### C item ######################################
for(i in 1:nrow(Sams_Comparison_C)){
  Sams_Comparison_C$Rate[i] = 0
  Sams_Comparison_C$Rate[i] <-abs(as.numeric(Sams_Comparison_C[i,"Forecast Amount"])-as.numeric(Sams_Comparison_C[i,"Actual Amount"]))/as.numeric(Sams_Comparison_C[i,"Forecast Amount"])
}


for(i in 1:ncol(Sams_Comparison_C)){
  Sams_Comparison_C[is.na(Sams_Comparison_C[,i]),i]<-0
}

Sams_Comparison_C <- Sams_Comparison_C[order(-as.numeric(Sams_Comparison_C[,"Rate"])),]
Sams_Comparison_C <- data.frame(Sams_Comparison_C)
rownames(Sams_Comparison_C) <- 1:nrow(Sams_Comparison_C)

C10 <- c(Sams_Comparison_C[1:10,"AccountSKU"])
Summary[62:71,1] <- C10

# this part is for inaccurate SKUs only
Summary[38:47,2] <-Sams_Comparison_A[1:10,"Actual.Amount"]
Summary[38:47,4] <-Sams_Comparison_A[1:10,"Forecast.Amount"]
Summary[38:47,6] <-as.numeric(Summary[38:47,2])-as.numeric(Summary[38:47,4])
Summary[38:47,7] <-Sams_Comparison_A[1:10,"Rate"]
Summary[38:47,8] <-Sams_Comparison_A[1:10,"Actual.Units"]
Summary[38:47,10] <-Sams_Comparison_A[1:10,"Forecast.Units"]
Summary[38:47,12] <- as.numeric(Summary[38:47,8])-as.numeric(Summary[38:47,10])


Summary[50:59,2] <-Sams_Comparison_B[1:10,"Actual.Amount"]
Summary[50:59,4] <-Sams_Comparison_B[1:10,"Forecast.Amount"]
Summary[50:59,6] <-as.numeric(Summary[50:59,2])-as.numeric(Summary[50:59,4])
Summary[50:59,7] <-Sams_Comparison_B[1:10,"Rate"]
Summary[50:59,8] <-Sams_Comparison_B[1:10,"Actual.Units"]
Summary[50:59,10] <-Sams_Comparison_B[1:10,"Forecast.Units"]
Summary[50:59,12] <- as.numeric(Summary[50:59,8])-as.numeric(Summary[50:59,10])


Summary[62:71,2] <-Sams_Comparison_C[1:10,"Actual.Amount"]
Summary[62:71,4] <-Sams_Comparison_C[1:10,"Forecast.Amount"]
Summary[62:71,6] <-as.numeric(Summary[62:71,2])-as.numeric(Summary[62:71,4])
Summary[62:71,7] <-Sams_Comparison_C[1:10,"Rate"]
Summary[62:71,8] <-Sams_Comparison_C[1:10,"Actual.Units"]
Summary[62:71,10] <-Sams_Comparison_C[1:10,"Forecast.Units"]
Summary[62:71,12] <- as.numeric(Summary[62:71,8])-as.numeric(Summary[62:71,10])



############### FINISH COPY  ################################################



setwd("//192.168.1.41/03.Demand Planning/Elvis/Forecast Accuracy Raw Data")
write.csv(Summary,"Sams_Club.csv")
