## AMZ Demand Forecast Summary Report

library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read and prepare AMZ Price Data
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")

Date = "testing"
Col1 = 56    # Half June of 2019
Col2 = 57

File_Name <- paste('ZINUSCOM_', Date, sep = "")
ZINUS_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 2, cols=c(2,3,4,6,Col1:Col2))
#Overstock_Sales <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Sales Budget', startRow = 32, cols=c(1:18))

ZINUS_SISR <- ZINUS_SISR[-c(1:11),]


for (i in 1:ncol(ZINUS_SISR)){
  ZINUS_SISR[is.na(ZINUS_SISR[,i]),i] <- 0
}


for (i in 1:nrow(ZINUS_SISR)){
  ZINUS_SISR$Amount[i] = 0
  ZINUS_SISR$Amount[i]<-sum(as.numeric(ZINUS_SISR[i,5:ncol(ZINUS_SISR)]))
}

Summary <- data.frame(matrix(nrow=71,ncol=12))

Summary[c(2:6,8:22,24,35,37,49,61),1] <- c("Zinus", "A","B","C","Total","Category","Foam Mattresses", "Spring Mattresses", 
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
Summary[3,2] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == "Actual Sales Total" & ZINUS_SISR$ABC == 'A'), "Amount"]))
Summary[4,2] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Actual Sales Total' & ZINUS_SISR$ABC == 'B'), "Amount"]))
Summary[5,2] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Actual Sales Total' & ZINUS_SISR$ABC == 'C'), "Amount"]))
Summary[6,2] <- sum(as.numeric(Summary[3:5,2]))

for(i in 3:5){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[6,2])
}

######## 2) Forecast ########
Summary[3,4] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Sales Total' & ZINUS_SISR$ABC == 'A'), "Amount"]))
Summary[4,4] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Sales Total' & ZINUS_SISR$ABC == 'B'), "Amount"]))
Summary[5,4] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Sales Total' & ZINUS_SISR$ABC == 'C'), "Amount"]))
Summary[6,4] <- sum(as.numeric(Summary[3:5,4]))

for(i in 3:6){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
}

######## 3) Deviation ########
for(i in 3:6){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

######## 4) SKU ########
Summary[3,8] <- length(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Number' & ZINUS_SISR$ABC == 'A'), "Amount"])
Summary[4,8] <- length(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Number' & ZINUS_SISR$ABC == 'B'), "Amount"])
Summary[5,8] <- length(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Number' & ZINUS_SISR$ABC == 'C'), "Amount"])
Summary[6,8] <- sum(as.numeric(Summary[3:5,8]))

for(i in 3:5){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[6,8])
}

################ 2. Cateogry Summary ################
Category <- c('Foam Mattresses', 'Spring Mattresses', 'SmartBases', 'Frames for Mattress', 'Bunk Beds', 'Guest Beds', 
              'Frames for Box Spring', 'Box Springs', 'Platform Beds', 'Furnitures', 'Toppers', 'Pillows', 'Others')

######## 1) Actual Sales ######## 
for(i in 9:21){
  Summary[i,2] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Actual Sales Total' & ZINUS_SISR$Middle.Category == Category[i-8]), "Amount"]))
}
Summary[22,2] <- sum(as.numeric(Summary[9:21,2]))

for(i in 9:21){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[22,2])
}

######## 2) Forecast ######## 
for(i in 9:21){
  Summary[i,4] <- sum(as.numeric(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Sales Total' & ZINUS_SISR$Middle.Category == Category[i-8]), "Amount"]))
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
  Summary[i,8] <- length(ZINUS_SISR[which(ZINUS_SISR$Account == 'Forecast Number' & ZINUS_SISR$Middle.Category == Category[i-8]), "Amount"])
}
Summary[22,8] <- sum(as.numeric(Summary[9:21,8]))

for(i in 9:21){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[22,8])
}

################ 2. Top 10 Summary ################

Actual_Amount <- ZINUS_SISR[which(ZINUS_SISR$Account == "Actual Sales Total"),]
Actual_Units <- ZINUS_SISR[which(ZINUS_SISR$Account =="Zinus.com"),]
Actual <- merge(Actual_Units[,c("ABC","ZINUS.SKU","Amount")],Actual_Amount[,c("ZINUS.SKU","Amount")], by="ZINUS.SKU", all.x = TRUE)
colnames(Actual)[3] <- "ActualUnits"
colnames(Actual)[4] <- "ActualAmount"
Actual <- Actual[order(-as.numeric(Actual[,"ActualAmount"])),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)

Top10 <- c(Actual[1:10,"ZINUS.SKU"])
Summary[25:34,1] <- Top10


######## 1-1) Actual Sales $ ######## 
Summary[25:34,2] <- Actual[1:10,"ActualAmount"]
Summary[35,2] <- sum(as.numeric(Summary[25:34,2]))

for(i in 25:35){
  Summary[i,3] <- as.numeric(Summary[i,2]) / sum(as.numeric(Actual[1:nrow(Actual),"ActualAmount"]))
}

######## 1-2) Forecast $ ########
for(i in 25:34){
  Summary[i,4] <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Sales Total" & ZINUS_SISR$ZINUS.SKU == Top10[i-24]), "Amount"]
}

Summary[35,4] <- sum(as.numeric(Summary[25:34,4]))

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

######## 2-1) Actual Sales Units ######## 
Summary[25:34,8] <- Actual[1:10,"ActualUnits"]
Summary[35,8] <- sum(as.numeric(Summary[25:34,8]))

for(i in 25:35){
  Summary[i,9] <- as.numeric(Summary[i,8]) / sum(as.numeric(Actual[1:nrow(Actual),"ActualUnits"]))
}

######## 2-2) Forecast Units ########
for(i in 25:34){
  Summary[i,10] <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Number" & ZINUS_SISR$ZINUS.SKU == Top10[i-24]), "Amount"]
}
Summary[35,10] <- sum(as.numeric(Summary[25:34,10]))

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
################ 2. Top 10 Inaccurate SKUs  Summary by Class################
################# A item #####################################

Forecast_A_Amount <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Sales Total"&ZINUS_SISR$ABC == "A"),]
Forecast_A_Units <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Number"&ZINUS_SISR$ABC == "A"),]
Forecast_A <-merge (Forecast_A_Units[,c("ABC","ZINUS.SKU","Amount")],Forecast_A_Amount[,c("ZINUS.SKU","Amount")],by="ZINUS.SKU",all.x=TRUE)
colnames(Forecast_A)[3] <-"ForecastUnits"
colnames(Forecast_A)[4] <-"ForecastAmount"

Forecast_A <- data.frame(Forecast_A)
rownames(Forecast_A) <- 1:nrow(Forecast_A)
ZINUS_Comparison_A <- merge(Actual, Forecast_A[,-c(2)], by.x = "ZINUS.SKU", by.y ="ZINUS.SKU", all.y = TRUE)

for(i in 1:nrow(ZINUS_Comparison_A)){
  ZINUS_Comparison_A$Rate[i] = 0
  ZINUS_Comparison_A$Rate[i] <-abs(as.numeric(ZINUS_Comparison_A[i,"ForecastAmount"])-as.numeric(ZINUS_Comparison_A[i,"ActualAmount"]))/as.numeric(ZINUS_Comparison_A[i,"ForecastAmount"])
}

for(i in 1:ncol(ZINUS_Comparison_A)){
  ZINUS_Comparison_A[is.na(ZINUS_Comparison_A[,i]),i]<-0
}

ZINUS_Comparison_A <- ZINUS_Comparison_A[order(-as.numeric(ZINUS_Comparison_A[,"Rate"])),]
ZINUS_Comparison_A <- data.frame(ZINUS_Comparison_A)
rownames(ZINUS_Comparison_A) <- 1:nrow(ZINUS_Comparison_A)

A10 <- c(ZINUS_Comparison_A[1:10,"ZINUS.SKU"])

Summary[38:47,1] <- A10

################# B item #####################################

Forecast_B_Amount <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Sales Total"&ZINUS_SISR$ABC == "B"),]
Forecast_B_Units <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Number"&ZINUS_SISR$ABC == "B"),]
Forecast_B <-merge (Forecast_B_Units[,c("ABC","ZINUS.SKU","Amount")],Forecast_B_Amount[,c("ZINUS.SKU","Amount")],by="ZINUS.SKU",all.x=TRUE)
colnames(Forecast_B)[3] <-"ForecastUnits"
colnames(Forecast_B)[4] <-"ForecastAmount"

Forecast_B <- data.frame(Forecast_B)
rownames(Forecast_B) <- 1:nrow(Forecast_B)
ZINUS_Comparison_B <- merge(Actual, Forecast_B[,-c(2)], by.x = "ZINUS.SKU", by.y ="ZINUS.SKU", all.y = TRUE)

for(i in 1:nrow(ZINUS_Comparison_B)){
  ZINUS_Comparison_B$Rate[i] = 0
  ZINUS_Comparison_B$Rate[i] <-abs(as.numeric(ZINUS_Comparison_B[i,"ForecastAmount"])-as.numeric(ZINUS_Comparison_B[i,"ActualAmount"]))/as.numeric(ZINUS_Comparison_B[i,"ForecastAmount"])
}

for(i in 1:ncol(ZINUS_Comparison_B)){
  ZINUS_Comparison_B[is.na(ZINUS_Comparison_B[,i]),i]<-0
}

ZINUS_Comparison_B <- ZINUS_Comparison_B[order(-as.numeric(ZINUS_Comparison_B[,"Rate"])),]
ZINUS_Comparison_B <- data.frame(ZINUS_Comparison_B)
rownames(ZINUS_Comparison_B) <- 1:nrow(ZINUS_Comparison_B)

B10 <- c(ZINUS_Comparison_B[1:10,"ZINUS.SKU"])

Summary[50:59,1] <- B10

################# C item #####################################

Forecast_C_Amount <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Sales Total"&ZINUS_SISR$ABC == "C"),]
Forecast_C_Units <- ZINUS_SISR[which(ZINUS_SISR$Account == "Forecast Number"&ZINUS_SISR$ABC == "C"),]
Forecast_C <-merge (Forecast_C_Units[,c("ABC","ZINUS.SKU","Amount")],Forecast_C_Amount[,c("ZINUS.SKU","Amount")],by="ZINUS.SKU",all.x=TRUE)
colnames(Forecast_C)[3] <-"ForecastUnits"
colnames(Forecast_C)[4] <-"ForecastAmount"

Forecast_C <- data.frame(Forecast_C)
rownames(Forecast_C) <- 1:nrow(Forecast_C)
ZINUS_Comparison_C <- merge(Actual, Forecast_C[,-c(2)], by.x = "ZINUS.SKU", by.y ="ZINUS.SKU", all.y = TRUE)

for(i in 1:nrow(ZINUS_Comparison_C)){
  ZINUS_Comparison_C$Rate[i] = 0
  ZINUS_Comparison_C$Rate[i] <-abs(as.numeric(ZINUS_Comparison_C[i,"ForecastAmount"])-as.numeric(ZINUS_Comparison_C[i,"ActualAmount"]))/as.numeric(ZINUS_Comparison_C[i,"ForecastAmount"])
}

for(i in 1:ncol(ZINUS_Comparison_C)){
  ZINUS_Comparison_C[is.na(ZINUS_Comparison_C[,i]),i]<-0
}

ZINUS_Comparison_C <- ZINUS_Comparison_C[order(-as.numeric(ZINUS_Comparison_C[,"Rate"])),]
ZINUS_Comparison_C <- data.frame(ZINUS_Comparison_C)
rownames(ZINUS_Comparison_C) <- 1:nrow(ZINUS_Comparison_C)

C10 <- c(ZINUS_Comparison_C[1:10,"ZINUS.SKU"])

Summary[62:71,1] <- C10

# this part is for inaccurate SKUs only
Summary[38:47,2] <- ZINUS_Comparison_A[1:10,"ActualAmount"]
Summary[38:47,4] <- ZINUS_Comparison_A[1:10,"ForecastAmount"]
Summary[38:47,6] <-as.numeric(Summary[38:47,2])-as.numeric(Summary[38:47,4])
Summary[38:47,7] <- ZINUS_Comparison_A[1:10,"Rate"]
Summary[38:47,8] <- ZINUS_Comparison_A[1:10,"ActualUnits"]
Summary[38:47,10] <- ZINUS_Comparison_A[1:10,"ForecastUnits"]
Summary[38:47,12] <- as.numeric(Summary[38:47,8])-as.numeric(Summary[38:47,10])


Summary[50:59,2] <- ZINUS_Comparison_B[1:10,"ActualAmount"]
Summary[50:59,4] <- ZINUS_Comparison_B[1:10,"ForecastAmount"]
Summary[50:59,6] <-as.numeric(Summary[50:59,2])-as.numeric(Summary[50:59,4])
Summary[50:59,7] <- ZINUS_Comparison_B[1:10,"Rate"]
Summary[50:59,8] <- ZINUS_Comparison_B[1:10,"ActualUnits"]
Summary[50:59,10] <- ZINUS_Comparison_B[1:10,"ForecastUnits"]
Summary[50:59,12] <- as.numeric(Summary[50:59,8])-as.numeric(Summary[50:59,10])


Summary[62:71,2] <- ZINUS_Comparison_C[1:10,"ActualAmount"]
Summary[62:71,4] <- ZINUS_Comparison_C[1:10,"ForecastAmount"]
Summary[62:71,6] <-as.numeric(Summary[62:71,2])-as.numeric(Summary[62:71,4])
Summary[62:71,7] <- ZINUS_Comparison_C[1:10,"Rate"]
Summary[62:71,8] <- ZINUS_Comparison_C[1:10,"ActualUnits"]
Summary[62:71,10] <- ZINUS_Comparison_C[1:10,"ForecastUnits"]
Summary[62:71,12] <- as.numeric(Summary[62:71,8])-as.numeric(Summary[62:71,10])


############### FINISH COPY  ################################################



setwd("//192.168.1.41/03.Demand Planning/Elvis/Forecast Accuracy Raw Data")
write.csv(Summary,"Zinus.com.csv")
