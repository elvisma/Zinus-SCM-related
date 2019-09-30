
## AMZ Demand Forecast Summary Report

library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read and prepare AMZ Price Data
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")

Date = "testing"

Col1 = 85   # Half June 2019
Col2 = 86


File_Name <- paste('costcocom_', Date, sep = "")
CSC_SISR_Original <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 3)
CSC_SISR <- CSC_SISR_Original[,c(1,2,3,8,9,Col1:Col2)]
#Overstock_Sales <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Sales Budget', startRow = 32, cols=c(1:18))


for(i in 1:ncol(CSC_SISR)){
 
  CSC_SISR[is.na(CSC_SISR[,i]),i] <- 0
}


for (i in 1:nrow(CSC_SISR)){
  CSC_SISR$Units[i] = 0
  CSC_SISR$Units[i]<-sum(as.numeric(CSC_SISR[i,c(6:(ncol(CSC_SISR)))]))
}


CSC_SISR$Amount <- as.numeric(CSC_SISR$New.Unit.Price) * as.numeric(CSC_SISR$Units)

Summary <- data.frame(matrix(nrow=71,ncol=12))

Summary[c(2:6,8:22,24,35,37,49,61),1] <- c("Costco", "A","B","C","Total","Category","Foam Mattresses", "Spring Mattresses", 
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
Summary[3,2] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Actual' & CSC_SISR$ABC == 'A'), 'Amount']))
Summary[4,2] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Actual' & CSC_SISR$ABC == 'B'), 'Amount']))
Summary[5,2] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Actual' & CSC_SISR$ABC == 'C'), 'Amount']))
Summary[6,2] <- sum(as.numeric(Summary[3:5,2]))

for(i in 3:5){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[6,2])
}

######## 2) Forecast ########
Summary[3,4] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'A'), 'Amount']))
Summary[4,4] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'B'), 'Amount']))
Summary[5,4] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'C'), 'Amount']))
Summary[6,4] <- sum(as.numeric(Summary[3:5,4]))

for(i in 3:6){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
}

######## 3) Deviation ########
for(i in 3:6){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

######## 4) SKU ########
Summary[3,8] <- length(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'A'), "Amount"])
Summary[4,8] <- length(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'B'), "Amount"])
Summary[5,8] <- length(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$ABC == 'C'), "Amount"])
Summary[6,8] <- sum(as.numeric(Summary[3:5,8]))

for(i in 3:5){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[6,8])
}

################ 2. Cateogry Summary ################
Category <- c('Foam Mattresses', 'Spring Mattresses', 'SmartBases', 'Frames for Mattress', 'Bunk Beds', 'Guest Beds', 
             'Frames for Box Spring', 'Box Springs', 'Platform Beds', 'Furnitures', 'Toppers', 'Pillows', 'Others')

######## 1) Actual Sales ######## 
for(i in 9:21){
  Summary[i,2] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Actual' & CSC_SISR$Middle.Category == Category[i-8]), "Amount"]))
}
Summary[22,2] <- sum(as.numeric(Summary[9:21,2]))

for(i in 9:21){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[22,2])
}

######## 2) Forecast ######## 
for(i in 9:21){
  Summary[i,4] <- sum(as.numeric(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$Middle.Category == Category[i-8]), "Amount"]))
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
  Summary[i,8] <- length(CSC_SISR[which(CSC_SISR$Account == 'Forecast' & CSC_SISR$Middle.Category == Category[i-8]), "Amount"])
}
Summary[22,8] <- sum(as.numeric(Summary[9:21,8]))

for(i in 9:21){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[22,8])
}

################ 2. Top 10 Summary ################
# Top10 Hot Seller
Actual <- CSC_SISR[which(CSC_SISR$Account == "Actual"),]
Actual <- Actual[order(-as.numeric(Actual[,ncol(Actual)])),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)
Actual_toMerge <-Actual[,c("Zinus.SKU.","Units","Amount")]

Top10 <- c(Actual[1:10,3])
Summary[25:34,1] <- Top10

######## 1-1) Actual Sales $ ######## 
Summary[25:34,2] <- Actual[1:10,ncol(Actual)]
Summary[35,2] <- sum(as.numeric(Summary[25:34,2]))

for(i in 25:35){
  Summary[i,3] <- as.numeric(Summary[i,2]) / sum(as.numeric(Actual[1:nrow(Actual),ncol(Actual)]))
}

######## 1-2) Forecast $ ########
for(i in 25:34){
  Summary[i,4] <- CSC_SISR[which(CSC_SISR$Account == "Forecast" & CSC_SISR$`Zinus.SKU#` == Top10[i-24]), 'Amount']
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
Summary[25:34,8] <- Actual[1:10,"Units"]
Summary[35,8] <- sum(as.numeric(Summary[25:34,8]))

for(i in 25:35){
  Summary[i,9] <- as.numeric(Summary[i,8]) / sum(as.numeric(Actual[1:nrow(Actual),"Units"]))
}

######## 2-2) Forecast Units ########
for(i in 25:34){
  Summary[i,10] <- CSC_SISR[which(CSC_SISR$Account == "Forecast" & CSC_SISR$`Zinus.SKU#` == Top10[i-24]), 'Units']
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


################ 3. Inaccurate SKUs for ABC classes  ############################
###########################################
################ A ITEM ####################
Forecast_A <- CSC_SISR[which(CSC_SISR$Account == "Forecast"&CSC_SISR$ABC == 'A'),]
Forecast_A <- data.frame(Forecast_A)
rownames(Forecast_A) <- 1:nrow(Forecast_A)
Forecast_A_toMerge <- Forecast_A[,c("ABC","Zinus.SKU.","Units","Amount")]
    
CSC_Comparison_A <- merge( Forecast_A_toMerge, Actual_toMerge,by.x = "Zinus.SKU.", by.y ="Zinus.SKU.", all.x = TRUE)

colnames(CSC_Comparison_A)[3] <- "Forecast Units"
colnames(CSC_Comparison_A)[4] <- "Forecast Amount"
colnames(CSC_Comparison_A)[5] <- "Actual Units"
colnames(CSC_Comparison_A)[6] <- "Actual Amount"

for(i in 1:nrow(CSC_Comparison_A)){
  CSC_Comparison_A$Rate[i] = 0
  CSC_Comparison_A$Rate[i] <-abs((CSC_Comparison_A[i,"Forecast Amount"]-CSC_Comparison_A[i,"Actual Amount"]))/CSC_Comparison_A[i,"Forecast Amount"]
}

for(i in 1:ncol(CSC_Comparison_A)){
  CSC_Comparison_A[is.na(CSC_Comparison_A[,i]),i]<-0
}

CSC_Comparison_A <- CSC_Comparison_A[order(-as.numeric(CSC_Comparison_A[,"Rate"])),]
CSC_Comparison_A <- data.frame(CSC_Comparison_A)
rownames(CSC_Comparison_A) <- 1:nrow(CSC_Comparison_A)

A10 <- c(CSC_Comparison_A[1:10,"Zinus.SKU."])
Summary[38:47,1] <- A10
################ B ITEM ############################
Forecast_B <- CSC_SISR[which(CSC_SISR$Account == "Forecast"&CSC_SISR$ABC == 'B'),]
Forecast_B <- data.frame(Forecast_B)
rownames(Forecast_B) <- 1:nrow(Forecast_B)
Forecast_B_toMerge <- Forecast_B[,c("ABC","Zinus.SKU.","Units","Amount")]

CSC_Comparison_B <- merge( Forecast_B_toMerge, Actual_toMerge,by.x = "Zinus.SKU.", by.y ="Zinus.SKU.", all.x = TRUE)

colnames(CSC_Comparison_B)[3] <- "Forecast Units"
colnames(CSC_Comparison_B)[4] <- "Forecast Amount"
colnames(CSC_Comparison_B)[5] <- "Actual Units"
colnames(CSC_Comparison_B)[6] <- "Actual Amount"

for(i in 1:nrow(CSC_Comparison_B)){
  CSC_Comparison_B$Rate[i] = 0
  CSC_Comparison_B$Rate[i] <-abs((CSC_Comparison_B[i,"Forecast Amount"]-CSC_Comparison_B[i,"Actual Amount"]))/CSC_Comparison_B[i,"Forecast Amount"]
}

for(i in 1:ncol(CSC_Comparison_B)){
  CSC_Comparison_B[is.na(CSC_Comparison_B[,i]),i]<-0
}

CSC_Comparison_B <- CSC_Comparison_B[order(-as.numeric(CSC_Comparison_B[,"Rate"])),]
CSC_Comparison_B <- data.frame(CSC_Comparison_B)
rownames(CSC_Comparison_B) <- 1:nrow(CSC_Comparison_B)

B10 <- c(CSC_Comparison_B[1:10,"Zinus.SKU."])
Summary[50:59,1] <- B10
######################### C ITEM  ##################################################################
Forecast_C <- CSC_SISR[which(CSC_SISR$Account == "Forecast"&CSC_SISR$ABC == 'C'),]
Forecast_C <- data.frame(Forecast_C)
rownames(Forecast_C) <- 1:nrow(Forecast_C)
Forecast_C_toMerge <- Forecast_C[,c("ABC","Zinus.SKU.","Units","Amount")]

CSC_Comparison_C <- merge( Forecast_C_toMerge, Actual_toMerge,by.x = "Zinus.SKU.", by.y ="Zinus.SKU.", all.x = TRUE)

colnames(CSC_Comparison_C)[3] <- "Forecast Units"
colnames(CSC_Comparison_C)[4] <- "Forecast Amount"
colnames(CSC_Comparison_C)[5] <- "Actual Units"
colnames(CSC_Comparison_C)[6] <- "Actual Amount"

for(i in 1:nrow(CSC_Comparison_C)){
  CSC_Comparison_C$Rate[i] = 0
  CSC_Comparison_C$Rate[i] <-abs((CSC_Comparison_C[i,"Forecast Amount"]-CSC_Comparison_C[i,"Actual Amount"]))/CSC_Comparison_C[i,"Forecast Amount"]
}

for(i in 1:ncol(CSC_Comparison_C)){
  CSC_Comparison_C[is.na(CSC_Comparison_C[,i]),i]<-0
}

CSC_Comparison_C <- CSC_Comparison_C[order(-as.numeric(CSC_Comparison_C[,"Rate"])),]
CSC_Comparison_C <- data.frame(CSC_Comparison_C)
rownames(CSC_Comparison_C) <- 1:nrow(CSC_Comparison_C)

C10 <- c(CSC_Comparison_C[1:10,"Zinus.SKU."])
Summary[62:71,1] <- C10

# this part is for inaccurate ABC SKUs only
Summary[38:47,2] <- CSC_Comparison_A[1:10,"Actual.Amount"]
Summary[38:47,4] <- CSC_Comparison_A[1:10,"Forecast.Amount"]
Summary[38:47,6] <-as.numeric(Summary[38:47,2])-as.numeric(Summary[38:47,4])
Summary[38:47,7] <- CSC_Comparison_A[1:10,"Rate"]
Summary[38:47,8] <- CSC_Comparison_A[1:10,"Actual.Units"]
Summary[38:47,10] <- CSC_Comparison_A[1:10,"Forecast.Units"]
Summary[38:47,12] <- as.numeric(Summary[38:47,8])-as.numeric(Summary[38:47,10])


Summary[50:59,2] <- CSC_Comparison_B[1:10,"Actual.Amount"]
Summary[50:59,4] <- CSC_Comparison_B[1:10,"Forecast.Amount"]
Summary[50:59,6] <-as.numeric(Summary[50:59,2])-as.numeric(Summary[50:59,4])
Summary[50:59,7] <- CSC_Comparison_B[1:10,"Rate"]
Summary[50:59,8] <- CSC_Comparison_B[1:10,"Actual.Units"]
Summary[50:59,10] <- CSC_Comparison_B[1:10,"Forecast.Units"]
Summary[50:59,12] <- as.numeric(Summary[50:59,8])-as.numeric(Summary[50:59,10])


Summary[62:71,2] <- CSC_Comparison_C[1:10,"Actual.Amount"]
Summary[62:71,4] <- CSC_Comparison_C[1:10,"Forecast.Amount"]
Summary[62:71,6] <-as.numeric(Summary[62:71,2])-as.numeric(Summary[62:71,4])
Summary[62:71,7] <- CSC_Comparison_C[1:10,"Rate"]
Summary[62:71,8] <- CSC_Comparison_C[1:10,"Actual.Units"]
Summary[62:71,10] <- CSC_Comparison_C[1:10,"Forecast.Units"]
Summary[62:71,12] <- as.numeric(Summary[62:71,8])-as.numeric(Summary[62:71,10])






setwd("//192.168.1.41/03.Demand Planning/Elvis/Forecast Accuracy Raw Data")
write.csv(Summary,"Costco.csv")
