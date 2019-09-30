
## AMZ Demand Forecast Summary Report

library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read and prepare AMZ Price Data
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/Actual Sales")

Date = "testing"
Col1 = 83   # Half June of 2019
Col2 = 84

File_Name <- paste('WMT_', Date, sep = "")
WMT_SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Demand', startRow = 4, cols=c(2,3,4,6,7,Col1:Col2))


for (i in 1:ncol(WMT_SISR)){
  WMT_SISR[is.na(WMT_SISR[,i]),i] <- 0
}

      
for (i in 1:nrow(WMT_SISR)){
  WMT_SISR$Units[i] = 0
  WMT_SISR$Units[i]<-sum(as.numeric(WMT_SISR[i,6:ncol(WMT_SISR)]))
}

WMT_SISR$Amount <- as.numeric(WMT_SISR$Dropship.Cost) * as.numeric(WMT_SISR$Units)



Summary <- data.frame(matrix(nrow=71,ncol=12))

Summary[c(2:6,8:22,24,35,37,49,61),1] <- c("Walmart", "A","B","C","Total","Category","Foam Mattresses", "Spring Mattresses", 
                                           "SmartBases", "Frames for Mattresses", "Bunk Beds", "Guest Beds", "Frames for Box Spring",
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
Summary[3,2] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Walmart.COM' & WMT_SISR$ABC == 'A'), 'Amount']))
Summary[4,2] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Walmart.COM' & WMT_SISR$ABC == 'B'), 'Amount']))
Summary[5,2] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Walmart.COM' & WMT_SISR$ABC == 'C'), 'Amount']))
Summary[6,2] <- sum(as.numeric(Summary[3:5,2]))

for(i in 3:5){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[6,2])
}

######## 2) Forecast ########
Summary[3,4] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'A'), 'Amount']))
Summary[4,4] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'B'), 'Amount']))
Summary[5,4] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'C'), 'Amount']))
Summary[6,4] <- sum(as.numeric(Summary[3:5,4]))

for(i in 3:6){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
}

######## 3) Deviation ########
for(i in 3:6){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

######## 4) SKU ########
Summary[3,8] <- length(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'A'), "Amount"])
Summary[4,8] <- length(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'B'), "Amount"])
Summary[5,8] <- length(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$ABC == 'C'), "Amount"])
Summary[6,8] <- sum(as.numeric(Summary[3:5,8]))

for(i in 3:5){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[6,8])
}

################ 2. Cateogry Summary ################
Category <- c('Foam Mattresses', 'Spring Mattresses', 'SmartBases', 'Frames for Mattresses', 'Bunk Beds', 'Guest Beds', 
             'Frames for Box Spring', 'Box Springs', 'Platform Beds', 'Furnitures', 'Toppers', 'Pillows', 'Others')

######## 1) Actual Sales ######## 
for(i in 9:21){
  Summary[i,2] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Walmart.COM' & WMT_SISR$Middle.Category == Category[i-8]), "Amount"]))
}
Summary[22,2] <- sum(as.numeric(Summary[9:21,2]))

for(i in 9:21){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[22,2])
}

######## 2) Forecast ######## 
for(i in 9:21){
  Summary[i,4] <- sum(as.numeric(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$Middle.Category == Category[i-8]), "Amount"]))
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
  Summary[i,8] <- length(WMT_SISR[which(WMT_SISR$Account == 'Forecast' & WMT_SISR$Middle.Category == Category[i-8]), "Amount"])
}
Summary[22,8] <- sum(as.numeric(Summary[9:21,8]))

for(i in 9:21){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[22,8])
}

################ 2. Top 10 Summary ################

Actual <- WMT_SISR[which(WMT_SISR$Account == "Walmart.COM"),]
Actual <- Actual[order(-as.numeric(Actual[,"Amount"])),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)

Top10 <- c(Actual[1:10,3])
Summary[25:34,1] <- Top10


######## 1-1) Actual Sales $ ######## 
Summary[25:34,2] <- Actual[1:10,"Amount"]
Summary[35,2] <- sum(as.numeric(Summary[25:34,2]))

for(i in 25:35){
  Summary[i,3] <- as.numeric(Summary[i,2]) / sum(as.numeric(Actual[1:nrow(Actual),"Amount"]))
}

######## 1-2) Forecast $ ########
for(i in 25:34){
  Summary[i,4] <- WMT_SISR[which(WMT_SISR$Account == "Forecast" & WMT_SISR$`Zinus.SKU#` == Top10[i-24]), "Amount"]
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
  Summary[i,10] <- WMT_SISR[which(WMT_SISR$Account == "Forecast" & WMT_SISR$`Zinus.SKU#` == Top10[i-24]), "Units"]
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
################ 2. Top 10 Inaccurate SKUs  by Class ################

Actual_toMerge <-Actual[,c("Zinus.SKU.","Units","Amount")]

#####################  A ITEM ###############################
Forecast_A <- WMT_SISR[which(WMT_SISR$Account == "Forecast"&WMT_SISR$ABC == "A"),]
Forecast_A <- data.frame(Forecast_A)
rownames(Forecast_A) <- 1:nrow(Forecast_A)
Forecast_A_toMerge <- Forecast_A[,c("ABC","Zinus.SKU.","Units","Amount")]

WMT_Comparison_A <- merge( Forecast_A_toMerge, Actual_toMerge,by.x = "Zinus.SKU.", by.y ="Zinus.SKU.", all.x = TRUE)

colnames(WMT_Comparison_A)[3] <- "Forecast Units"
colnames(WMT_Comparison_A)[4] <- "Forecast Amount"
colnames(WMT_Comparison_A)[5] <- "Actual Units"
colnames(WMT_Comparison_A)[6] <- "Actual Amount"

for(i in 1:nrow(WMT_Comparison_A)){
  WMT_Comparison_A $Rate[i] = 0
  WMT_Comparison_A$Rate[i] <-abs(as.numeric(WMT_Comparison_A[i,"Forecast Amount"])-as.numeric(WMT_Comparison_A[i,"Actual Amount"]))/as.numeric(WMT_Comparison_A[i,"Forecast Amount"])
}

for(i in 1:ncol(WMT_Comparison_A)){
  WMT_Comparison_A[is.na(WMT_Comparison_A[,i]),i]<-0
}

WMT_Comparison_A <- WMT_Comparison_A[order(-as.numeric(WMT_Comparison_A[,"Rate"])),]
WMT_Comparison_A <- data.frame(WMT_Comparison_A)
rownames(WMT_Comparison_A) <- 1:nrow(WMT_Comparison_A)

A10 <- c(WMT_Comparison_A[1:10,"Zinus.SKU."])
Summary[38:47,1] <- A10

################## B ITEM #################################

Forecast_B <- WMT_SISR[which(WMT_SISR$Account == "Forecast"&WMT_SISR$ABC == "B"),]
Forecast_B <- data.frame(Forecast_B)
rownames(Forecast_B) <- 1:nrow(Forecast_B)
Forecast_B_toMerge <- Forecast_B[,c("ABC","Zinus.SKU.","Units","Amount")]

WMT_Comparison_B <- merge( Forecast_B_toMerge, Actual_toMerge,by.x = "Zinus.SKU.", by.y ="Zinus.SKU.", all.x = TRUE)

colnames(WMT_Comparison_B)[3] <- "Forecast Units"
colnames(WMT_Comparison_B)[4] <- "Forecast Amount"
colnames(WMT_Comparison_B)[5] <- "Actual Units"
colnames(WMT_Comparison_B)[6] <- "Actual Amount"

for(i in 1:nrow(WMT_Comparison_B)){
  WMT_Comparison_B$Rate[i] = 0
  WMT_Comparison_B$Rate[i] <-abs(as.numeric(WMT_Comparison_B[i,"Forecast Amount"])-as.numeric(WMT_Comparison_B[i,"Actual Amount"]))/as.numeric(WMT_Comparison_B[i,"Forecast Amount"])
}

for(i in 1:ncol(WMT_Comparison_B)){
  WMT_Comparison_B[is.na(WMT_Comparison_B[,i]),i]<-0
}

WMT_Comparison_B <- WMT_Comparison_B[order(-as.numeric(WMT_Comparison_B[,"Rate"])),]
WMT_Comparison_B <- data.frame(WMT_Comparison_B)
rownames(WMT_Comparison_B) <- 1:nrow(WMT_Comparison_B)

B10 <- c(WMT_Comparison_B[1:10,"Zinus.SKU."])
Summary[50:59,1] <- B10
################## C ITEM #################################
Forecast_C <- WMT_SISR[which(WMT_SISR$Account == "Forecast"&WMT_SISR$ABC == "C"),]
Forecast_C <- data.frame(Forecast_C)
rownames(Forecast_C) <- 1:nrow(Forecast_C)
Forecast_C_toMerge <- Forecast_C[,c("ABC","Zinus.SKU.","Units","Amount")]

WMT_Comparison_C <- merge(Forecast_C_toMerge, Actual_toMerge,by.x = "Zinus.SKU.", by.y ="Zinus.SKU.", all.x = TRUE)

colnames(WMT_Comparison_C)[3] <- "Forecast Units"
colnames(WMT_Comparison_C)[4] <- "Forecast Amount"
colnames(WMT_Comparison_C)[5] <- "Actual Units"
colnames(WMT_Comparison_C)[6] <- "Actual Amount"

for(i in 1:nrow(WMT_Comparison_C)){
  WMT_Comparison_C$Rate[i] = 0
  WMT_Comparison_C$Rate[i] <-abs(as.numeric(WMT_Comparison_C[i,"Forecast Amount"])-as.numeric(WMT_Comparison_C[i,"Actual Amount"]))/as.numeric(WMT_Comparison_C[i,"Forecast Amount"])
}

for(i in 1:ncol(WMT_Comparison_C)){
  WMT_Comparison_C[is.na(WMT_Comparison_C[,i]),i]<-0
}

WMT_Comparison_C <- WMT_Comparison_C[order(-as.numeric(WMT_Comparison_C[,"Rate"])),]
WMT_Comparison_C <- data.frame(WMT_Comparison_C)
rownames(WMT_Comparison_C) <- 1:nrow(WMT_Comparison_C)

C10 <- c(WMT_Comparison_C[1:10,"Zinus.SKU."])
Summary[62:71,1] <- C10

###############################################

# this part is for inaccurate ABC SKUs only
Summary[38:47,2] <- WMT_Comparison_A[1:10,"Actual.Amount"]
Summary[38:47,4] <- WMT_Comparison_A[1:10,"Forecast.Amount"]
Summary[38:47,6] <-as.numeric(Summary[38:47,2])-as.numeric(Summary[38:47,4])
Summary[38:47,7] <- WMT_Comparison_A[1:10,"Rate"]
Summary[38:47,8] <- WMT_Comparison_A[1:10,"Actual.Units"]
Summary[38:47,10] <- WMT_Comparison_A[1:10,"Forecast.Units"]
Summary[38:47,12] <- as.numeric(Summary[38:47,8])-as.numeric(Summary[38:47,10])


Summary[50:59,2] <- WMT_Comparison_B[1:10,"Actual.Amount"]
Summary[50:59,4] <- WMT_Comparison_B[1:10,"Forecast.Amount"]
Summary[50:59,6] <-as.numeric(Summary[50:59,2])-as.numeric(Summary[50:59,4])
Summary[50:59,7] <- WMT_Comparison_B[1:10,"Rate"]
Summary[50:59,8] <- WMT_Comparison_B[1:10,"Actual.Units"]
Summary[50:59,10] <- WMT_Comparison_B[1:10,"Forecast.Units"]
Summary[50:59,12] <- as.numeric(Summary[50:59,8])-as.numeric(Summary[50:59,10])


Summary[62:71,2] <- WMT_Comparison_C[1:10,"Actual.Amount"]
Summary[62:71,4] <- WMT_Comparison_C[1:10,"Forecast.Amount"]
Summary[62:71,6] <-as.numeric(Summary[62:71,2])-as.numeric(Summary[62:71,4])
Summary[62:71,7] <- WMT_Comparison_C[1:10,"Rate"]
Summary[62:71,8] <- WMT_Comparison_C[1:10,"Actual.Units"]
Summary[62:71,10] <- WMT_Comparison_C[1:10,"Forecast.Units"]
Summary[62:71,12] <- as.numeric(Summary[62:71,8])-as.numeric(Summary[62:71,10])



############### FINISH COPY  ################################################



setwd("//192.168.1.41/03.Demand Planning/Elvis/Forecast Accuracy Raw Data")
write.csv(Summary,"Walmart.csv")
