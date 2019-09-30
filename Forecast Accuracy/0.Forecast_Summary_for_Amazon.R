
## AMZ Demand Forecast Summary Report

library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read and prepare AMZ Price Data
setwd("//192.168.1.41/09.Sales/David L/Amazon Demand Forecast")

Combined <- read.xlsx('AMZ Demand Forecast_All_06172019.xlsm', sheet = 'Master', startRow=3)

Col1=36  # half June 2019
Col2=37

Combined <- Combined[,c(1,4:6,11,13,Col1:Col2)]




for(i in 1:ncol(Combined)){
  Combined[is.na(Combined[,i]),i]<-0
}




Combined <- Combined[,-c(2,3)]



for (i in 1:ncol(Combined)){
  Combined[is.na(Combined[,i]),i]<-0
}


for (i in 1:nrow(Combined)){
  
  Combined$Amount[i] = 0
  Combined$Amount[i]<-sum(as.numeric(Combined[i,c(5:ncol(Combined))]))
}




# Combined$Amount<-0
# for (i in 1:nrow(Combined)){
  # Combined$Amount[i]<-sum(as.numeric(Combined[i,c(5:ncol(Combined))]))
# }

# Combined$Amount<-as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)
# Combined$Amount<-as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)
# Combined$Amount<-as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)

# FM.SM$Amount<-as.numeric(FM.SM$`Week.-4`)+as.numeric(FM.SM$`Week.-3`)+as.numeric(FM.SM$`Week.-2`)+as.numeric(FM.SM$`Week.-1`)
# Combined$Amount<-as.numeric(Combined$`Week.-4`)+as.numeric(Combined$`Week.-3`)+as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)
# Combined$Amount<-as.numeric(Combined$`Week.-4`)+as.numeric(Combined$`Week.-3`)+as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)

# Combined$Amount<-as.numeric(Combined$`Week.-5`)+as.numeric(Combined$`Week.-4`)+as.numeric(Combined$`Week.-3`)+as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)
# Combined$Amount<-as.numeric(Combined$`Week.-5`)+as.numeric(Combined$`Week.-4`)+as.numeric(Combined$`Week.-3`)+as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)
# Combined$Amount<-as.numeric(Combined$`Week.-5`)+as.numeric(Combined$`Week.-4`)+as.numeric(Combined$`Week.-3`)+as.numeric(Combined$`Week.-2`)+as.numeric(Combined$`Week.-1`)


Summary <- data.frame(matrix(nrow=92,ncol=12))

Summary[c(3:7,9:13,15:19,21:25,27:31,33:37,39:43,45,56,58,70,82),1] <- c("Total","A","B","C","Total","Foam Mattress","A","B","C","Total", 
                                  "Spring Mattress","A","B","C","Total","Frame","A","B","C","Total","Box Spring","A","B","C","Total",
                                  "Platform Bed","A","B","C","Total","Others","A","B","C","Total","Top 20","Total","A inaccurate 10","B inaccurate 10","C inaccurate 10")

Summary[c(3,9,15,21,27,33,39,45,58,70,82),2] <- "Actual"
Summary[c(3,9,15,21,27,33,39,45,58,70,82),4] <- "Forecast"
Summary[c(3,9,15,21,27,33,39,45,58,70,82),6] <- "Deviation"
Summary[c(58,70,82),7] <-"Inaccurancy Rate"
Summary[c(3,9,15,21,27,33,39),8] <- "SKU"
Summary[c(45,58,70,82),8] <-"Actual #"
Summary[c(45,58,70,82),10] <- "Forecast #"
Summary[c(45,58,70,82),12] <- "Deviation #"


#Summary <- data.frame(matrix(nrow=34,ncol=3))
#colnames(Summary) <- c('Actual', 'Forecast', 'SKU')

# Acutal $

Summary[4,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$Class == 'A'), "Amount"]))
Summary[5,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$Class == 'B'), "Amount"]))
Summary[6,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$Class == 'C'), "Amount"]))
Summary[7,2] <- sum(as.numeric(Summary[4:6,2]))

for(i in 4:6){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[7,2])
}

Summary[10,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'A'), "Amount"]))
Summary[11,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'B'), "Amount"]))
Summary[12,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'C'), "Amount"]))
Summary[13,2] <- sum(as.numeric(Summary[10:12,2]))

for(i in 10:12){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[13,2])
}

Summary[16,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'A'), "Amount"]))
Summary[17,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'B'), "Amount"])) 
Summary[18,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'C'), "Amount"]))
Summary[19,2] <- sum(as.numeric(Summary[16:18,2]))

for(i in 16:18){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[19,2])
}

Summary[22,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'A'), "Amount"])) 
Summary[23,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'B'), "Amount"])) 
Summary[24,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'C'), "Amount"]))
Summary[25,2] <- sum(as.numeric(Summary[22:24,2]))

for(i in 22:24){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[25,2])
}

Summary[28,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'A'), "Amount"]))
Summary[29,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'B'), "Amount"]))
Summary[30,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'C'), "Amount"]))
Summary[31,2] <- sum(as.numeric(Summary[28:30,2]))

for(i in 28:30){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[31,2])
}

Summary[34,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'A'), "Amount"])) 
Summary[35,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'B'), "Amount"])) 
Summary[36,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'C'), "Amount"]))
Summary[37,2] <- sum(as.numeric(Summary[34:36,2]))

for(i in 34:36){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[37,2])
}

Summary[40,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Others' & Combined$Class == 'A'), "Amount"])) 
Summary[41,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Others' & Combined$Class == 'B'), "Amount"])) 
Summary[42,2] <- sum(as.numeric(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Others' & Combined$Class == 'C'), "Amount"]))
Summary[43,2] <- sum(as.numeric(Summary[40:42,2]))

for(i in 40:42){
  Summary[i,3] <- as.numeric(Summary[i,2]) / as.numeric(Summary[43,2])
}

## Forecast $

Summary[4,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$Class == 'A'), "Amount"]))
Summary[5,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$Class == 'B'), "Amount"]))
Summary[6,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$Class == 'C'), "Amount"]))
Summary[7,4] <- sum(as.numeric(Summary[4:6,4]))

Summary[10,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'A'), "Amount"]))
Summary[11,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'B'), "Amount"]))
Summary[12,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'C'), "Amount"]))
Summary[13,4] <- sum(as.numeric(Summary[10:12,4]))

Summary[16,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'A'), "Amount"]))
Summary[17,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'B'), "Amount"])) 
Summary[18,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'C'), "Amount"]))
Summary[19,4] <- sum(as.numeric(Summary[16:18,4]))

Summary[22,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'A'), "Amount"])) 
Summary[23,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'B'), "Amount"])) 
Summary[24,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'C'), "Amount"]))
Summary[25,4] <- sum(as.numeric(Summary[22:24,4]))

Summary[28,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'A'), "Amount"]))
Summary[29,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'B'), "Amount"]))
Summary[30,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'C'), "Amount"]))
Summary[31,4] <- sum(as.numeric(Summary[28:30,4]))

Summary[34,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'A'), "Amount"])) 
Summary[35,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'B'), "Amount"])) 
Summary[36,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'C'), "Amount"]))
Summary[37,4] <- sum(as.numeric(Summary[34:36,4]))

Summary[40,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Others' & Combined$Class == 'A'), "Amount"])) 
Summary[41,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Others' & Combined$Class == 'B'), "Amount"])) 
Summary[42,4] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$SISR.Category == 'Others' & Combined$Class == 'C'), "Amount"]))
Summary[43,4] <- sum(as.numeric(Summary[40:42,4]))

for(i in c(4:7,10:13,16:19,22:25,28:31,34:37,40:43)){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
}

for(i in c(4:7,10:13,16:19,22:25,28:31,34:37,40:43)){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

## SKU
Summary[4,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$Class == 'A'), "Amount"])
Summary[5,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$Class == 'B'), "Amount"])
Summary[6,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$Class == 'C'), "Amount"])
Summary[7,8] <- sum(as.numeric(Summary[4:6,8]))

for(i in 4:6){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[7,8])
}

Summary[10,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'A'), "Amount"])
Summary[11,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'B'), "Amount"])
Summary[12,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Foam Mattress' & Combined$Class == 'C'), "Amount"])
Summary[13,8] <- sum(as.numeric(Summary[10:12,8]))

for(i in 10:12){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[13,8])
}

Summary[16,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'A'), "Amount"])
Summary[17,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'B'), "Amount"])
Summary[18,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Spring Mattress' & Combined$Class == 'C'), "Amount"])
Summary[19,8] <- sum(as.numeric(Summary[16:18,8]))

for(i in 16:18){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[19,8])
}

Summary[22,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'A'), "Amount"]) 
Summary[23,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'B'), "Amount"])
Summary[24,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Frame' & Combined$Class == 'C'), "Amount"])
Summary[25,8] <- sum(as.numeric(Summary[22:24,8]))

for(i in 22:24){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[25,8])
}

Summary[28,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'A'), "Amount"])
Summary[29,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'B'), "Amount"])
Summary[30,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Box Spring' & Combined$Class == 'C'), "Amount"])
Summary[31,8] <- sum(as.numeric(Summary[28:30,8]))

for(i in 28:30){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[31,8])
}

Summary[34,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'A'), "Amount"])
Summary[35,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'B'), "Amount"])
Summary[36,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Platform Bed' & Combined$Class == 'C'), "Amount"])
Summary[37,8] <- sum(as.numeric(Summary[34:36,8]))

for(i in 34:36){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[37,8])
}

Summary[40,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Others' & Combined$Class == 'A'), "Amount"])
Summary[41,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Others' & Combined$Class == 'B'), "Amount"]) 
Summary[42,8] <- length(Combined[which(Combined$List == 'Actual $' & Combined$SISR.Category == 'Others' & Combined$Class == 'C'), "Amount"])
Summary[43,8] <- sum(as.numeric(Summary[40:42,8]))

for(i in 40:42){
  Summary[i,9] <- as.numeric(Summary[i,8]) / as.numeric(Summary[43,8])
}

################ 2. Top 10 Summary ################
# Top 10 hot seller
Actual <- Combined[which(Combined$List == "Actual $"),]
Actual_QTY <-Combined[which(Combined$List=="Actual #"),]
Actual <- merge(Actual_QTY[,c("Amazon.SKU","Class","Amount")],Actual[,c("Amazon.SKU","Amount")],by.x ="Amazon.SKU",by.y="Amazon.SKU",all.y=TRUE)
colnames(Actual)[3] <-"Units"
colnames(Actual)[4] <-"Amount"
Actual <- Actual[order(-as.numeric(Actual[,"Amount"])),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)

Top10 <- c(Actual[1:10,"Amazon.SKU"])
Summary[46:55,1] <- Top10

# Top10 hot seller
Summary[46:55,2] <- Actual[1:10,"Amount"]
Summary[56,2] <- sum(as.numeric(Summary[46:55,2]))

for(i in 46:56){
  Summary[i,3] <- as.numeric(Summary[i,2]) / sum(as.numeric(Actual[1:nrow(Actual),"Amount"]))
}


######## 1-2) Forecast $ ########
for(i in 46:55){
  Summary[i,4] <- Combined[which(Combined$List == "Forecast $" & Combined$Amazon.SKU == Top10[i-45]), "Amount"]
}
Summary[56,4] <- sum(as.numeric(Summary[46:55,4]))

for(i in 46:55){
  Summary[i,5] <- (as.numeric(Summary[i,4]) - as.numeric(Summary[i,2])) / as.numeric(Summary[i,4])
  if(!is.finite(Summary[i,5])){
    Summary[i,5] <- NaN
  }
}

######## 1-3) Deviation $ ########
for(i in 46:55){
  Summary[i,6] <- as.numeric(Summary[i,2]) - as.numeric(Summary[i,4])
}

######## 2-1) Actual Sales Units ######## 
Summary[46:55,8] <- Actual[1:10,"Units"]
Summary[56,8] <- sum(as.numeric(Summary[46:55,8]))

for(i in 46:56){
  Summary[i,9] <- as.numeric(Summary[i,8]) / sum(as.numeric(Actual[1:nrow(Actual),"Units"]))
}

######## 2-2) Forecast Units ########
for(i in 46:55){
  Summary[i,10] <- Combined[which(Combined$List == "Forecast #" & Combined$Amazon.SKU == Top10[i-45]), "Amount"]
}
Summary[56,10] <- sum(as.numeric(Summary[46:55,10]))

for(i in 46:55){
  Summary[i,11] <- (as.numeric(Summary[i,10]) - as.numeric(Summary[i,8])) / as.numeric(Summary[i,10])
  if(!is.finite(Summary[i,11])){
    Summary[i,11] <- NaN
  }
}

######## 2-3) Deviation Units ########
for(i in 46:56){
  Summary[i,12] <- as.numeric(Summary[i,8]) - as.numeric(Summary[i,10])
}

# Top 10 inaccurate forecast SKUs by Class
##Set Actual Comparison for all classes
Actual_Comparison <- Actual
colnames(Actual_Comparison)[3] <- "ActualUnits"
colnames(Actual_Comparison)[4] <- "ActualAmount"
########## A item  #####################
Forecast_A_Amount <- Combined[which(Combined$List == "Forecast $" & Combined$Class == "A"),]
Forecast_A_Units <- Combined[which(Combined$List == "Forecast #" & Combined$Class == "A"),]
Forecast_A <- merge(Forecast_A_Units[,c("Class","Amazon.SKU","Amount")],Forecast_A_Amount[,c("Amazon.SKU","Amount")], by="Amazon.SKU", all.x=TRUE)
colnames(Forecast_A)[3] <-"ForecastUnits"
colnames(Forecast_A)[4] <-"ForecastAmount"
Forecast_A <- data.frame(Forecast_A)
rownames(Forecast_A) <- 1:nrow(Forecast_A)

Comparison_A <- merge(Actual_Comparison,Forecast_A[,-c(2)], by.x ="Amazon.SKU",by.y="Amazon.SKU", all.y = TRUE)

for(i in 1:nrow(Comparison_A)){
  Comparison_A$Rate[i] = 0
  Comparison_A$Rate[i] <-abs(as.numeric(Comparison_A[i,"ForecastUnits"])-as.numeric(Comparison_A[i,"ActualUnits"]))/as.numeric(Comparison_A[i,"ForecastUnits"])
}

for(i in 1:ncol(Comparison_A)){
  Comparison_A[is.na(Comparison_A[,i]),i]<-0
}

Comparison_A <-Comparison_A[order(-as.numeric(Comparison_A[,"Rate"])),]
Comparison_A <- data.frame(Comparison_A)
rownames(Comparison_A) <- 1:nrow(Comparison_A)

A10 <- c(Comparison_A[1:10,"Amazon.SKU"])

################### B item ##############F
Forecast_B_Amount <- Combined[which(Combined$List == "Forecast $" & Combined$Class == "B"),]
Forecast_B_Units <- Combined[which(Combined$List == "Forecast #" & Combined$Class == "B"),]
Forecast_B <- merge(Forecast_B_Units[,c("Class","Amazon.SKU","Amount")],Forecast_B_Amount[,c("Amazon.SKU","Amount")], by="Amazon.SKU", all.x=TRUE)
colnames(Forecast_B)[3] <-"ForecastUnits"
colnames(Forecast_B)[4] <-"ForecastAmount"
Forecast_B <- data.frame(Forecast_B)
rownames(Forecast_B) <- 1:nrow(Forecast_B)

Comparison_B <- merge(Actual_Comparison,Forecast_B[,-c(2)], by.x ="Amazon.SKU",by.y="Amazon.SKU", all.y = TRUE)

for(i in 1:nrow(Comparison_B)){
  Comparison_B$Rate[i] = 0
  Comparison_B$Rate[i] <-abs(as.numeric(Comparison_B[i,"ForecastUnits"])-as.numeric(Comparison_B[i,"ActualUnits"]))/as.numeric(Comparison_B[i,"ForecastUnits"])
}

for(i in 1:ncol(Comparison_B)){
  Comparison_B[is.na(Comparison_B[,i]),i]<-0
}

Comparison_B <-Comparison_B[order(-as.numeric(Comparison_B[,"Rate"])),]
Comparison_B <- data.frame(Comparison_B)
rownames(Comparison_B) <- 1:nrow(Comparison_B)

B10 <- c(Comparison_B[1:10,"Amazon.SKU"])


############ C item ###################
Forecast_C_Amount <- Combined[which(Combined$List == "Forecast $" & Combined$Class == "C"),]
Forecast_C_Units <- Combined[which(Combined$List == "Forecast #" & Combined$Class == "C"),]
Forecast_C <- merge(Forecast_C_Units[,c("Class","Amazon.SKU","Amount")],Forecast_C_Amount[,c("Amazon.SKU","Amount")], by="Amazon.SKU", all.x=TRUE)
colnames(Forecast_C)[3] <-"ForecastUnits"
colnames(Forecast_C)[4] <-"ForecastAmount"
Forecast_C <- data.frame(Forecast_C)
rownames(Forecast_C) <- 1:nrow(Forecast_C)

Comparison_C <- merge(Actual_Comparison,Forecast_C[,-c(2)], by.x ="Amazon.SKU",by.y="Amazon.SKU", all.y = TRUE)

for(i in 1:nrow(Comparison_C)){
  Comparison_C$Rate[i] = 0
  Comparison_C$Rate[i] <-abs(as.numeric(Comparison_C[i,"ForecastUnits"])-as.numeric(Comparison_C[i,"ActualUnits"]))/as.numeric(Comparison_C[i,"ForecastUnits"])
}

for(i in 1:ncol(Comparison_C)){
  Comparison_C[is.na(Comparison_C[,i]),i]<-0
}

Comparison_C <-Comparison_C[order(-as.numeric(Comparison_C[,"Rate"])),]
Comparison_C <- data.frame(Comparison_C)
rownames(Comparison_C) <- 1:nrow(Comparison_C)

C10 <- c(Comparison_C[1:10,"Amazon.SKU"])



# Top20 inaccurate SKUs
# this part is for inaccurate ABC SKUs only
Summary[59:68,1] <- A10
Summary[71:80,1] <- B10
Summary[83:92,1] <- C10


Summary[59:68,2] <- Comparison_A[1:10,"ActualAmount"]
Summary[59:68,4] <- Comparison_A[1:10,"ForecastAmount"]
Summary[59:68,6] <-as.numeric(Summary[59:68,2])-as.numeric(Summary[59:68,4])
Summary[59:68,7] <- Comparison_A[1:10,"Rate"]
Summary[59:68,8] <- Comparison_A[1:10,"ActualUnits"]
Summary[59:68,10] <- Comparison_A[1:10,"ForecastUnits"]
Summary[59:68,12] <- as.numeric(Summary[59:68,8])-as.numeric(Summary[59:68,10])


Summary[71:80,2] <- Comparison_B[1:10,"ActualAmount"]
Summary[71:80,4] <- Comparison_B[1:10,"ForecastAmount"]
Summary[71:80,6] <-as.numeric(Summary[71:80,2])-as.numeric(Summary[71:80,4])
Summary[71:80,7] <- Comparison_B[1:10,"Rate"]
Summary[71:80,8] <- Comparison_B[1:10,"ActualUnits"]
Summary[71:80,10] <- Comparison_B[1:10,"ForecastUnits"]
Summary[71:80,12] <- as.numeric(Summary[71:80,8])-as.numeric(Summary[71:80,10])


Summary[83:92,2] <- Comparison_C[1:10,"ActualAmount"]
Summary[83:92,4] <- Comparison_C[1:10,"ForecastAmount"]
Summary[83:92,6] <-as.numeric(Summary[83:92,2])-as.numeric(Summary[83:92,4])
Summary[83:92,7] <- Comparison_C[1:10,"Rate"]
Summary[83:92,8] <- Comparison_C[1:10,"ActualUnits"]
Summary[83:92,10] <- Comparison_C[1:10,"ForecastUnits"]
Summary[83:92,12] <- as.numeric(Summary[83:92,8])-as.numeric(Summary[83:92,10])



############### FINISH COPY  ################################################


setwd("//192.168.1.41/03.Demand Planning/Elvis/Forecast Accuracy Raw Data")
write.csv(Summary, 'Amazon.csv')
