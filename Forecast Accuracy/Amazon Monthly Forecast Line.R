
## AMZ Monthly Forecast Line

library(xlsx)
library(openxlsx)
options(scipen = 999)


setwd("//192.168.1.41/09.Sales/David L/Amazon Demand Forecast")

Combined <- read.xlsx('AMZ Demand Forecast_All_06172019.xlsm', sheet = 'Master', startRow=3)

Col1=36  # whole June2019 
Col2=39

Combined <- Combined[,c(1,4:6,11,13,Col1:Col2)]



for(i in 1:ncol(Combined)){
  Combined[is.na(Combined[,i]),i]<-0
}




Combined <- Combined[,-c(2,3)]
for (i in 1:nrow(Combined)){
  
  Combined$Amount[i] = 0
  Combined$Amount[i]<-sum(as.numeric(Combined[i,c(5:ncol(Combined))]))
}


Combined_Forecast_Line <- data.frame(matrix(nrow=5,ncol=2))
Combined_Forecast_Line[c(2,3,4,5),1] <- c("A","B","C","Total")
Combined_Forecast_Line[1,2] <-"Monthly Forecast Line"


Combined_Forecast_Line[2,2] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$Class == 'A'), "Amount"]))
Combined_Forecast_Line[3,2] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$Class == 'B'), "Amount"]))
Combined_Forecast_Line[4,2] <- sum(as.numeric(Combined[which(Combined$List == 'Forecast $' & Combined$Class == 'C'), "Amount"]))
Combined_Forecast_Line[5,2] <- sum(as.numeric(Combined_Forecast_Line[2:4,2]))

# copy the data frame Combined_Forecast_Line every month
