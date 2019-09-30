
## AMZ Demand Forecast Summary Report

library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read and prepare AMZ Price Data
setwd("C:/Users/Elvis Ma/Dropbox (Zinus Inc.)/COSTCO.COM WKLY")

Date = "01282019"
Col = 168

File_Name <- paste('costcocom_', Date, sep = "")
SISR <- read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 3, cols=c(1,2,3,7,8,(Col-12):Col))
SISR[is.na(SISR[,4]),4] <- 0

for(i in 6:ncol(SISR)){
  SISR[is.na(SISR[,i]),i] <- 0
}

Actual <- SISR[which(SISR$Account == "Actual"),]

for(i in 1:nrow(Actual)){
  Actual$Actual.Sum[i] <- sum(as.numeric(Actual[i,c(6:ncol(Actual))]))
}

for(i in 1:nrow(Actual)){
  Actual[i,19] <- sum(as.numeric(Actual[i,6:18]))
}

Actual$Amount <- Actual$Actual.Sum * Actual$Unit.Price

Actual <- Actual[order(-as.numeric(Actual[,20])),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)

for(i in 1:nrow(Actual)){
  Actual$Percentage[i] <- Actual[i,20] / sum(Actual[,20])
}

Actual$Cumulative[1] <- Actual[1,21]

for(i in 2:nrow(Actual)){
  Actual$Cumulative[i] <- Actual[i-1,22] + Actual[i,21]
}

for(i in 1:nrow(Actual)){
  if(Actual[i,22] >= 0 & Actual[i,22] <= 0.8) {
    Actual$ABC[i] <- 'A'
  }
  else if(Actual[i,22] > 0.8 & Actual[i,22] <= 0.95) {
    Actual$ABC[i] <- 'B'
  }
  else {
    Actual$ABC[i] <- 'C'
  }
}

setwd("//192.168.1.116/Shared/Accpacadv/Young/1) Tuesday")
write.csv(Actual[,c(3,20,22,2)], 'Costco_ABC.csv')
