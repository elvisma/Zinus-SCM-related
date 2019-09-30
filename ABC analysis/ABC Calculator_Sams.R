
## Sam's Club Demand Forecast Summary Report

library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read and prepare AMZ Price Data
setwd("C:/Users/Elvis Ma/Desktop")

Date = "01012019"
Col = 61

File_Name <- paste('samsclub_', Date, sep = "")
SISR <- openxlsx::read.xlsx(paste(File_Name, '.xlsm', sep = ""), sheet = 'Master', startRow = 3, cols=c(2,3,6,10,(Col-12):Col))

for(i in 5:ncol(SISR)){
  SISR[is.na(SISR[,i]),i] <- 0
}

Actual <- SISR[which(SISR$Account == "Retail Sales"),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)

for(i in 1:nrow(Actual)){
  Actual$Actual.Sum[i] <- sum(as.numeric(Actual[i,c(5:ncol(Actual))]))
}

for(i in 1:nrow(Actual)){
  Actual[i,18] <- sum(as.numeric(Actual[i,5:17]))
}

Actual <- Actual[order(-as.numeric(Actual[,18])),]
Actual <- data.frame(Actual)
rownames(Actual) <- 1:nrow(Actual)

for(i in 1:nrow(Actual)){
  Actual$Percentage[i] <- Actual[i,18] / sum(Actual[,18])
}

Actual$Cumulative[1] <- Actual[1,19]

for(i in 2:nrow(Actual)){
  Actual$Cumulative[i] <- Actual[i-1,20] + Actual[i,19]
}

for(i in 1:nrow(Actual)){
  if(Actual[i,20] >= 0 & Actual[i,20] <= 0.8) {
    Actual$ABC[i] <- 'A'
  }
  else if(Actual[i,20] > 0.8 & Actual[i,20] <= 0.95) {
    Actual$ABC[i] <- 'B'
  }
  else {
    Actual$ABC[i] <- 'C'
  }
}

setwd("//192.168.1.116/Shared/Accpacadv/Young/1) Tuesday")
write.csv(Actual[,c(3,18,20,21)], 'Sams_ABC.csv')
