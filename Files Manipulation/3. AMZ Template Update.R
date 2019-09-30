
library(xlsx)
library(openxlsx)
options(scipen = 999)

setwd("//192.168.1.116/Shared/Accpacadv/Young")
Format <- read.xlsx("AMZ_SKU.xlsx", sheet = 'Sheet1')
Format <- Format[,c(2,3,5)]

## From Current Week AMZ Inv
AMZ_FC_Needed <- AMZ_FC_2017[,c(1,3:ncol(AMZ_FC_2017))]

## From Current Week ZIN Inv
ZIN_Inv_Needed <- Zinus_Inv_2017[,c(1,3:ncol(Zinus_Inv_2017))] 

## Combine Data
Format <- merge(Format, AMZ_FC_Needed, by.x = 'Zinus.SKU', by.y = 'Zinus.SKU#', all.x = TRUE)
Format <- merge(Format, ZIN_Inv_Needed, by.x = 'Zinus.SKU', by.y = 'Zinus.SKU#', all.x = TRUE)

Format[,4:ncol(Format)][is.na(Format[,4:ncol(Format)])] <- 0
Format[,4:ncol(Format)] <- sapply(Format[,4:ncol(Format)], as.numeric)

## July (~7/22/2018)
Final_Inv <- Format[,1:55]
Final_Inv[,4:55] <- Format[,4:55] + Format[,56:ncol(Format)]

setwd("//192.168.1.116/Shared/Accpacadv/Young/0) Monday/Inventory update for AMZ Template & AM SISRs")
write.csv(Final_Inv, "Final Inv for AMZ Template.csv")
