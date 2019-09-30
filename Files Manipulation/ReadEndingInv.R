library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read SISR
# 1. Frame
startcol=35 ### 6/2/2019
endcol=startcol+25
setwd("C:/Users/Elvis Ma/Desktop/Weekly Work/space_planning")
SISR_Date = "AMZ"

SB_SISR_Name <- paste('SISR_SB_',SISR_Date, sep = "")
BX_SISR_Name <- paste('SISR_BX_',SISR_Date, sep = "")
FR_SISR_Name <- paste('SISR_FR_',SISR_Date, sep = "")
FM_SISR_Name <- paste('SISR_FM_',SISR_Date, sep = "")
SM_SISR_Name <- paste('SISR_SM_',SISR_Date, sep = "")
IM_SISR_Name <- paste('SISR_IM_',SISR_Date, sep = "")
PB_SISR_Name <- paste('SISR_PB_',SISR_Date, sep = "")
OT_SISR_Name <- paste('SISR_OT_',SISR_Date, sep = "")
FN_SISR_Name <- paste('SISR_FN_',SISR_Date, sep = "")

SB_SISR <- read.xlsx(paste(SB_SISR_Name, '.xlsm', sep = ""), sheet = 'SB', startRow = 5,cols = c(7,12,startcol:endcol))
BX_SISR <- read.xlsx(paste(BX_SISR_Name, '.xlsm', sep = ""), sheet = 'BX', startRow = 5,cols = c(7,12,startcol:endcol))
FR_SISR <- read.xlsx(paste(FR_SISR_Name, '.xlsm', sep = ""), sheet = 'FR', startRow = 5,cols = c(7,12,startcol:endcol))
FM_SISR <- read.xlsx(paste(FM_SISR_Name, '.xlsm', sep = ""), sheet = 'FM', startRow = 5,cols = c(7,12,startcol:endcol))
SM_SISR <- read.xlsx(paste(SM_SISR_Name, '.xlsm', sep = ""), sheet = 'SM', startRow = 5,cols = c(7,12,startcol:endcol))
IM_SISR <- read.xlsx(paste(IM_SISR_Name, '.xlsm', sep = ""), sheet = 'IM', startRow = 5,cols = c(7,12,startcol:endcol))
PB_SISR <- read.xlsx(paste(PB_SISR_Name, '.xlsm', sep = ""), sheet = 'PB', startRow = 5,cols = c(7,12,startcol:endcol))
OT_SISR <- read.xlsx(paste(OT_SISR_Name, '.xlsm', sep = ""), sheet = 'OT', startRow = 5,cols = c(7,12,startcol:endcol))
FN_SISR <- read.xlsx(paste(FN_SISR_Name, '.xlsm', sep = ""), sheet = 'FN', startRow = 5,cols = c(7,12,startcol:endcol))

SB_Inv <-SB_SISR[which(SB_SISR$Account=='Total Ending Inventory'),-c(2)]
BX_Inv <-BX_SISR[which(BX_SISR$Account=='Total Ending Inventory'),-c(2)]
FR_Inv <-FR_SISR[which(FR_SISR$Account=='Total Ending Inventory'),-c(2)]
FM_Inv <-FM_SISR[which(FM_SISR$Account=='Total Ending Inventory'),-c(2)]
SM_Inv <-SM_SISR[which(SM_SISR$Account=='Total Ending Inventory'),-c(2)]
IM_Inv <-IM_SISR[which(IM_SISR$Account=='Total Ending Inventory'),-c(2)]
PB_Inv <-PB_SISR[which(PB_SISR$Account=='Total Ending Inventory'),-c(2)]
OT_Inv <-OT_SISR[which(OT_SISR$Account=='Total Ending Inventory'),-c(2)]
FN_Inv <-FN_SISR[which(FN_SISR$Account=='Total Ending Inventory'),-c(2)]

Inv_list=list(SB_Inv,BX_Inv,FR_Inv,FM_Inv,SM_Inv,IM_Inv,PB_Inv,OT_Inv,FN_Inv)
Combine_Inv=do.call("rbind", Inv_list)
Combine_Inv[,c(2:ncol(Combine_Inv))]=sapply(Combine_Inv[,c(2:ncol(Combine_Inv))],as.numeric)

setwd("C:/Users/Elvis Ma/Desktop")
xlsx::write.xlsx(Combine_Inv,"Zinus_Inv.xlsx",row.names=FALSE)

