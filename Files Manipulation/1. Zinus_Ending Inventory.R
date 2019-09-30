
library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read Inventory
# Box Spring
BX_Inv <- BX_SISR[which(BX_SISR$Account == 'Total Ending Inventory'), c(7,12,39:90)]

# Frame
FR_Inv <- FR_SISR[which(FR_SISR$Account == 'Total Ending Inventory'), c(7,12,39:90)]

# Foam Mattress
FM_Inv <- FM_SISR[which(FM_SISR$Account == 'Total Ending Inventory'), c(7,12,39:90)]

# Spring Mattress
SM_Inv <- SM_SISR[which(SM_SISR$Account == 'Total Ending Inventory'), c(7,12,39:90)]

# Platform Bed
PB_Inv <- PB_SISR[which(PB_SISR$Account == 'Total Ending Inventory'), c(7,12,39:90)]

# Others
OT_Inv <- OT_SISR[which(OT_SISR$Account == 'Total Ending Inventory'), c(7,12,39:90)]

# Furniture
FN_Inv <- FN_SISR[which(FN_SISR$Account == 'Total Ending Inventory'), c(7,12,39:90)]

## Total
Zinus_Inv <- rbind(BX_Inv, FR_Inv)
Zinus_Inv <- rbind(Zinus_Inv, FM_Inv)
Zinus_Inv <- rbind(Zinus_Inv, SM_Inv)
Zinus_Inv <- rbind(Zinus_Inv, PB_Inv)
Zinus_Inv <- rbind(Zinus_Inv, OT_Inv)
Zinus_Inv <- rbind(Zinus_Inv, FN_Inv)

Zinus_Inv_2017 <- Zinus_Inv

setwd("//192.168.1.116/Shared/Accpacadv/Young/0) Monday/Inventory update for AMZ Template & AM SISRs")
write.csv(Zinus_Inv_2017, "Zinus_Inv.csv")
