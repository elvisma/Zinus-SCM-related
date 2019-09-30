
library(xlsx)
library(openxlsx)
options(scipen = 999)

## Read Inventory
# Box Spring
BX_Inv <- BX_AISR[which(BX_AISR$Account == '(SCM) Fulfilment Center'), c(7,13,40:91)]

# Frame
FR_Inv <- FR_AISR[which(FR_AISR$Account == '(SCM) Fulfilment Center'), c(7,13,40:91)]

# Foam Mattress
FM_Inv <- FM_AISR[which(FM_AISR$Account == '(SCM) Fulfilment Center'), c(7,13,40:91)]

# Spring Mattress
SM_Inv <- SM_AISR[which(SM_AISR$Account == '(SCM) Fulfilment Center'), c(7,13,40:91)]

# Platform Bed
PB_Inv <- PB_AISR[which(PB_AISR$Account == '(SCM) Fulfilment Center'), c(7,13,40:91)]

# Others
OT_Inv <- OT_AISR[which(OT_AISR$Account == '(SCM) Fulfilment Center'), c(7,13,40:91)]

# Furniture
FN_Inv <- FN_AISR[which(FN_AISR$Account == '(SCM) Fulfilment Center'), c(7,13,40:91)]

AMZ_FC <- rbind(BX_Inv, FR_Inv)
AMZ_FC <- rbind(AMZ_FC, FM_Inv)
AMZ_FC <- rbind(AMZ_FC, SM_Inv)
AMZ_FC <- rbind(AMZ_FC, PB_Inv)
AMZ_FC <- rbind(AMZ_FC, OT_Inv)
AMZ_FC <- rbind(AMZ_FC, FN_Inv)

AMZ_FC_2017 <- AMZ_FC
