
library(xlsx)
library(openxlsx)
options(scipen = 999)

### 12.26.2017
SISR_Folder_Date = "06.09.2019"
SISR_Date = "20190609"

SISR_Folder = paste("//192.168.1.41/03.Demand Planning/SISR ", SISR_Folder_Date, sep="")

#########################################################################################
## Read SISR
# 1. Frame
setwd(paste(SISR_Folder,"/Box Spring", sep = ""))

BX_SISR_Name <- paste('SISR_BX_',SISR_Date, sep = "")
SB_SISR_Name <- paste('SISR_SB_',SISR_Date, sep = "")

BX_SISR <- read.xlsx(paste(BX_SISR_Name, '.xlsm', sep = ""), sheet = 'BX', startRow = 5)
SB_SISR <- read.xlsx(paste(SB_SISR_Name, '.xlsm', sep = ""), sheet = 'SB', startRow = 5)

BX_SISR$SISR.Category <- "BX"
SB_SISR$SISR.Category <- "SB"

# 2. Mattress
setwd(paste(SISR_Folder,"/Mattress", sep=""))

FM_SISR_Name <- paste('SISR_FM_',SISR_Date, sep = "")
SM_SISR_Name <- paste('SISR_SM_',SISR_Date, sep = "")

FM_SISR <- read.xlsx(paste(FM_SISR_Name, '.xlsm', sep = ""), sheet = 'FM', startRow = 5)
SM_SISR <- read.xlsx(paste(SM_SISR_Name, '.xlsm', sep = ""), sheet = 'SM', startRow = 5)

FM_SISR$SISR.Category <- "FM"
SM_SISR$SISR.Category <- "SM"

# 3. Platform Bed
setwd(paste(SISR_Folder,"/Platform Bed", sep=""))

PB_SISR_Name <- paste('SISR_PB_',SISR_Date, sep = "")
OT_SISR_Name <- paste('SISR_OT_',SISR_Date, sep = "")
FN_SISR_Name <- paste('SISR_FN_',SISR_Date, sep = "")
FR_SISR_Name <- paste('SISR_FR_',SISR_Date, sep = "")

PB_SISR <- read.xlsx(paste(PB_SISR_Name, '.xlsm', sep = ""), sheet = 'PB', startRow = 5)
OT_SISR <- read.xlsx(paste(OT_SISR_Name, '.xlsm', sep = ""), sheet = 'OT', startRow = 5)
FN_SISR <- read.xlsx(paste(FN_SISR_Name, '.xlsm', sep = ""), sheet = 'FN', startRow = 5)
FR_SISR <- read.xlsx(paste(FR_SISR_Name, '.xlsm', sep = ""), sheet = 'FR', startRow = 5)

PB_SISR$SISR.Category <- "PB"
OT_SISR$SISR.Category <- "OT"
FN_SISR$SISR.Category <- "FN"
FR_SISR$SISR.Category <- "FR"
#########################################################################################


#####################################################################################
################Starting from here just added for AMZ tracker update#################
##change headers' names###
#colnames(BX_SISR) <- colnames(FM_SISR)
#colnames(FR_SISR) <- colnames(FM_SISR)
# Combine All###
SISRs <- rbind(FM_SISR[,c(ncol(FM_SISR),2:64)], SM_SISR[,c(ncol(SM_SISR),2:64)])
SISRs <- rbind(SISRs, FR_SISR[,c(ncol(FR_SISR),2:64)])
SISRs <- rbind(SISRs, BX_SISR[,c(ncol(BX_SISR),2:64)])
SISRs <- rbind(SISRs, PB_SISR[,c(ncol(PB_SISR),2:64)])
SISRs <- rbind(SISRs, OT_SISR[,c(ncol(OT_SISR),2:64)])
SISRs <- rbind(SISRs, FN_SISR[,c(ncol(FN_SISR),2:64)])
SISRs <- rbind(SISRs, SB_SISR[,c(ncol(SB_SISR),2:64)])
SISRs <- SISRs[which(SISRs$'Zinus.SKU#' != 'NA'),]
SISRs <- SISRs[,c(1,7,12,13:ncol(SISRs))]
SISRs <- SISRs[which(SISRs$Account != 'NA'),]

SISR_Inv <- SISRs[which(SISRs$Account == 'Total Ending Inventory'),]

