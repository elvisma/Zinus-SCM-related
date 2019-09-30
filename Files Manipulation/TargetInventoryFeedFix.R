
library(xlsx)
library(openxlsx)
library(dplyr)
library('scales')
library(purrr)
options(scipen = 999)

setwd("C:/Users/Elvis Ma/Desktop")
#### Target_tracy (846 file)
#### Target_CHS (846 file)
### TR (excel)
### CHS (excel)
Target_tracy <- read.delim("target_tr.txt",header=FALSE, sep="*",stringsAsFactors = FALSE)
for (i in seq_along(Target_tracy)){
  Target_tracy[which(is.na(Target_tracy[,i])),i] <-''
}
write.csv(Target_tracy,"TR.csv",row.names = FALSE)



Target_CHS <- read.delim("target_chs.txt",header=FALSE, sep="*",stringsAsFactors = FALSE)
for (i in seq_along(Target_CHS)){
  Target_CHS[which(is.na(Target_CHS[,i])),i] <-''
}
write.csv(Target_CHS,"CHS.csv",row.names = FALSE)


#####################################################

TR <-read.csv("TR.csv")
CHS <-read.csv("CHS.csv")
for (i in seq_along(TR)){
  TR[which(is.na(TR[,i])),i] <-''
}


for (i in seq_along(CHS)){
  CHS[which(is.na(CHS[,i])),i] <-''
}

write.table(TR,"TR.txt", sep = "*", col.names =FALSE, row.names=FALSE, quote = FALSE)
write.table(CHS,"CHS.txt", sep = "*", col.names =FALSE, row.names=FALSE, quote = FALSE)


