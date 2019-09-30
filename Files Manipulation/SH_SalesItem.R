library(xlsx)
library(openxlsx)
library(dplyr)
library(lubridate)
options(scipen = 999)

start_week <-"12/30/2018"
end_week <-"01/20/2019"
date_range <- seq(as.Date(start_week,format="%m/%d/%Y"),as.Date(end_week,format="%m/%d/%Y"),"week")

# subject to change later
setwd("C:/Users/Elvis Ma/Desktop/Little Projects")
# make sure the header is text format
SH_SalesItem <-xlsx::read.xlsx("SH_SalesItem.xlsx",sheetName ="SH_SalesItem")
col_names <- colnames(SH_SalesItem)[4:ncol(SH_SalesItem)]
trimmed_names <- gsub(pattern = "X",replacement = "",col_names)
trimmed_names <- gsub(pattern = "\\.",replacement = "/",trimmed_names)
names(SH_SalesItem)[4:ncol(SH_SalesItem)] <-trimmed_names

for (i in 4:ncol(SH_SalesItem)){
  SH_SalesItem[is.na(SH_SalesItem[,i]),i] <- 0
} 

##### to fix ####
#date_names <- names(SH_SalesItem)[4:ncol(SH_SalesItem)]
#date_names <-mdy(date_names)
#colnames(SH_SalesItem)[4:ncol(SH_SalesItem)] <-date_names

#### Wayfair SH combine Castlegate US & Wayfair.com
WFDSV_SH <-SH_SalesItem[which(SH_SalesItem$Customer.name=="Wayfair.com"),-c(1,3)]
CGUS_SH <-SH_SalesItem[which(SH_SalesItem$Customer.name=="Castlegate US"),-c(1,3)]
rownames(WFDSV_SH) <-1:nrow(WFDSV_SH)
rownames(CGUS_SH) <- 1:nrow(CGUS_SH)

USWF_SH <- merge(WFDSV_SH,CGUS_SH,by = "ZINUS.SKU",all = TRUE)
for (i in 2:ncol(USWF_SH)){
  USWF_SH[is.na(USWF_SH[,i]),i] <- 0
} 

USWF_SH[,2:ncol(USWF_SH)] <- sapply(USWF_SH[,2:ncol(USWF_SH)],as.numeric)

Wayfair_SH <-USWF_SH[,1:ncol(WFDSV_SH)]
names(Wayfair_SH) <-gsub(pattern="\\.x",replacement="",names(Wayfair_SH))
Wayfair_SH[,2:ncol(Wayfair_SH)] <- USWF_SH[,2:ncol(Wayfair_SH)] + USWF_SH[,(ncol(Wayfair_SH)+1):ncol(USWF_SH)]
############## Finish line for Wayfair #############
###################### TESTING CLEANING #############
library(tidyr)
Wayfair_gather <- gather(Wayfair_SH,Date,`Actual QTY`,-`ZINUS.SKU`)
Wayfair_gather$Date <- as.Date(Wayfair_gather$Date,format = "%m/%d/%Y")
##get the week number
for (i in 1:nrow(Wayfair_gather)){
  if (isoweek(Wayfair_gather$Date[i])==52){
    Wayfair_gather$Week[i] = 1
  }
  else
    {Wayfair_gather$Week[i] <-isoweek(Wayfair_gather$Date[i])+1}
}

Wayfair_update <-Wayfair_gather[which(Wayfair_gather$Date%in%date_range),]

Wayfair_group <-group_by(Wayfair_gather,ZINUS.SKU)
Wayfair_summ <-summarize(Wayfair_group,`Actual QTY`=sum(`Actual QTY`))
###################### TESTING CLEANING 

