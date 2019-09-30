library(xlsx)
library(openxlsx)
library(dplyr)
options(scipen=999)

setwd("C:/Users/Elvis Ma/Dropbox (Zinus Inc.)/WAYFAIR WKLY")

US_Date ="09092019"
UK_Date ="09092019"
US_Col = 96 ### 9/1/2019
UK_Col = 147 ### 9/1/2019

US_WF_Name <- paste("USWayfair_",US_Date,sep = "")
UK_WF_Name <- paste("UKWayfair_",UK_Date,sep = "")


US_WF <-read.xlsx(paste(US_WF_Name,".xlsm",sep = ""),sheet = "Demand",startRow = 2, cols = c(2,3,6,7,8,(US_Col-12):US_Col))
UK_WF <-read.xlsx(paste(UK_WF_Name,".xlsm",sep = ""),sheet = "Demand",startRow = 2, cols = c(2,3,5,6,7,(UK_Col-12):UK_Col))

for(i in 6:ncol(US_WF)){
  US_WF[is.na(US_WF[,i]),i] <-0  
}

for(i in 6:ncol(UK_WF)){
  UK_WF[is.na(UK_WF[,i]),i] <-0  
}

US_Actual <- filter(US_WF,Account=="Actual Sales Total")
UK_Actual <- filter(UK_WF,Account=="Actual Sales Total")

for(i in 1:nrow(US_Actual)){
  US_Actual$Sum[i]=0
  US_Actual$Sum[i] <-sum(as.numeric(US_Actual[i,c(6:ncol(US_Actual))]))
}

for(i in 1:nrow(UK_Actual)){
  UK_Actual$Sum[i]=0
  UK_Actual$Sum[i] <-sum(as.numeric(UK_Actual[i,c(6:ncol(UK_Actual))]))
}

US_Actual <-US_Actual[order(-as.numeric(US_Actual[,ncol(US_Actual)])),]
UK_Actual <-UK_Actual[order(-as.numeric(UK_Actual[,ncol(UK_Actual)])),]

US_Actual$Percentage <-US_Actual$Sum/sum(US_Actual$Sum)
UK_Actual$Percentage <-UK_Actual$Sum/sum(UK_Actual$Sum)

US_Actual$Cumulative[1]=US_Actual$Percentage[1]
for(i in 2:nrow(US_Actual)){
  US_Actual$Cumulative[i] <- US_Actual$Cumulative[i-1]+US_Actual$Percentage[i]
}
for(i in 1:nrow(US_Actual)){
  if(US_Actual$Cumulative[i]>=0 & US_Actual$Cumulative[i]<=0.8)
  {US_Actual$ABC[i] <-"A"}
  else if(US_Actual$Cumulative[i]>0.8 & US_Actual$Cumulative[i]<=0.95)
  {US_Actual$ABC[i] <-"B"}
  else
  {US_Actual$ABC[i] <-"C"}
}

UK_Actual$Cumulative[1]=UK_Actual$Percentage[1]
for(i in 2:nrow(UK_Actual)){
  UK_Actual$Cumulative[i] <- UK_Actual$Cumulative[i-1]+UK_Actual$Percentage[i]
}
for(i in 1:nrow(UK_Actual)){
  if(UK_Actual$Cumulative[i]>=0 & UK_Actual$Cumulative[i]<=0.8)
  {UK_Actual$ABC[i] <-"A"}
  else if(UK_Actual$Cumulative[i]>0.8 & UK_Actual$Cumulative[i]<=0.95)
  {UK_Actual$ABC[i] <-"B"}
  else
  {UK_Actual$ABC[i] <-"C"}
}

xlsx::write.xlsx(US_Actual[,c(3,20,21,22)], 'ABC.xlsx',sheetName="US CG and .com",row.names=FALSE)
xlsx::write.xlsx(UK_Actual[,c(3,20,21,22)], 'ABC.xlsx',sheetName="UK CG", row.names=FALSE, append=TRUE)
