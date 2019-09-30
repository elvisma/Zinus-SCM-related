#install and load packages
#AND Install JAVA
pkg <- c("XLConnect")
new.pkg <- pkg[!(pkg %in% installed.packages())]

if (length(new.pkg)) {
  install.packages(new.pkg)
}
install.packages("XLConnect")
install.packages("devtools")
library(devtools)
install_version("XLConnectJars", version = "0.2-12", repos = "http://cran.us.r-project.org") 
install_version("XLConnect", version = "0.2-12", repos = "http://cran.us.r-project.org")
library(XLConnect)
install.packages('xlsx', lib='\\network\R\library')
library(xlsx, lib='\\network\R\library'))
install.packages('dplyr')
library(dplyr)
#load excel workbook and get sheet names
excel_1 <- loadWorkbook('ZINUS.xlsx')
sheet_names <- getSheets(excel_1)
names(sheet_names) <- sheet_names

# put sheets into a list of data frames
sheet_list <- lapply(sheet_names, function(.sheet){readWorksheet(object=excel_1, .sheet)})


df <- data.frame(item_num=character(), updated_cost=character())
for (i in 1:length(sheet_list)) {
data<-sheet_list[[i]]
item_num_temp<-data$Col2
item_num_temp<- tail(item_num_temp, -2)
updated_temp<-data$From.SD.11.22
updated_temp<- tail(updated_temp, -2)
df_append<-data.frame(item_num=item_num_temp,updated_cost=updated_temp)
df<-rbind(df,df_append)
}
write.csv(df,'FOB.csv')


excel_2 <- loadWorkbook('OLB.xlsx')
sheet_names <- getSheets(excel_2)
names(sheet_names) <- sheet_names

# put sheets into a list of data frames
sheet_list_1 <- lapply(sheet_names, function(.sheet){readWorksheet(object=excel_2, .sheet)})

df1 <- data.frame(item_num=character(), updated_cost=character())
for (i in 2:length(sheet_list_1)) {
  data<-sheet_list_1[[i]]
  item_num_temp<-data$Col2
  item_num_temp<- tail(item_num_temp, -2)
  updated_temp<-data$From.SD.11.22
  updated_temp<- tail(updated_temp, -2)
  df_append<-data.frame(item_num=item_num_temp,updated_cost=updated_temp)
  df1<-rbind(df1,df_append)
}
df<-rbind(df,df1)
write.csv(df,'FOB.csv')

excel_3 <- read.csv('FOB.csv')
unique<-df[!duplicated(excel_3),]
write.csv(unique,'FOB.csv')