library(xlsx)
library(openxlsx)
library(plyr)
options(scipen = 999)

setwd("//zinussvr02/Shared/Accpacadv/Young/Forecast Accuracy Summary")
file_list <- list.files(pattern = "*.csv")
for(i in 1:length(file_list)){
  assign(file_list[i],
         read.csv(file_list[i])
  )
}

file_name=list()
for(i in 1:length(file_list)){
  file_name[i] = gsub(pattern = ".csv",replacement = "",file_list[i])
}



setwd("//zinussvr02/Shared/Accpacadv/Young/Forecast Accuracy Summary")
  xlsx::write.xlsx(read.csv(file_list[1]),"Combined Forecast.xlsx",sheetName = file_name[[1]],row.names = FALSE)
for(i in 2:length(file_list)){
  xlsx::write.xlsx(read.csv(file_list[i]),"Combined Forecast.xlsx",sheetName = file_name[[i]],append=TRUE,row.names = FALSE)
}