library(xlsx)
library(openxlsx)
library(plyr)
options(scipen = 999)

setwd("//192.168.1.41/03.Demand Planning/SISR 09.23.2019/MasterPO/")
file_list <- list.files(pattern = "*.csv")
date <-"09232019" ####### date to set 


###### separate data ##########
for(i in 1:length(file_list)){
  assign(file_list[i],
         read.csv(file_list[i])
  )
}
############ combined data ##############
######### using do.call() to combine #######
data <- do.call("rbind",lapply(file_list,function(file)
                                         read.csv(file)
  ))

######## using Reduce() to combine ########
#data <-Reduce("rbind",lapply(file_list,function(file)
  #                                    read.csv(file)
 # ))

########## change headers' names
colnames(data)[1] <- "PO"
colnames(data) <-gsub(pattern = "\\.",replacement = " ",colnames(data))

######## write the file #################
xlsx::write.xlsx(data[,-c(ncol(data))],paste("MasterPO_",date,".xlsx",sep = ''),row.names = FALSE)