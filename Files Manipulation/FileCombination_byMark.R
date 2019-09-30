install.packages("xlsx")
install.packages("gdata")
install.packages("rJava")

require(gdata)
require(rJava)
require(xlsx)

#Read csv
path <- "C:/Users/MarkHo/Desktop/INV Feed History/EndingFeed"
setwd(path)
filenames <- list.files(path)
master_df <- do.call("rbind", lapply(filenames, read.csv, header = FALSE))
write.csv(master_df, 'master_df.csv')


#Read xlsx
#df = read.xlsx("Zinus_OUT_OESHIP_20170511075648539324011.xlsx", sheetIndex = 1, colNames=TRUE)
path <-  "C:/Users/MarkHo/Desktop/Summer"
setwd(path)
filenames <- list.files(path,pattern='*.xlsx')
master_df <- do.call("rbind", lapply(filenames, read.xlsx, sheetIndex=1, colNames=TRUE, rowIndex=NULL))
write.xlsx(master_df, paste('Shipment',"-", Sys.Date(),".xlsx",sep=""),row.names=FALSE)
           