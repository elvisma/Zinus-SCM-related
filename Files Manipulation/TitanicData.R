setwd("C:/Users/Elvis Ma/Downloads")
train <- read.csv("train.csv", header = TRUE)
Test <-read.csv("test.csv", header =TRUE)

test.survived <- data.frame(Survived = rep("none",nrow(Test)),Test[,])

### ? function()
data.combined <-rbind(train, test.survived)

#get the structure of data frame
str(data.combined)?fa

library(RCurl)
library(bitops)
url <- "http://www.basketball-reference.com/boxscores/201506140GSW.html"
data <- readLines(url)

install.packages("bitops")

library(rvest)
library(xml2)
page <- read_html(url)
table <- html_nodes(page, ".stats_table")[3]
rows <- html_nodes(table, "tr")
cells <- html_nodes(rows, "td a")
teams <- html_text(cells)

extractRow <- function(rows, i){
  if(i == 1){
    return
  }
  row <- rows[i]
  tag <- "td"
  if(i == 2){
    tag <- "th"
  }
  items <- html_nodes(row, tag)
  html_text(items)
}

scrapeData <- function(team){
  teamData <- html_nodes(page, paste("#",team,"_basic", sep=""))
  rows <- html_nodes(teamData, "tr")
  lapply(seq_along(rows), extractRow, rows=rows) 
}

data <- lapply(teams, scrapeData)