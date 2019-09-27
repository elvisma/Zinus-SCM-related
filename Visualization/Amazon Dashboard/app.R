
##### Amazon Dashboard 2 #####

# Read Libraries
library("readxl")
library("reshape")
options(scipen = 999)
library("rsconnect")
library("shiny")
library("stringi")
library("highcharter")
library("shinythemes")
library("shinydashboard")
library("dplyr")
library("DT")

By_Date <- as.Date('2019/6/2') # last week
Import_Plan <- as.Date('2019/9/1') # from SISR/AISR
DI_Arrival <- as.Date('2019/9/8') # AISR 1 week before forecast DI 
By_Col <- 173 #AMZ Forecast tab By_Date --> Column calculate selling history (EACH WEEK +1)
Sum_Col <- 75 # make sure it cover til Last week (EACH WEEK +1)
Front_Date <- "Updated Week: 6/2/2019-6/8/2019" # Last Week

############################################################################
## Date for Flags
data_flags <- data_frame(
  date = By_Date,
  title = "Last Week"
)

data_flags_2 <- data_frame(
  date = Import_Plan,
  title = "Import Plan Week"
)

data_flags_3 <- data_frame(
  date = DI_Arrival,
  title = "last DI Arrival Week"
)

############################################################################
## 1) Read AMZ Tracker
#setwd("C:/Users/Elvis Ma/Desktop/Visualization tools/Amazon dashboard")
AMZ_DF <- read_excel('AMZ Tracker.xlsx', sheet = 'AMZ Forecast', skip = 2)

# AMZ Tracker: Trend of Inventory, WOS, Net PPM, ASP, Accuracy of Forecast
AMZ_Tracker <- read_excel('AMZ Tracker.xlsx', sheet = 1)
AMZ_Tracker <- AMZ_Tracker[c(2:17,33:38),-c(1)] #### cover to 12/22/2019

names(AMZ_Tracker) = seq(as.Date("2017/12/24"), as.Date("2019/12/22"), "week")  ##### as.Date() needs to match the column of previous code
colnames(AMZ_Tracker)[1] <- "List"
AMZ_Tracker[,2:ncol(AMZ_Tracker)] <- sapply(AMZ_Tracker[,2:ncol(AMZ_Tracker)], as.numeric)

DF <- data.frame(matrix(nrow=ncol(AMZ_Tracker)-1,ncol=1))
DF[,1] <- seq(as.Date("2017/12/31"), as.Date("2019/12/22"), "week")
colnames(DF)[1] <- "Dates"

DF$Actual.Sales <- as.numeric(AMZ_Tracker[10,-1])
DF$AMZ.Forecast <- as.numeric(AMZ_Tracker[15,-1])
DF$SCM.Forecast <- as.numeric(AMZ_Tracker[16,-1])
DF$Shipped.COGS <- as.numeric(AMZ_Tracker[11,-1])
DF$DI <- as.numeric(AMZ_Tracker[4,-1])
DF$Direct <- as.numeric(AMZ_Tracker[5,-1])
DF$DDS <- as.numeric(AMZ_Tracker[6,-1])
DF$Merchant <- as.numeric(AMZ_Tracker[7,-1])
DF$ZinusInv <- as.numeric(AMZ_Tracker[1,-1])
DF$AMZFC <- as.numeric(AMZ_Tracker[22,-1])
DF$ASP <- as.numeric(AMZ_Tracker[12,-1])
DF$Markup <- as.numeric(AMZ_Tracker[13,-1]) * 100
DF$Net.PPM <- as.numeric(AMZ_Tracker[14,-1]) * 100
DF$AMZ.ZIN.Inv <- as.numeric(AMZ_Tracker[17,-1])
DF$Actual.Forecast <- as.numeric(AMZ_Tracker[18,-1])
DF$ZINWOS <- as.numeric(AMZ_Tracker[19,-1])
DF$AMZWOS <- as.numeric(AMZ_Tracker[20,-1])
DF$AMZ.ZIN.WOS <- as.numeric(AMZ_Tracker[21,-1])

DF_Summary <- DF[1:Sum_Col,]
DF_Summary <- DF_Summary[order(DF_Summary[,1], decreasing = TRUE),]
DF_Summary <- data.frame(DF_Summary)
rownames(DF_Summary) <- 1:nrow(DF_Summary)
DF_Summary2 <- DF_Summary[,c(1,14,12,11,10,18,17,3,4,2,5,6,7,8,9)]
DF_Summary2[,2] <- DF_Summary2[,2] / 100
datatable(DF_Summary2)

line <- rep(43.5, nrow(DF))
line2 <- rep(46.5, nrow(DF))
#############################################################################
## 1) Read Top SKUs Tracker
#setwd("C:/Users/Elvis Ma/Desktop/Elvis Local/Young/2) Wednesday/AMZ Dashobard/")
TopSKU_Tracker <-read_excel("Top SKUs tracker for Dashboard.xlsx",sheet = "Top10 BY $",skip=1)
#### cover to the end of Aug 2019
names(TopSKU_Tracker) = seq(as.Date("2017/12/17"), as.Date("2019/8/25"), "week")
colnames(TopSKU_Tracker)[1] <- "ZINUSSKU"
colnames(TopSKU_Tracker)[2] <- "List"
TopSKU_Tracker[,3:ncol(TopSKU_Tracker)] <- sapply(TopSKU_Tracker[,3:ncol(TopSKU_Tracker)], as.numeric)

TopSKU_List <-read_excel("Top SKUs tracker for Dashboard.xlsx",sheet = "TOP SKU",skip=0)
TopSKUs <- c(TopSKU_List$`ZINUS SKU`)

############################################################################
## 2) Cleaning Data for "Price & Past Sales History"
AMZ_DFS <- AMZ_DF[,c(1,6,7,11,13:By_Col)]

Dates <- seq(as.Date("2016/4/10"), By_Date, "week")

names(AMZ_DFS) = seq(as.Date("2016/4/10"), By_Date, "week")
colnames(AMZ_DFS)[1] <- "SISR.Category"
colnames(AMZ_DFS)[2] <- "Class"
colnames(AMZ_DFS)[3] <- "Size"
colnames(AMZ_DFS)[4] <- "Amazon.SKU"
colnames(AMZ_DFS)[5] <- "List"

AMZ_DFS <- AMZ_DFS[,4:length(AMZ_DFS)]
AMZ_DFS[,3:ncol(AMZ_DFS)] <- sapply(AMZ_DFS[,3:ncol(AMZ_DFS)], as.numeric)

Price <- AMZ_DFS[which(AMZ_DFS$List == "Price"),-2]
Actual <- AMZ_DFS[which(AMZ_DFS$List == "Actual #"),-2]
Inv <- AMZ_DFS[which(AMZ_DFS$List == "AMZ + Zin Inv"),-2]

n <- Price$Amazon.SKU
Price_SKU <- as.data.frame(t(Price[,-1]))
colnames(Price_SKU) <- n
Price_SKU$Date <-as.Date(rownames(Price_SKU))
for (i in 1:(ncol(Price_SKU)-1)){
  Price_SKU[is.na(Price_SKU[,i]),i] <- 0
}

n1 <- Actual$Amazon.SKU
Actual_SKU <- as.data.frame(t(Actual[,-1]))
colnames(Actual_SKU) <- n1
Actual_SKU$Date <-as.Date(rownames(Actual_SKU))
for (i in 1:(ncol(Actual_SKU)-1)){
  Actual_SKU[is.na(Actual_SKU[,i]),i] <- 0
}

n2 <- Inv$Amazon.SKU
Inv_SKU <- as.data.frame(t(Inv[,-1]))
colnames(Inv_SKU) <- n2
Inv_SKU$Date <-as.Date(rownames(Inv_SKU))
for (i in 1:(ncol(Inv_SKU)-1)){
  Inv_SKU[is.na(Inv_SKU[,i]),i] <- 0
}

##########################################################################
##########################################################################
##########################################################################
## Building UI
ui <- navbarPage("Amazon Dashboard", theme = shinytheme("flatly"),
        tabPanel("Overview",
                 fluidPage(
                   h4(Front_Date),
                   br(),
                   h3("Trend of Inventory"),
                   br(),
                   fluidRow(
                    column(width = 12,
                        tabsetPanel(
                          tabPanel("Inventory", highchartOutput("DBInventory", height = "480px")),
                          tabPanel("WOS", highchartOutput("DBWOS", height = "480px"))
                   ))),
                   hr(),
#################### Top SKUs Testing UI######################                   
                   h3("Trend of Top 10 SKUs"),
                   br(),
                   sidebarLayout(
                     sidebarPanel(width = 2,
                                  selectInput("SKU_2", label = "SKU:", choices = unique(TopSKUs))),
                     mainPanel(width = 10,
                        tabsetPanel(
                          tabPanel("Inventory", highchartOutput("TopSKUInventory", height = "480px")),
                          tabPanel("WOS", highchartOutput("TopSKUWOS", height = "480px"))
                   ))),                  
                  hr(),
#################### Top SKUs Testing UI######################  

                   h3("Trend of Net PPM & ASP"),
                   br(),
                   fluidRow(
                     column(width = 12,
                            tabsetPanel(
                              tabPanel("Net PPM", highchartOutput("AMZNetPPM", height = "480px")),
                              tabPanel("ASP", highchartOutput("AMZASP", height = "480px"))
                   ))),
                   hr(),
                 
                   h3("Summary"),
                   br(),
                   fluidRow(column(11, DT::dataTableOutput("mytable9")), height="500px"),
                   br(),
                   hr(),
                   h3("Top 10 SKU & Product Line (Whole Category)"),
                   br(),
                   sidebarLayout(
                     sidebarPanel(width=2, 
                                  radioButtons("button", label = h4("Period"), choices = list("Past 12 Months" = 51, "Past 4 Months" = 15, "Past 1 Month" = 3, "Last Week" = 0))),
                     mainPanel(width=9,
                               tabsetPanel(
                                 tabPanel("SKU", br(), fluidRow(highchartOutput("top10bar", height = "380px")), br(), fluidRow(column(width=12, DT::dataTableOutput("Top10Total2"), height = "400px"))),
                                 tabPanel("Product Line", br(), fluidRow(highchartOutput("top10Gbar", height = "380px")), br(), fluidRow(column(width=12,DT::dataTableOutput("Top10Group2"), height = "400px")))
                               ))
                   ),
                   hr(),br()
                 )),
        
        tabPanel("Sales",
                 fluidPage(
                   sidebarLayout(
                     sidebarPanel(width = 2,
                                  radioButtons("radio", label = h4("Category"),
                                               choices = list("Foam Mattress" = "Foam Mattress", "Spring Mattress" = "Spring Mattress", "Frame" = "Frame", "Box Spring" = "Box Spring", "Platform Bed" = "Platform Bed", "Others" = "Others")),
                                  radioButtons("radio2", label = h4("Period"),
                                               choices = list("Past 12 Months" = 51, "Past 4 Months" = 15, "Past 1 Month" = 3, "Last Week" = 0))
                     ),
                     mainPanel(width = 9,
                               h3("Top 10 SKU & Product Line (by Category)"),
                               br(),
                               tabsetPanel(
                                 tabPanel("SKU", br(), fluidRow(highchartOutput("top10Tbar", height = "380px")), br(), fluidRow(column(width=12, DT::dataTableOutput("Top10TT"), height = "400px"))),
                                 tabPanel("Product Line", br(), fluidRow(highchartOutput("top10GGbar", height = "380px")), br(), fluidRow(column(width=12,DT::dataTableOutput("Top10GG"), height = "400px")))
                               ),
                               br()
                     )),
                   hr(),
                   sidebarLayout(
                     sidebarPanel(width = 2,
                                  selectInput("SKU_1", label = "SKU:", choices = unique(AMZ_DF$Amazon.SKU))             
                     ),
                     mainPanel(width = 10,
                               h3("Price, Past Sales & Inventory History"),
                               br(),
                               highchartOutput("AMZTrack", height = "500px")
                     )
                   ),
                   hr(),br()
                 ))
      )

##########################################################################
##########################################################################
##########################################################################
## Building Server
server <- function(input, output) {

  #Dashboard: Summary Table
  output$mytable9 <- DT::renderDataTable(DT::datatable(DF_Summary2, options = list(scrollX = TRUE, pageLength = 8, autoWidth = TRUE, columnDefs = list(list(width = '120px', targets = 1)))) %>%
                                          formatCurrency(c(3:5,8:15),'$') %>%
                                          formatRound(c(6:7),1) %>%
                                          formatPercentage(2,2) %>% 
                                          formatDate(1,'toLocaleDateString'))

  #class = 'cell-border stripe', options = list(dom = 't', ordering = F)
  
  #Dashboard: Inventory
  output$DBInventory <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(DF$Dates, DF$AMZ.ZIN.Inv, name = "Zinus + AMZ Inv", id = "ZinusAMZ") %>%
      hc_add_series_times_values(DF$Dates, DF$AMZFC, name = "AMZ FC") %>%
      hc_add_series_times_values(DF$Dates, DF$ZinusInv, name = "Zinus Inv (STD)") %>%
      hc_add_series_times_values(DF$Dates, DF$AMZ.Forecast, name = "AMZ Forecast") %>%
      hc_add_series_times_values(DF$Dates, DF$SCM.Forecast, name = "SCM Forecast") %>%
      hc_add_series_times_values(DF$Dates, DF$Actual.Sales, name = "Actual Sales", color="#f45b5b") %>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      #hc_add_series(data_flags, mapping=hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Current Date", color="#f45b5b") %>%
      #hc_add_series(data_flags_2, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Import Plan", color="#f45b5b") %>%
      #hc_add_series(data_flags_3, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "DI Arrival", color="#f45b5b") %>%
      hc_add_theme(hc_theme_gridlight())
  })
################################### testing ############################################
  
  SKU_DF <- reactive({
    Needed2 <- TopSKU_Tracker[which(TopSKU_Tracker$ZINUSSKU == input$SKU_2),]
    Needed2 <- Needed2[,-c(1)]
    SKU_DF<-data.frame(matrix(nrow=ncol(Needed2)-1,ncol=1))
    SKU_DF[,1] <- seq(as.Date("2017/12/31"), as.Date("2019/8/25"), "week")
    colnames(SKU_DF)[1] <- "Dates"
    SKU_DF$ZinusInv <- as.numeric(Needed2[1,-1])
    SKU_DF$AMZFC <- as.numeric(Needed2[2,-1])
    SKU_DF$AMZ.ZIN.Inv <- as.numeric(Needed2[3,-1])
    SKU_DF$SCM.Forecast <- as.numeric(Needed2[5,-1])
    SKU_DF$Actual.Sales <- as.numeric(Needed2[6,-1])
    ##############################################
    
    SKU_DF$Zinus.WOS <- as.numeric(Needed2[7,-1])
    SKU_DF$AMZ.WOS <- as.numeric(Needed2[8,-1])
    SKU_DF$AMZ.ZIN.WOS <- as.numeric(Needed2[9,-1])
    Selected <- SKU_DF
    Selected
    
  })
  
  output$TopSKUInventory <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$AMZ.ZIN.Inv, name = "Zinus + AMZ Inv", id = "ZinusAMZ") %>%
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$AMZFC, name = "AMZ FC") %>%
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$ZinusInv, name = "Zinus Inv") %>%
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$SCM.Forecast, name = "SCM Forecast") %>%
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$Actual.Sales, name = "Actual Sales", color="#f45b5b") %>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      #hc_add_series(data_flags, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Current Date", color="#f45b5b") %>%
      #hc_add_series(data_flags_2, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Import Plan", color="#f45b5b") %>%
      #hc_add_series(data_flags_3, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "DI Arrival", color="#f45b5b") %>%
      hc_add_theme(hc_theme_gridlight())
      
    
    
  })
  
  output$TopSKUWOS <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$AMZ.ZIN.WOS, name = "Zinus + AMZ WOS", id = "ZinusAMZ") %>%
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$AMZ.WOS, name = "AMZ WOS") %>%
      hc_add_series_times_values(SKU_DF()$Dates, SKU_DF()$Zinus.WOS, name = "Zinus WOS") %>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      #hc_add_series(data_flags, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Current Date", color="#f45b5b") %>%
      #hc_add_series(data_flags_2, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Import Plan", color="#f45b5b") %>%
      #hc_add_series(data_flags_3, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "DI Arrival", color="#f45b5b") %>%
      hc_add_theme(hc_theme_gridlight())
    
    
    
  })

  
####################################testing ending #####################################  
  #Dashboard: WOS
  output$DBWOS <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(DF$Dates, DF$AMZ.ZIN.WOS, name = "AMZ + Zinus WOS", id = "ZinusAMZ") %>%
      hc_add_series_times_values(DF$Dates, DF$AMZWOS, name = "AMZ WOS") %>%
      hc_add_series_times_values(DF$Dates, DF$ZINWOS, name = "Zinus WOS") %>%
      #hc_add_series(data_flags, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Current Date", color="#f45b5b") %>%
      #hc_add_series(data_flags_2, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "Import Plan", color="#f45b5b") %>%
      #hc_add_series(data_flags_3, hcaes(x = date), type = "flags", onSeries = "ZinusAMZ", name = "DI Arrival", color="#f45b5b") %>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      hc_add_theme(hc_theme_gridlight())
  })
  
  #Dashboard: NetPPM
  output$AMZNetPPM <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_add_series_times_values(DF$Dates, DF$Net.PPM, name = "Net PPM", color = '#7cb5ec') %>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_add_series_times_values(DF$Dates, line, name = 'Net_PPM 43.5%', color = '#f45b5b') %>%
      hc_add_series_times_values(DF$Dates, line2, name = 'Net_PPM 46.5%', color = '#f45b5b') %>%
      hc_rangeSelector(selected = 4) %>%
      hc_add_theme(hc_theme_gridlight())
  })
  
  #Dashboard: ASP
  output$AMZASP <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_add_series_times_values(DF$Dates, DF$ASP, name = "ASP") %>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      hc_add_theme(hc_theme_gridlight())
  })
  
  ## Dashboard: Top 10 SKU
  Top10Total <- reactive({
    
    StartDate = By_Col - as.numeric(input$button)
    Needed <- AMZ_DF[,c(1,10,13,StartDate:By_Col)]
    
    Needed <- Needed[which(Needed$List == 'Actual $'), ]
    Top10Total <- Needed[,1:2]
    
    colnames(Top10Total)[1] <- "SISR.Category"
    colnames(Top10Total)[2] <- "SKU"
    
    for(i in 1:nrow(Needed)){
      Top10Total$Total[i] <- round(sum(as.numeric(Needed[i,4:ncol(Needed)])),2)
    }
    
    Top10Total[is.na(Top10Total[,3]),3] <- 0
    Top10Total <- Top10Total[order(-Top10Total[,3]),]
    Top10Total <- data.frame(Top10Total)
    
    for(i in 1:nrow(Top10Total)){
      Top10Total$Percentage[i] <- round(Top10Total[i,3] / sum(Top10Total[,3]), 4)
    }
    
    Top10Total$Cumulative[1] <- Top10Total[1,4]
    
    for(i in 2:nrow(Top10Total)){
      Top10Total$Cumulative[i] <- Top10Total[i-1,5] + Top10Total[i,4]
    }
    
    colnames(Top10Total)[4] <- "% of Total"
    colnames(Top10Total)[5] <- "Cumulative %"
    
    Selected <- Top10Total[, c(2:5)]
    rownames(Selected) <- 1:nrow(Selected)
    
    Selected
  })
  
  output$Top10Total2 <- DT::renderDataTable(DT::datatable(Top10Total(), filter = 'bottom') %>%
                                          formatCurrency(2,'$') %>%
                                          formatPercentage(c(3:4),2))
  
  output$top10bar <- renderHighchart({
    highchart() %>%
      hc_add_series_df(data = Top10Total()[1:10,], type = "bar", x = SKU, y = Total) %>%
      hc_xAxis(list(categories = Top10Total()[1:10,1])) %>%
      hc_legend(enabled = FALSE) %>%
      hc_tooltip(valueDecimals = 2, valuePrefix = "$")
  })
  
  ## Dashboard: Top 10 Product Line
  Top10Group <- reactive({
    
    StartDate = By_Col - as.numeric(input$button)
    Needed_G <- AMZ_DF[,c(1,3,13,StartDate:By_Col)]
    
    Needed_G <- Needed_G[which(Needed_G$List == 'Actual $'), ]
    Top10Total_G <- Needed_G[,1:2]
    
    colnames(Top10Total_G)[1] <- "SISR.Category"
    colnames(Top10Total_G)[2] <- "SKU"
    
    for(i in 1:nrow(Needed_G)){
      Top10Total_G$Total[i] <- round(sum(as.numeric(Needed_G[i,4:ncol(Needed_G)])),2)
    }
    
    basic_summ <- group_by(Top10Total_G, `SKU`)
    basic_summ <- summarise(basic_summ, Total = sum(`Total`))
    basic_summ[is.na(basic_summ[,2]),2] <- 0
    
    basic_summ <- merge(basic_summ, unique(Top10Total_G[,1:2]), by = 'SKU', all.x = TRUE)
    basic_summ <- basic_summ[order(-basic_summ[,2]),]
    basic_summ <- data.frame(basic_summ)
    
    for(i in 1:nrow(basic_summ)){
      basic_summ$Percentage[i] <- round(basic_summ[i,2] / sum(basic_summ[,2]), 4)
    }
    
    basic_summ$Cumulative[1] <- basic_summ[1,4]
    
    for(i in 2:nrow(basic_summ)){
      basic_summ$Cumulative[i] <- basic_summ[i-1,5] + basic_summ[i,4]
    }
    
    colnames(basic_summ)[4] <- "% of Total"
    colnames(basic_summ)[5] <- "Cumulative %"
    
    basic_summ <- basic_summ[,c(3,1:2,4:5)]
    
    Selected_G <- basic_summ[, c(2:5)]
    rownames(Selected_G) <- 1:nrow(Selected_G)
    
    Selected_G
  })
  
  output$Top10Group2 <- DT::renderDataTable(DT::datatable(Top10Group(), filter = 'bottom') %>%
                                          formatCurrency(2,'$') %>%
                                          formatPercentage(c(3:4),2))

  output$top10Gbar <- renderHighchart({
    highchart() %>%
      hc_add_series_df(data = Top10Group()[1:10,], type = "bar", x = SKU, y = Total) %>%
      hc_xAxis(list(categories = Top10Group()[1:10,1])) %>%
      hc_legend(enabled = FALSE) %>%
      hc_tooltip(valueDecimals = 2, valuePrefix = "$")
  })

  ## Sales: Top 10 SKU
  Top10T <- reactive({
    
    StartDate = By_Col - as.numeric(input$radio2)
    Needed <- AMZ_DF[,c(1,10,13,StartDate:By_Col)]
    
    Needed <- Needed[which(Needed$List == 'Actual $'), ]
    Top10Total <- Needed[,1:2]
    
    colnames(Top10Total)[1] <- "SISR.Category"
    colnames(Top10Total)[2] <- "SKU"
    
    for(i in 1:nrow(Needed)){
      Top10Total$Total[i] <- round(sum(as.numeric(Needed[i,4:ncol(Needed)])),2)
    }
    
    Top10Total[is.na(Top10Total[,3]),3] <- 0
    Top10Total <- Top10Total[order(-Top10Total[,3]),]
    Top10Total <- data.frame(Top10Total)
    
    for(i in 1:nrow(Top10Total)){
      Top10Total$Percentage[i] <- round(Top10Total[i,3] / sum(Top10Total[,3]), 4)
    }
    
    Top10Total$Cumulative[1] <- Top10Total[1,4]
    
    for(i in 2:nrow(Top10Total)){
      Top10Total$Cumulative[i] <- Top10Total[i-1,5] + Top10Total[i,4]
    }
    
    colnames(Top10Total)[4] <- "% of Total"
    colnames(Top10Total)[5] <- "Cumulative %"
    
    Selected <- Top10Total[which(Top10Total$SISR.Category == input$radio), c(2:5)]
    rownames(Selected) <- 1:nrow(Selected)
    
    Selected
  })
  
  output$Top10TT <- DT::renderDataTable(DT::datatable(Top10T(), filter = 'bottom') %>%
                                          formatCurrency(2,'$') %>%
                                          formatPercentage(c(3:4),2))
  
  output$top10Tbar <- renderHighchart({
    highchart() %>%
      hc_add_series_df(data = Top10T()[1:10,], type = "bar", x = SKU, y = Total) %>%
      hc_xAxis(list(categories = Top10T()[1:10,1])) %>%
      hc_legend(enabled = FALSE) %>%
      hc_tooltip(valueDecimals = 2, valuePrefix = "$")
  })
  
  ## Sales: Top 10 Product Line
  Top10G <- reactive({
    
    StartDate = By_Col - as.numeric(input$radio2)
    Needed_G <- AMZ_DF[,c(1,3,13,StartDate:By_Col)]
    
    Needed_G <- Needed_G[which(Needed_G$List == 'Actual $'), ]
    Top10Total_G <- Needed_G[,1:2]
    
    colnames(Top10Total_G)[1] <- "SISR.Category"
    colnames(Top10Total_G)[2] <- "SKU"
    
    for(i in 1:nrow(Needed_G)){
      Top10Total_G$Total[i] <- round(sum(as.numeric(Needed_G[i,4:ncol(Needed_G)])),2)
    }
    
    basic_summ <- group_by(Top10Total_G, `SKU`)
    basic_summ <- summarise(basic_summ, Total = sum(`Total`))
    basic_summ[is.na(basic_summ[,2]),2] <- 0
    
    basic_summ <- merge(basic_summ, unique(Top10Total_G[,1:2]), by = 'SKU', all.x = TRUE)
    basic_summ <- basic_summ[order(-basic_summ[,2]),]
    basic_summ <- data.frame(basic_summ)
    
    for(i in 1:nrow(basic_summ)){
      basic_summ$Percentage[i] <- round(basic_summ[i,2] / sum(basic_summ[,2]), 4)
    }
    
    basic_summ$Cumulative[1] <- basic_summ[1,4]
    
    for(i in 2:nrow(basic_summ)){
      basic_summ$Cumulative[i] <- basic_summ[i-1,5] + basic_summ[i,4]
    }
    
    colnames(basic_summ)[4] <- "% of Total"
    colnames(basic_summ)[5] <- "Cumulative %"
    
    basic_summ <- basic_summ[,c(3,1:2,4:5)]
    
    Selected_G <- basic_summ[which(basic_summ$SISR.Category == input$radio), c(2:5)]
    rownames(Selected_G) <- 1:nrow(Selected_G)
    
    Selected_G
  })
  
  output$Top10GG <- DT::renderDataTable(DT::datatable(Top10G(), filter = 'bottom') %>%
                                          formatCurrency(2,'$') %>%
                                          formatPercentage(c(3:4),2))
  
  output$top10GGbar <- renderHighchart({
    highchart() %>%
      hc_add_series_df(data = Top10G()[1:10,], type = "bar", x = SKU, y = Total) %>%
      hc_xAxis(list(categories = Top10G()[1:10,1])) %>%
      hc_legend(enabled = FALSE) %>%
      hc_tooltip(valueDecimals = 2, valuePrefix = "$")
  })
  
  # Sales: SKU Tracker: Price, Past Sales, Inventory
  output$AMZTrack <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(Price_SKU$Date, 
                                 Price_SKU[,input$SKU_1],
                                 name = "Price", color = '#f7a35c', yAxis = 0) %>%
      hc_add_series_times_values(Actual_SKU$Date, 
                                 Actual_SKU[,input$SKU_1],
                                 name = "Actual #", color = '#7cb5ec', yAxis = 1) %>%
      hc_add_series_times_values(Inv_SKU$Date, 
                                 Inv_SKU[,input$SKU_1],
                                 name = "AMZ + Zin Inv", color = '#e4d354', yAxis = 1) %>%
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      hc_add_theme(hc_theme_gridlight())
  })
}

## Deploy App
shinyApp(ui, server)
