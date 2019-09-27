##### read libraries #####

library(readxl)
library(shiny)
library(xlsx)
library(openxlsx)
options(scipen = 999)
library("stringi")
library("highcharter")
library(rsconnect)
library("shinythemes")
library("shinydashboard")
library("dplyr")
library("DT")
library(tidyr)
####################################

####################################
## read massage data
setwd("C:/Users/Elvis Ma/Desktop/Visualization tools/Forecast Accuracy dashboard")

Date_Range <- "Updated: 6/17/2019"

############################################
#########################################
## Building UI
ui <- navbarPage("Forecast Accuracy Summary (SCM Team)", theme = shinytheme("sandstone"),
                 tabPanel("Actual $ vs. Forecast $",
                     fluidPage(
                        sidebarLayout(
                           sidebarPanel(width =2,
                                       checkboxGroupInput("variable1","Choose Accounts:",
                                                          c("Amazon"="AMZ","Costco","Walmart","Sam's Club","Wayfair","Overstock","Target","Homedepot","Zinus.com","MACY's.com"="Macys","CHEWY.com"="CHEWY.COM"),
                                                          selected = c("Walmart","Sam's Club","Wayfair","Costco","Zinus.com")
                                       ),
                                       checkboxGroupInput("variable2","Choose Classes:",
                                                          c("Class A"="A","Class B"="B","Class C"="C","TOTAL"),
                                                          selected = c("TOTAL")
                                       ),
                                       selectInput("dateRange","Date Range Selection:",
                                                   c("Half Jun 2019","May 2019","Apr 2019","Mar 2019","Feb 2019","Jan 2019","Dec 2018","Nov 2018"),
                                                   selected = c("Half Jun 2019")
                                       ),
                                       downloadButton("downloadData","Download")
                           ),
                           mainPanel(width = 10,
                                    h3(Date_Range),
                                    br(),
                                    tabsetPanel(
                                      tabPanel("Chart",highchartOutput("ComparisonChart",height = "480px")),
                                      tabPanel("Table",dataTableOutput("ComparisonTable"))
                                    ))),
                                    hr(),
                          ### Newly Added ###
                          h3("Forecast Accuracy & MAPE (total classes)"),
                          br(),
                          sidebarLayout(
                            sidebarPanel(width = 2,
                                         radioButtons("variable5",label=h4("Choose Accounts:"),
                                                            choices=list("Amazon"="AMZ","Costco","Walmart","Sam's Club","Wayfair","Overstock","Target","Homedepot","Zinus.com","MACY's.com"="Macys","CHEWY.com"="CHEWY.COM")
                                                           
                                         )
                              
                            ),
                            
                            mainPanel(width=10,
                                      br(),
                                      highchartOutput("MAPEtracker",height = "480px")
                                      )
                          ),
                          br(),
                          hr(),
                          hr(),br()
                          ### Newly Added ended ###
                          )),
                 tabPanel("Rolling Actual  vs. Monthly Rolling Forecast",
                          sidebarPanel(width =2,
                                       checkboxGroupInput("variable3","Choose Accounts:",
                                                          c("Amazon"="AMZ","Costco","Walmart","Sam's Club","Wayfair","Overstock","Target","Homedepot","Zinus.com","MACY's.com"="Macys","CHEWY.com"="CHEWY.COM"),
                                                          selected = c("Walmart","Sam's Club","Wayfair","Costco","Zinus.com")
                                       ),
                                       checkboxGroupInput("variable4","Choose Classes:",
                                                          c("Class A"="A","Class B"="B","Class C"="C","TOTAL"),
                                                          selected = c("TOTAL")
                                       ),
                                       selectInput("dateRange2","Date Range Selection:",
                                                   c("Half Jun 2019","May 2019","Apr 2019","Mar 2019","Feb 2019","Jan 2019","Dec 2018","Nov 2018"),
                                                   selected = c("Half Jun 2019")
                                       )
                          ),
                          mainPanel(width = 10,
                                    h3(Date_Range),
                                    br(),
                                    tabsetPanel(
                                      tabPanel("Accomplishment By $", highchartOutput("DollarChart", height = "480px")),
                                      tabPanel("Accomplishment Rate By %",highchartOutput("PercentChart", height = "480px"))
                                    )
                                    
                          ))
                 
                 
)
######################################
##
#######################################################################

server <-function(input,output) {
  
  output$ComparisonTable <-DT::renderDataTable(
    DT::datatable(ReportFunction()[,c("ABC","ACCOUNTS","Actual","Forecast","Monthly Forecast Line","ActualPercent","DifferencePercent")]) %>%
      formatCurrency(c(3,4,5),"$") %>%
      formatPercentage(c(6,7),2)
  )
  
  # output$ComparisonChart <- renderPlotly(
  
  #  plot_ly(ReportFunction(),x=~ReportFunction()$Combine,y=~ReportFunction()$Actual,text=paste("$",round(as.numeric(ReportFunction()$Actual)/1000000,2),"M"),textposition="auto",type ="bar", name = "Actual $",marker = list(color='#2cd061')) %>%
  
  #   add_trace(y=~ReportFunction()$Forecast, type ="bar",name="Forecast $",text=paste("$",round(as.numeric(ReportFunction()$Forecast)/1000000,2),"M"),textposition = "auto",marker = list(color='#7cb5ec')) %>%
  #  layout(xaxis = list(title = "BY Channel & Class"), yaxis = list(title = 'BY $ Amount'), barmode = 'group')
  
  #)
  
  output$ComparisonChart <-renderHighchart({
    highchart()%>%
      hc_title(text = "Actual Sales V.S. Rolling Forecast")%>%
      hc_tooltip(valueDecimals = 2) %>%
      
      
      #hc_yAxis_multiples(create_yaxis(naxis = 2, heights = c(3, 1))) %>% 
      #hc_yAxis(title=list(text="% Inaccuracy Rate"))%>%
      hc_yAxis_multiples(
        
        list(title=list(text="$ Amount")),
        list(title=list(text="Forecast MAPE"),opposite=TRUE)
        
      ) %>% 
      #hc_yAxis(title=list("% Inaccuracy","$ Amount"))%>%
      hc_xAxis(list(categories = ReportFunction()$Combine)) %>%
      hc_plotOptions(line=list(
                               dataLabels=list(enabled=TRUE),
                               enableMouseTracking = TRUE
      ))%>%
      
      
      hc_add_series(data = ReportFunction()$Forecast, type = "column", name="Forecast $",yAxis=0) %>%
      #hc_add_series_df(data = ReportFunction(), type = "line", x = Combine, y = variance , group="Inaccuracy Rate",yAxis=1) %>%
      hc_add_series(data=ReportFunction()$variance,type="spline",name="Forecast MAPE",yAxis=1)%>%
      hc_add_series(data = ReportFunction()$Actual, type = "column", name="Actual $",yAxis=0) %>%
      hc_legend(enabled = TRUE) %>%
      hc_add_theme(hc_theme_gridlight())
    
  })
  
  
  # output$DollarChart <-renderPlotly(
  #  plot_ly(ComparisonFunction(),x=~ComparisonFunction()$Combine,y=~ComparisonFunction()$Actual,text=paste("$",round(as.numeric(ComparisonFunction()$Actual)/1000000,2),"M"),textposition="auto",type ="bar",name = "Actual $",marker = list(color='#2cd061')) %>%
  
  #   add_trace(y=~ComparisonFunction()$Difference, text=paste("$",round(as.numeric(ComparisonFunction()$Difference)/1000000,2),"M"),textposition="auto",type ="bar",name="$ Gap to Fill",marker = list(color="#f7a35c")) %>%
  #  layout(xaxis = list(title = "BY Channel & Class"), yaxis = list(title = 'Accomplishment By $'), barmode = 'stack')
  
  #)
  
  output$DollarChart <- renderHighchart({
    
    highchart()%>%
      
      hc_title(text = "Sales Accomplishment $")%>%
      #hc_xAxis(list(categories = ComparisonFunction()$Combine)) %>%
      hc_yAxis(title=list(text="$ Amount"))%>%
      
      hc_xAxis(list(categories = ComparisonFunction()$Combine)) %>%
      
      hc_tooltip(valueDecimals = 2, valuePrefix = "$")%>%
      hc_plotOptions(column=list(color="#32CD32",
                                 dataLabels=list(enabled=FALSE),
                                 stacking="normal",
                                 enableMouseTracking = TRUE
      ))%>%
      
      hc_add_series(data = ComparisonFunction()$`Monthly Forecast Line`, type = "line",  name="Monthly Forecast $") %>%
      hc_add_series(data = ComparisonFunction()$Actual, type = "column", name="Accomplished $") %>%
      
      
      hc_legend(enabled = TRUE)
  })
  
  
  
  # output$PercentChart<-renderPlotly(
  #  plot_ly(ComparisonFunction(),x=~ComparisonFunction()$Combine,y=~ComparisonFunction()$ActualPercent,text=paste(round(100*ComparisonFunction()$ActualPercent,1),"%"),textposition="auto",type ="bar",name = "Actual %", marker=list(color='#2cd061')) %>%
  
  #   add_trace(y=~ComparisonFunction()$DifferencePercent,text=paste(round(100*ComparisonFunction()$DifferencePercent,1),"%"),textposition="auto", textposition="auto",name="% Gap to Fill",marker=list(color="#f7a35c")) %>%
  #  layout(xaxis = list(title = "BY Channel & Class"), yaxis = list(title = 'Accomplishment By %'), barmode = 'stack')
  
  #)
  
  
  output$PercentChart <- renderHighchart({
    
    highchart()%>%
      
      hc_title(text = "Sales Accomplishment Rate %")%>%
      #hc_xAxis(list(categories = ComparisonFunction()$Combine)) %>%
      hc_xAxis(list(categories = ComparisonFunction()$Combine)) %>%
      hc_yAxis(title=list(text="% Accomplishment")
      )%>%
      hc_tooltip(pointFormat="{point.percentage:.1f}%</b>")%>%
      hc_plotOptions(column=list(
        dataLabels=list(enabled=FALSE),
        stacking="normal",
        enableMouseTracking = TRUE
      ))%>%
      hc_add_series(data = ComparisonFunction()$DifferencePercent, type = "column", name ="% Gap to fill") %>%
      hc_add_series(data = ComparisonFunction()$ActualPercent, type = "column",  name ="Accomplishment %") %>%
      
      
      
      #hc_tooltip(valueDecimals = 2, )%>%
      
      
      #hc_series(list(name="Monthly Forecast",data=ComparisonFunction()$DifferencePercent),
      #         list(name="Actual Sales",data=ComparisonFunction()$ActualPercent)
      
      # )%>%
      hc_legend(enabled = TRUE) %>%
      hc_add_theme(hc_theme_gridlight())
  })
  #########################################
  output$downloadData <-downloadHandler(
    ###### this function only works in the browser window !!!!!!!!! #############
    filename = function() {
      paste(input$dateRange, length(input$variable1),"Accounts Comparison.csv", sep = " ")
    },
    
    
    content = function(file) {
      write.csv(ReportFunction(),file,row.names = FALSE)
    }
  )
  #########################################
  
  output$MAPEtracker <-renderHighchart(
    highchart()%>%
      
      hc_title(text = "Time Series Performance")%>%
      ######
      hc_xAxis(list(categories = MAPEFunction()$Month)) %>%
      
      hc_yAxis(title=list(text="Forecast Accuracy & MAPE")
      )%>%

      #hc_yAxis(title=list("% Inaccuracy","$ Amount"))%>%
      #hc_tooltip(pointFormat="{point.percentage:.1f}%</b>")%>%
      hc_plotOptions(spline=list(color="#FF0000",
        dataLabels=list(enabled=TRUE),
        enableMouseTracking = TRUE
      ))%>%
      hc_plotOptions(line=list(color="#32CD32",
                                 dataLabels=list(enabled=TRUE),
                                 enableMouseTracking = TRUE
      ))%>%
      hc_add_series(data=MAPEFunction()$roundMAPE, type="spline",name="Forecast MAPE") %>%
      hc_add_series(data=MAPEFunction()$roundFA, type="line",name="Forecast Accuracy") %>%
      
      #hc_add_series_df(MAPEFunction(), type="spline",x=Month,y=roundMAPE, group="Forecast MAPE") %>%
      #hc_add_series_df(MAPEFunction(), type="line",x=Month,y=roundFA, group="Forecast Accuracy") %>%
      
      #hc_add_series(data = Costco_MAPE$roundMAPE, type = "spline", name="Costco",yAxis=0) %>%
      #hc_add_series(data = WF_MAPE$roundMAPE, type = "spline", name="Wayfair",yAxis=0) %>%
      #hc_add_series(data = SC_MAPE$roundMAPE, type = "spline", name="Sam's Club",yAxis=0) %>%
      #hc_add_series(data = WMT_MAPE$roundMAPE, type = "spline", name="Walmart",yAxis=0) %>%
      #hc_add_series(data = ZINUS_MAPE$roundMAPE, type = "spline", name="Zinus.com",yAxis=0) %>%
      #hc_add_series(data = OS_MAPE$roundMAPE, type = "spline", name="Overstock",yAxis=0) %>%
      #hc_add_series(data = Target_MAPE$roundMAPE, type = "spline", name="Target",yAxis=0) %>%
      #hc_add_series(data = HD_MAPE$roundMAPE, type = "spline", name="Homedepot",yAxis=0) %>%
      #hc_add_series(data = AMZ_MAPE$roundMAPE, type = "spline", name="Amazon",yAxis=0) %>%
      #hc_add_series(data = Macys_MAPE$roundMAPE, type = "spline", name="Macys",yAxis=0) %>%
      
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE)%>%
      #hc_rangeSelector(selected = 4)
      #hc_legend(enabled = TRUE) %>%
      hc_add_theme(hc_theme_sandsignika())
  )
  
  
  MAPEFunction <-reactive({
    validate(
      need(!is.null(input$variable5), "PLEASE SELECT ACCOUNTS TO COMPARE")
    )
    MAPE_Data <-read_xlsx("Data for visualization.xlsx",sheet="ForecastAccuracy", skip=0)
    MAPE_Data <- select(MAPE_Data,Month,Account, roundMAPE,roundFA)
    MAPE_Data$Month=as.character(MAPE_Data$Month)
    MAPE_Return <- filter(MAPE_Data,Account==input$variable5)
    MAPE_Return <-as.data.frame(MAPE_Return)
    return(MAPE_Return)
    
  })
  
  
  ReportFunction <- reactive({
    validate(
      need(!is.null(input$variable1), "PLEASE SELECT ACCOUNTS TO COMPARE")
    )
    validate(
      need(!is.null(input$variable2),"PLEASE SELECT CLASSES TO COMPARE")
    )
    #########################################################################################################
    Report_Data <- read_xlsx("Data for visualization.xlsx", sheet = input$dateRange, skip=0)
    Report_Data$Combine <-paste(Report_Data$ACCOUNTS,Report_Data$ABC,sep=" ")
    #Report_Data$variance <-paste(round(Report_Data$variance,2)*100,"%",sep='')
    ########################################################################################################
    Report_DF <- Report_Data[which(Report_Data$ABC%in%input$variable2&Report_Data$ACCOUNTS%in%input$variable1),]
    Report_DF <-as.data.frame(Report_DF)
    
    return(Report_DF)
  })
  
  ComparisonFunction <- reactive({
    validate(
      need(!is.null(input$variable3), "PLEASE SELECT ACCOUNTS TO COMPARE")
    )
    validate(
      need(!is.null(input$variable4),"PLEASE SELECT CLASSES TO COMPARE")
    )
    
    #########################################################################################################
    Comparison_Data <- read_xlsx("Data for visualization.xlsx", sheet = input$dateRange2, skip=0)
    Comparison_Data$Combine <- paste(Comparison_Data$ACCOUNTS,Comparison_Data$ABC,sep=" ")
    #######################################################################################################
    Comparison_DF <- Comparison_Data[which(Comparison_Data$ABC%in%input$variable4&Comparison_Data$ACCOUNTS%in%input$variable3),]
    for(i in 1:nrow(Comparison_DF))
      if(Comparison_DF$Difference[i]<0)
        Comparison_DF$Difference[i]=0
    for(i in 1:nrow(Comparison_DF))
      if(Comparison_DF$DifferencePercent[i]<0)
      {Comparison_DF$DifferencePercent[i]=0
      Comparison_DF$ActualPercent[i]=1}
    
    Comparison_DF 
  })
  
  
}


shinyApp(ui, server)


