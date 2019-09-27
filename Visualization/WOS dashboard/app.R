#########################################
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
library(ggplot2)
library(plyr)

######## set file date ########
WOS_Report_Date <-"04072019"
Update_Date <- "Updated Week: 4/14/2019"
Last_Week <- "4/7/2019"
#Last_Week_Date <-as.Date("12/2/2018",format = "%m/%d/%Y")
## Date for Flags
#data_flags <- dplyr::data_frame(
  #Week_Date = Last_Week_Date,
  #title = "Last Week"
#)

############### Read the file #####################
#setwd("C:/Users/Elvis Ma/Desktop/Visualization tools/WOS dashboard")
ZINUS_WOS <- openxlsx::read.xlsx(paste("WOS_Report_",WOS_Report_Date,".xlsx",sep=""),sheet="Report_ZINUS")
WF_WOS <- openxlsx::read.xlsx(paste("WOS_Report_",WOS_Report_Date,".xlsx",sep=""),sheet="Report_Wayfair")

##################  zinus file #######################
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="A"),"Category/Class"] <-"Total_A"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="B"),"Category/Class"] <-"Total_B"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="C"),"Category/Class"] <-"Total_C"

ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="FM"),"Category/Class"] <-"FM_Total"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="SM"),"Category/Class"] <-"SM_Total"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="FR"),"Category/Class"] <-"FR_Total"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="BX"),"Category/Class"] <-"BX_Total"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="PB"),"Category/Class"] <-"PB_Total"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="OT"),"Category/Class"] <-"OT_Total"

ZINUS_WOS[which(is.na(ZINUS_WOS$`Category/Class`)),"Category/Class"] <-"Total_NA"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="NB"),"Category/Class"] <-"Total_NB"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="NC"),"Category/Class"] <-"Total_NC"

ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="Discontinued"),"Category/Class"] <-"Total_Discontinued"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="New"),"Category/Class"] <-"Total_New"
ZINUS_WOS[which(ZINUS_WOS$`Category/Class`=="Total_ZINUS"),"Category/Class"] <-"Total_Total"

################### Wayfair file ########################
WF_WOS[which(WF_WOS$`Category/Class`=="A"),"Category/Class"] <-"Total_A"
WF_WOS[which(WF_WOS$`Category/Class`=="B"),"Category/Class"] <-"Total_B"
WF_WOS[which(WF_WOS$`Category/Class`=="C"),"Category/Class"] <-"Total_C"

WF_WOS[which(WF_WOS$`Category/Class`=="FM"),"Category/Class"] <-"FM_Total"
WF_WOS[which(WF_WOS$`Category/Class`=="SM"),"Category/Class"] <-"SM_Total"
WF_WOS[which(WF_WOS$`Category/Class`=="FR"),"Category/Class"] <-"FR_Total"
WF_WOS[which(WF_WOS$`Category/Class`=="BX"),"Category/Class"] <-"BX_Total"
WF_WOS[which(WF_WOS$`Category/Class`=="PB"),"Category/Class"] <-"PB_Total"
WF_WOS[which(WF_WOS$`Category/Class`=="OT"),"Category/Class"] <-"OT_Total"

WF_WOS[which(WF_WOS$`Category/Class`=="Discontinued"),"Category/Class"] <-"Total_Discontinued"
WF_WOS[which(WF_WOS$`Category/Class`=="Total"),"Category/Class"] <-"Total_Total"

###################### Cleaning data #####################

 ############### ZINUS ############################ 
  ###### massage data ##########
  class_data1 <-ZINUS_WOS[which(ZINUS_WOS$List%in%c("Total Ending Inventory","Total Forecast Qty")&
                          ZINUS_WOS$`Category/Class`%in%grep(pattern = "^Total_*",ZINUS_WOS$`Category/Class`,value = TRUE)&
                          ZINUS_WOS$`Category/Class`!="Total_Total"),]
  Zinus_Class <- class_data1[,c(1,2)]
  
  Zinus_Class$`Last Week` <-class_data1[,c(Last_Week)] 
  
  #for(i in 1: nrow(Zinus_Class)){
   # Zinus_Class$Percentage[i] = Zinus_Class$`Last Week`[i]/sum(Zinus_Class[which(Zinus_Class[,1]==Zinus_Class[i,1]),"Last Week"])
   # Zinus_Class$Label[i] = scales::percent(Zinus_Class$Percentage[i])
  #}

  colnames(Zinus_Class) <- gsub(pattern = "Category/",replacement = '',colnames(Zinus_Class))
  Zinus_Class$Class <- gsub(pattern = "Total_",replacement = "",Zinus_Class$Class)
  
  ZinusInv_Class <-Zinus_Class[which(Zinus_Class$List=="Total Ending Inventory"),]
  ZinusSales_Class <-Zinus_Class[which(Zinus_Class$List=="Total Forecast Qty"),]


 ##### massage data ##########
 category_data1 <-ZINUS_WOS[which(ZINUS_WOS$List%in%c("Total Ending Inventory","Total Forecast Qty")&
                          ZINUS_WOS$`Category/Class`%in%grep(pattern = ".*_Total$",ZINUS_WOS$`Category/Class`,value = TRUE)&
                          ZINUS_WOS$`Category/Class`!="Total_Total"),]
 Zinus_Category <- category_data1[,c(1,2)]
  
  Zinus_Category$`Last Week` <-category_data1[,c(Last_Week)] 
  #for(i in 1: nrow(Zinus_Category)){
    #Zinus_Category$Percentage[i] = Zinus_Category$`Last Week`[i]/sum(Zinus_Category[which(Zinus_Category[,1]==Zinus_Category[i,1]),"Last Week"])
    #Zinus_Category$Label[i] = scales::percent(Zinus_Category$Percentage[i])
  #}
  
  colnames(Zinus_Category) <- gsub(pattern = "/Class",replacement = '',colnames(Zinus_Category))
  Zinus_Category$Category <- gsub(pattern = "_Total",replacement = "",Zinus_Category$Category)

  ZinusInv_Category <-Zinus_Category[which(Zinus_Category$List=="Total Ending Inventory"),]
  ZinusSales_Category <-Zinus_Category[which(Zinus_Category$List=="Total Forecast Qty"),]
  
################ WFCG #################################
  ###### massage data ##########
  class_data2 <-WF_WOS[which(WF_WOS$List%in%c("CG Ending Inventory","CG Forecast Number")&
                                  WF_WOS$`Category/Class`%in%grep(pattern = "^Total_*",WF_WOS$`Category/Class`,value = TRUE)&
                                  WF_WOS$`Category/Class`!="Total_Total"),]
  WF_Class <- class_data2[,c(1,2)]
  
  WF_Class$`Last Week` <-class_data2[,c(Last_Week)] 
  
  colnames(WF_Class) <- gsub(pattern = "Category/",replacement = '',colnames(WF_Class))
  WF_Class$Class <- gsub(pattern = "Total_",replacement = "",WF_Class$Class)
  
  WFInv_Class <-WF_Class[which(WF_Class$List=="CG Ending Inventory"),]
  WFSales_Class <-WF_Class[which(WF_Class$List=="CG Forecast Number"),]
  
  
  ##### massage data ##########
  category_data2 <-WF_WOS[which(WF_WOS$List%in%c("CG Ending Inventory","CG Forecast Number")&
                                     WF_WOS$`Category/Class`%in%grep(pattern = ".*_Total$",WF_WOS$`Category/Class`,value = TRUE)&
                                     WF_WOS$`Category/Class`!="Total_Total"),]
  
  WF_Category <- category_data2[,c(1,2)]
  WF_Category$`Last Week` <-category_data2[,c(Last_Week)] 
 
  colnames(WF_Category) <- gsub(pattern = "/Class",replacement = '',colnames(WF_Category))
  WF_Category$Category <- gsub(pattern = "_Total",replacement = "",WF_Category$Category)
  
  WFInv_Category <-WF_Category[which(WF_Category$List=="CG Ending Inventory"),]
  WFSales_Category <-WF_Category[which(WF_Category$List=="CG Forecast Number"),]
  

  
############## build UI #####################
ui <- navbarPage("WOS Report", theme = shinytheme("flatly"),
                 tabPanel("ZINUS",
                   fluidPage(
                     h4(Update_Date),
                     br(),
                     h3("Trend of Zinus Inventory"),
                     br(),
                     sidebarLayout(
                          sidebarPanel(width = 2,
                                    radioButtons("categoryInput", label = h4("Choose Category:"),
                                                 choices = list("All Cateogories"="Total","Foam Mattress" = "FM", "Spring Mattress" = "SM", "Frame" = "FR", "Box Spring" = "BX", "Platform Bed" = "PB", "Others" = "OT")
                                                ),
                                    selectInput("classInput", label = h4("Choose Class:"),
                                                 choices = list("Total","A","B","C","Discontinued","New","New A"="NA","New B"="NB","New C"="NC")
                                                 )
                                       ),
                          mainPanel(width = 10,
                                 highchartOutput("WOS_ZINUS",height = "480px")
                               
                                    )
                                ),
                      hr(),
  
   ########################## Pie chart for Category distribution ##################
                      h3("Category Distribution"),
                      br(),
                      fluidRow(
                        splitLayout(cellWidths = c("50%", "50%"), highchartOutput("ZinusInv_CategoryPie"), highchartOutput("ZinusSales_CategoryPie"))
                      ),
                      br(),
                      hr(),
                    
   ########################## Pie Chart for class distribution ############################                   
                      h3("Class Distribution"),
                      br(),
                      fluidRow(
                        splitLayout(cellWidths = c("50%", "50%"), highchartOutput("ZinusInv_ClassPie"), highchartOutput("ZinusSales_ClassPie"))
                      ),
                      br(),
                      hr(),
   
                      hr(),br()
   
                   ) ###### end of fluidPage
                   
                 ), ######## end of tab 
                 tabPanel("Wayfair CG",
                    fluidPage(
                      h4(Update_Date),
                      br(),
                      h3("Trend of Wayfair CG Inventory"),
                      br(),
                      
                      sidebarLayout(
                        sidebarPanel(width = 2,
                                     radioButtons("categoryInput2", label = h4("Choose Category:"),
                                                  choices = list("All Cateogories"="Total","Foam Mattress" = "FM", "Spring Mattress" = "SM", "Frame" = "FR", "Box Spring" = "BX", "Platform Bed" = "PB", "Others" = "OT")
                                     ),
                                     selectInput("classInput2", label = h4("Choose Class:"),
                                                 choices = list("Total","A","B","C","Discontinued")
                                     )),
                        mainPanel(width = 10,
                                  highchartOutput("WOS_WFCG",height = "480px")
                                  
                        )
                      ),
                      hr(),
                      
                      
########################## Pie chart for Category distribution ##################
                      h3("Category Distribution"),
                      br(),   
                      fluidRow(
                        splitLayout(cellWidths = c("50%", "50%"), highchartOutput("WFInv_CategoryPie"), highchartOutput("WFSales_CategoryPie"))
                      ),
                      br(),
                      hr(),
########################## Pie Chart for class distribution ############################                   
                      h3("Class Distribution"),
                      br(),
                      fluidRow(
                        splitLayout(cellWidths = c("50%", "50%"), highchartOutput("WFInv_ClassPie"), highchartOutput("WFSales_ClassPie"))
                      ),
                      br(),
                      hr(),
                      hr(),br()
                      
                            
                          ) ######## end of fluidPage
                 
    )######### end of Tab
)#### end of UI
  
 
############# build Server ################# 
server <-function(input,output){
  output$WOS_ZINUS <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(Zinus_Function()$date, Zinus_Function()$`Total Import Plan`, name = "Total Import Plan",yAxis=0) %>%
      hc_add_series_times_values(Zinus_Function()$date, Zinus_Function()$`Total Ending Inventory`, name = "Total Ending Inventory",yAxis=0,id="ZinusWeekofSupply") %>%
      hc_add_series_times_values(Zinus_Function()$date, Zinus_Function()$WOS, name = "WOS",yAxis=1) %>%
      hc_add_series_times_values(Zinus_Function()$date, Zinus_Function()$`Total Forecast QTY`, name = "Actual Sales/Forecast",yAxis=0,color="#f45b5b") %>%
      
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      #hc_add_series(data_flags, mapping=hcaes(x=Week_Date), type = "flags", onSeries = "ZinusWeekofSupply", name = "Last Week", color="#f45b5b") %>%
    
      hc_add_theme(hc_theme_gridlight())
    
  })
  
  output$WOS_WFCG <- renderHighchart({
    highchart(type = "stock") %>%
      hc_tooltip(valueDecimals = 2) %>%
      hc_yAxis_multiples(
        list(opposite = FALSE),
        list()
      ) %>%  
      hc_add_series_times_values(WF_Function()$date, WF_Function()$`Direct Import`, name = "Direct Import",yAxis=0) %>%
      hc_add_series_times_values(WF_Function()$date, WF_Function()$`CG Ending Inventory`, name = "CG Ending Inventory",yAxis=0) %>%
      hc_add_series_times_values(WF_Function()$date, WF_Function()$WOS, name = "WOS",yAxis=1) %>%
      
      hc_legend(align = 'right', verticalAlign = 'middle', layout = 'vertical', enabled = TRUE) %>%
      hc_rangeSelector(selected = 4) %>%
      #hc_add_series(data_flags, hcaes(x = date), type = "flags", onSeries = "WFWOS", name = "Current Week", color="#f45b5b") %>%
     
      hc_add_theme(hc_theme_gridlight())
    
  })
 
   Zinus_Function <- reactive({
     DF1 <- data.frame(matrix(ncol=0,nrow=(ncol(ZINUS_WOS)-2)))
     for(i in 3:ncol(ZINUS_WOS)){
       DF1$date[i-2]<-colnames(ZINUS_WOS)[i]
    
     }
     DF1$date <- as.Date(DF1$date,format = "%m/%d/%Y")
     Selected_Data <-ZINUS_WOS[which(ZINUS_WOS$`Category/Class`==paste(input$categoryInput,"_",input$classInput,sep = "")),]
     
     DF1$`Total Import Plan` <- as.numeric(Selected_Data[which(Selected_Data$List=="Total Import Plan"),3:ncol(Selected_Data)])
     DF1$`Total Forecast QTY` <- as.numeric(Selected_Data[which(Selected_Data$List=="Total Forecast Qty"),3:ncol(Selected_Data)])
     DF1$`Total Ending Inventory` <- as.numeric(Selected_Data[which(Selected_Data$List=="Total Ending Inventory"),3:ncol(Selected_Data)])
     DF1$`WOS` <- as.numeric(Selected_Data[which(Selected_Data$List=="WOS"),3:ncol(Selected_Data)])
     
     return(DF1) 
  })
   
   WF_Function <- reactive({
     DF2 <- data.frame(matrix(ncol=0,nrow=(ncol(WF_WOS)-2)))
     for(i in 3:ncol(WF_WOS)){
       DF2$date[i-2]<-colnames(WF_WOS)[i]
       
     }
     DF2$date <- as.Date(DF2$date,format = "%m/%d/%Y")
     Selected_Data <-WF_WOS[which(WF_WOS$`Category/Class`==paste(input$categoryInput2,"_",input$classInput2,sep = "")),]
     
     Append_Data <-t(Selected_Data)
     DF2$`Direct Import` <- as.numeric(Append_Data[c(3:nrow(Append_Data)),1])
     DF2$`CG Ending Inventory` <- as.numeric(Append_Data[c(3:nrow(Append_Data)),3])
     DF2$`WOS` <- as.numeric(Append_Data[c(3:nrow(Append_Data)),4])
     
     return(DF2) 
   })
   
   output$ZinusInv_ClassPie <-renderHighchart({
     ################ Set Color ########################
     colors <- revalue(ZinusInv_Class$Class, c("A" = "#32CD32","B" = "#ADFF2F","C" = "#98FB98","NA" = "#FF1493", "NB" = "#FF69B4", "NC" = "#FFB6C1",
                                               "New" = "#E44A12", "Discontinued" = "#A9A9A9"))
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusInv_Class$Class, values = ZinusInv_Class$`Last Week`,colors = colors)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "INVENTORY DISTRIBUTION BY CLASS (LAST WEEK)")
     #ggplot(data = Zinus_ClassData(), aes(x = "", y = Percentage, fill = Class )) + 
      # geom_bar(stat = "identity", position = position_fill()) +
       #geom_text(aes(label = Label), position = position_fill(vjust = 0.5)) +
       #coord_polar(theta = "y") +
       #facet_wrap(~List )  +
       #theme(
        # axis.title.x = element_blank(),
         #axis.title.y = element_blank()) + theme(legend.position='bottom') + guides(fill=guide_legend(nrow=2,byrow=TRUE))
   })
   output$ZinusSales_ClassPie <-renderHighchart({
     colors <- revalue(ZinusSales_Class$Class, c("A" = "#32CD32","B" = "#ADFF2F","C" = "#98FB98","NA" = "#FF1493", "NB" = "#FF69B4", "NC" = "#FFB6C1",
                                                  "New" = "#E44A12", "Discontinued" = "#A9A9A9" ))
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusSales_Class$Class, values = ZinusSales_Class$`Last Week`,colors = colors)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "SALES DISTRIBUTION BY CLASS (LAST WEEK)")
    
   })
   output$ZinusInv_CategoryPie<-renderHighchart({
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusInv_Category$Category, values = ZinusInv_Category$`Last Week`)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "INVENTORY DISTRIBUTION BY CATEGORY (LAST WEEK)")
   }) 
   
   output$ZinusSales_CategoryPie <-renderHighchart({
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = ZinusSales_Category$Category, values = ZinusSales_Category$`Last Week`)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "SALES DISTRIBUTION BY CATEGORY (LAST WEEK)")
     
   })
   
   
   
   output$WFInv_ClassPie <-renderHighchart({
     colors <- revalue(WFInv_Class$Class, c( "A" = "#32CD32","B" = "#ADFF2F","C" = "#98FB98", "Discontinued" = "#A9A9A9"))
                                          
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = WFInv_Class$Class, values = WFInv_Class$`Last Week`,colors = colors)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "WAYFAIR CG INVENTORY DISTRIBUTION BY CLASS (LAST WEEK)")
   })
   
   output$WFSales_ClassPie <-renderHighchart({
     colors <- revalue(WFSales_Class$Class, c( "A" = "#32CD32","B" = "#ADFF2F","C" = "#98FB98", "Discontinued" = "#A9A9A9"))
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = WFSales_Class$Class, values = WFSales_Class$`Last Week`,colors = colors)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "WAYFAIR CG SALES DISTRIBUTION BY CLASS (LAST WEEK)")
     
   })
   
   output$WFInv_CategoryPie<-renderHighchart({
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = WFInv_Category$Category, values = WFInv_Category$`Last Week`)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "WAYFAIR CG INVENTORY DISTRIBUTION BY CATEGORY (LAST WEEK)")
   }) 
   
   output$WFSales_CategoryPie <-renderHighchart({
     highchart() %>% 
       hc_chart(type = "pie") %>% 
       hc_add_series_labels_values(labels = WFSales_Category$Category, values = WFSales_Category$`Last Week`)%>% 
       hc_tooltip(pointFormat = paste('${point.y} <br/><b>{point.percentage:.1f}%</b>')) %>%
       hc_title(text = "WAYFAIR CG SALES DISTRIBUTION BY CATEGORY (LAST WEEK)")
     
   })
   

  
}

############ deploy app ################
shinyApp(ui,server)