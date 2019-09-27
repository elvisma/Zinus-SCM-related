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
library(plotly)

#####################################################
Updated_Week <-"12/09/2018"

#setwd("C:/Users/Elvis Ma/Desktop/Visualization tools/Overstock dashboard")

Overstock_Report <- openxlsx::read.xlsx("Overstock Rate Project.xlsm",sheet="MasterData")
Split_Data1 <- Overstock_Report[,c(1:6)]
Group_Data1 <- group_by(Split_Data1,Category,Inventory.Status)
Summ_Data1 <- summarise(Group_Data1,SKU.Qty=sum(SKU.Qty),Percentage.SKU=sum(Percentage.SKU))
Summ_Data1 <- as.data.frame(append(Summ_Data1,list(ABC="Total"),after = 1),stringsAsFactors = FALSE)
Summ_Data1$`Percentage.SKU.per.Class` <-Summ_Data1$Percentage.SKU
Split_Data1 <-rbind(Split_Data1,Summ_Data1)
Split_Data1$Combine <-paste(Split_Data1$Category,Split_Data1$ABC,sep=" ")


##################################################################
Split_Data2 <- Overstock_Report[c(1:28),c(8:ncol(Overstock_Report))]
Split_Data2$Combine <-paste(Split_Data2$Category,Split_Data2$ABC,sep=" ")

############ Buiding UI #############
ui<- navbarPage("Overstock Report", theme = shinytheme("flatly"),
                tabPanel("Overstock SKUs",
                         sidebarPanel(width = 2,
                                      checkboxGroupInput("variable1","Choose Categories:",
                                                         c("Spring Mattress","Foam Mattress","Frame","Box Spring","Platform Bed"="Platform","Others","Furniture"),
                                                         selected=c("Foam Mattress","Frame","Platform Bed","Others")
                                                         ),
                                      checkboxGroupInput("variable2","Choose Your Classes:",
                                                         c("A","B","C","Total"),
                                                         selected=c("A","Total")
                           
                           
                         )),
                         mainPanel(width = 10,
                                   h3(Updated_Week),
                                   br(),
                                   tabsetPanel(
                                     tabPanel("Inventory Status QTY",plotlyOutput("Overstock_Number")),
                                     tabPanel("Inventory Status % (Whole Category)",plotlyOutput("Overstock_Percent")),
                                     tabPanel("Inventory Status % (By Class)",plotlyOutput("Class_Percent"))
                                     
                                   ),
                                   br()
                         )
                ),
                tabPanel("Surplus $",
                         sidebarPanel(width = 2,
                                      checkboxGroupInput("variable3","Choose Categories:",
                                                         c("Spring Mattress","Foam Mattress","Frame","Box Spring","Platform Bed"="Platform","Others","Furniture"),
                                                         selected=c("Foam Mattress","Frame","Platform Bed","Others")
                                      ),
                                      checkboxGroupInput("variable4","Choose Classes:",
                                                         c("A","B","C","Total"),
                                                         selected=c("A","Total")
                                      )
                         ),
                         mainPanel(width = 10,
                                   h3(Updated_Week),
                                   br(),
                                   tabsetPanel(
                                     tabPanel("Surplus QTY",plotlyOutput("SurplusQty")),
                                     tabPanel("Surplus $Amount",plotlyOutput("SurplusDollar")),
                                     tabPanel("Surplus QTY (%)",plotlyOutput("QtyPercent")),
                                     tabPanel("Surplus $Amount (%)",plotlyOutput("AmountPercent"))
                                
                                   )
                                   
                                   )
                  
                  
                )
)



################## Building Server ##################

server<-function(input,output){
  output$Overstock_Number <- renderPlotly(
    plot_ly(data1_Function(), x=~data1_Function()$Combine,y=~data1_Function()$`Overstock #`, name="Overstock SKUs",type="bar",text=data1_Function()$`Overstock #`,textposition="auto",marker=list(color="#e58b89"))%>%
    add_trace(y=~data1_Function()$`Not Overstock #`,name="Not Overstock SKUs",type='bar',text=data1_Function()$`Not Overstock #`, textposition='auto',marker=list(color="#61d15c"))%>%
    layout(xaxis=list(title="By Categories & Classes"),yaxis=list(title="SKU QTY"), barmode='group')
  )
  output$Overstock_Percent <-renderPlotly(
    plot_ly(data1_Function(), x=~data1_Function()$Combine,y=~data1_Function()$`Overstock %`, name="Overstock SKUs %",type="bar",text=paste(round(data1_Function()$`Overstock %`*100,0),'%'),textposition="auto",marker=list(color="#e58b89"))%>%
    add_trace(y=~data1_Function()$`Not Overstock %`,name="Not Overstock SKUs %",type='bar',text=paste(round(data1_Function()$`Not Overstock %`*100,0),"%"), textposition='auto',marker=list(color="#61d15c"))%>%
    layout(xaxis=list(title="By Categories & Classes"),yaxis=list(title="SKU QTY (%)"), barmode='group')
  )
  output$Class_Percent <-renderPlotly(
    plot_ly(data1_Function(), x=~data1_Function()$Combine,y=~data1_Function()$`Overstock_Class %`, name="Overstock SKUs % By Class",type="bar",text=paste(round(data1_Function()$`Overstock_Class %`*100,0),'%'),textposition="auto",marker=list(color="#e58b89"))%>%
    add_trace(y=~data1_Function()$`Not Overstock_Class %`,name="Not Overstock SKUs% By Class",type='bar',text=paste(round(data1_Function()$`Not Overstock_Class %`*100,0),"%"), textposition='auto',marker=list(color="#61d15c"))%>%
    layout(xaxis=list(title="By Categories & Classes"),yaxis=list(title="By Percentage (%)"), barmode='stack')
    
  )
  output$SurplusQty <-renderPlotly(
    plot_ly(data2_Function(), x=~data2_Function()$Combine,y=~data2_Function()$`Surplus.of.Current.Week`, name="Surplus Inventory (QTY)",type="bar",text=paste(round(data2_Function()$`Surplus.of.Current.Week`/1000,0),"K"),textposition="auto",marker=list(color="#e58b89"))%>%
      add_trace(y=~data2_Function()$`Optimial.Inventory`,name="Optimal Inventory (QTY)",type='bar',text=paste(round(data2_Function()$`Optimial.Inventory`/1000,0),"K"), textposition='auto',marker=list(color="#61d15c"))%>%
      add_trace(y=~data2_Function()$`Total.Ending.Inventory`,name="Total Ending Inventory (QTY)",type='bar',text=paste(round(data2_Function()$`Total.Ending.Inventory`/1000,0),"K"), textposition='auto',marker=list(color="#acbde5"))%>%
      layout(xaxis=list(title="By Categories & Classes"),yaxis=list(title="Inventory level (QTY)"), barmode='group')
    
  )
  output$SurplusDollar <-renderPlotly(
    plot_ly(data2_Function(), x=~data2_Function()$Combine,y=~data2_Function()$`Surplus.$`, name="Surplus Inventory ($)",type="bar",text=paste(round(data2_Function()$`Surplus.$`/1000000,2),"M"),textposition="auto",marker=list(color="#e58b89"))%>%
      add_trace(y=~data2_Function()$`Optimize.Inventory.$`,name="Optimal Inventory ($)",type='bar',text=paste(round(data2_Function()$`Optimize.Inventory.$`/1000000,2),"M"), textposition='auto',marker=list(color="#61d15c"))%>%
      add_trace(y=~data2_Function()$`Total.Inventory.$`,name="Total Ending Inventory ($)",type='bar',text=paste(round(data2_Function()$`Total.Inventory.$`/1000000,2),"M"), textposition='auto',marker=list(color="#acbde5"))%>%
      layout(xaxis=list(title="By Categories & Classes"),yaxis=list(title="Inventory level ($)"), barmode='group')
  )
  output$QtyPercent <-renderPlotly(
      plot_ly(data2_Function(), x=~data2_Function()$Combine,y=~data2_Function()$`Surplus.%`, name="Surplus % (QTY)",type="bar",text=paste(round(data2_Function()$`Surplus.%`*100,0),'%'),textposition="auto",marker=list(color="#e58b89"))%>%
      add_trace(y=~data2_Function()$`Optimial.Inventory.%`,name="Optimal Inventory % (QTY)",type='bar',text=paste(round(data2_Function()$`Optimial.Inventory.%`*100,0),"%"), textposition='auto',marker=list(color="#61d15c"))%>%
      layout(xaxis=list(title="By Categories & Classes"),yaxis=list(title="Inventory by QTY (%)"), barmode='stack')
  )
  output$AmountPercent <-renderPlotly(
      plot_ly(data2_Function(), x=~data2_Function()$Combine,y=~data2_Function()$`Surplus.Percentage.$`, name="Surplus % (QTY)",type="bar",text=paste(round(data2_Function()$`Surplus.Percentage.$`*100,0),'%'),textposition="auto",marker=list(color="#e58b89"))%>%
      add_trace(y=~data2_Function()$`Optimial.Percentage.$`,name="Optimal Inventory % (QTY)",type='bar',text=paste(round(data2_Function()$`Optimial.Percentage.$`*100,0),"%"), textposition='auto',marker=list(color="#61d15c"))%>%
      layout(xaxis=list(title="By Categories & Classes"),yaxis=list(title="Inventory by Amount (%)"), barmode='stack')
  )
  data1_Function <-reactive({
    ########
    validate(
      need(!is.null(input$variable2), "PLEASE SELECT CLASSES TO COMPARE")
    )
     if(is.null(input$variable1)&!is.null(input$variable2)){
    # validate(!is.null(input$variable2),"Please select your visualization preferences")
      
      ##########################
      group_data <-group_by(Split_Data1,ABC,Inventory.Status)  
      df1<-summarise(group_data,SKU.Qty=sum(SKU.Qty))
      df1$Percentage.SKU <-0
      
      for(i in 1:nrow(df1)){
        df1$Percentage.SKU[i] = df1$SKU.Qty[i]/sum(df1$SKU.Qty[which(df1$ABC==df1$ABC[i])])
      }
      
      null_input <-df1[which(df1$ABC%in%input$variable2),]####
      null_input$Combine <- paste("All categories",null_input$ABC)
      Total_Overstock <-null_input[which(null_input$Inventory.Status=="Overstock"),]
      Total_NotOverstock <- null_input[which(null_input$Inventory.Status=="Not Overstock"),]
      df_merge0 <-cbind(Total_Overstock[,c(5,3,4)],Total_NotOverstock[,c(3,4)])
      colnames(df_merge0)[2:5] <-c("Overstock #","Overstock %","Not Overstock #","Not Overstock %")
      df_merge0$`Overstock_Class %`<-df_merge0$`Overstock %`
      df_merge0$`Not Overstock_Class %` <-df_merge0$`Not Overstock %`
      return(df_merge0)
    }
    else
    {df2 <- Split_Data1[which(Split_Data1$Category%in%input$variable1&Split_Data1$ABC%in%input$variable2),c("Combine","Inventory.Status","SKU.Qty","Percentage.SKU","Percentage.SKU.per.Class")]
     df_Overstock <- filter(df2,df2$Inventory.Status=="Overstock")
     df_NotOverstock <- filter(df2,df2$Inventory.Status=="Not Overstock")
     df_merge <- merge(df_Overstock[,c(1,3,4,5)],df_NotOverstock[,c(1,3,4,5)],by="Combine",all=TRUE)
     colnames(df_merge)[2:7] <-c("Overstock #","Overstock %","Overstock_Class %","Not Overstock #","Not Overstock %","Not Overstock_Class %")
     return(df_merge)
     }
  })
  
  ##############################################################
  ##################### practice ###########################
  ########################################################
  data2_Function <- reactive({
    validate(
      need(!is.null(input$variable4), "PLEASE SELECT CLASSES TO COMPARE")
    )
    if(is.null(input$variable3)&!is.null(input$variable4)){
      # validate(!is.null(input$variable2),"Please select your visualization preferences")
      
      ##########################
      Data2 <-filter(Split_Data2,ABC%in%input$variable4)
      group_data2 <-group_by(Data2,ABC)  
      
      df3 <-summarize_at(group_data2,c(4:ncol(group_data2)-1),sum, na.rm=TRUE)
      df3$Combine <- paste("All categories",df3$ABC)
      df3$`Surplus.%` <- df3$Surplus.of.Current.Week/df3$Total.Ending.Inventory
      df3$`Optimial.Inventory.%` <-df3$Optimial.Inventory/df3$Total.Ending.Inventory
      df3$`Surplus.Percentage.$` <-df3$`Surplus.$`/df3$`Total.Inventory.$`
      df3$`Optimial.Percentage.$` <-df3$`Optimize.Inventory.$`/df3$`Total.Inventory.$`
      return(df3)
    }
    else
    {df4 <- Split_Data2[which(Split_Data2$Category%in%input$variable3&Split_Data2$ABC%in%input$variable4),]

    return(df4)
    
    }
  })
  
}



################### Deploy App ##################
shinyApp(ui,server)

