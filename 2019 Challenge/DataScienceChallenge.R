#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#Get packages
library(zoo)
library(DataExplorer)
library(openxlsx)
library(tidyverse)
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#Load and Explore Data

Adds=read_csv("C:/Users/Greco/Desktop/DataAnalyst_Ecom_data_addsToCart.csv")
Counts=read_csv("C:/Users/Greco/Desktop/DataAnalyst_Ecom_data_sessionCounts.csv")

str(Counts)
str(Adds)
DataExplorer::create_report(Adds)
DataExplorer::create_report(Counts)
#+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#Clean and Prep Data

#Format Dates
Counts$dim_date<- as.Date(Counts$dim_date, "%m/%d/%Y")

#Create Year-Month Column
Counts$YearMonth <- as.yearmon(Counts$dim_date)
Adds$YearMonth <- as.yearmon(paste(Adds$dim_year, Adds$dim_month), "%Y %m")

#Drop redundant columns 
Adds<-select(Adds,-c(1,2))


#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#Aggregate Month * Device for Worksheet 1

#Select Metrics for Worksheet #1
df1<-select(Counts, sessions, transactions, QTY)
df1<-df1 %>% mutate(ECR= transactions / sessions)

#Aggregate by MEAN
aggMeans<-aggregate(df1,
                    by=list(Counts$dim_deviceCategory, Counts$YearMonth),
                    FUN = mean)      
#Cleanup  
aggMeans<-aggMeans%>% mutate(ECR= transactions / sessions)
aggMeans<-aggMeans%>% rename(Device = Group.1)
aggMeans<-aggMeans%>% rename(YearMonth = Group.2)
aggMeans<-aggMeans%>% rename(Sessions = sessions)
aggMeans<-aggMeans%>% rename(Transactions = transactions)
aggMeans<-aggMeans%>% rename(Date = YearMonth)


#Aggregate by SUM
aggSUM<-aggregate(df1,
                  by=list(Counts$dim_deviceCategory, Counts$YearMonth),
                  FUN = sum) 

#Cleanup  
aggSUM<-aggSUM%>% mutate(ECR= transactions / sessions)
aggSUM<-aggSUM%>% rename(Device = Group.1)
aggSUM<-aggSUM%>% rename(YearMonth = Group.2)
aggSUM<-aggSUM%>% rename(Sessions = sessions)
aggSUM<-aggSUM%>% rename(Transactions = transactions)
aggSUM<-aggSUM%>% rename(Date = YearMonth)

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#Month over month comparison for worksheet 2

#Summarize metrics by month
monthlycomp <- Counts %>%
  group_by(YearMonth) %>%
  summarize(
    MonthlySessions = sum(sessions),
    MonthlyTransactions = sum(transactions),
    MonthlyQTY = sum(QTY) ) %>%
  arrange(YearMonth)

#add monthly ECR
monthlycomp<-monthlycomp%>% mutate(MonthlyECR= MonthlyTransactions / MonthlySessions)

#Join tables to get addstocart
monthlycomp <- inner_join(monthlycomp, Adds,by="YearMonth")
monthlycomp<-monthlycomp%>% rename(MonthlyAddsToCart = addsToCart)

#limit last two months
TwoMonthReport<-monthlycomp[-c(1:10),]

#Relative differences

TwoMonthReport1 <- (TwoMonthReport[2,] - TwoMonthReport[1,])/TwoMonthReport[1,]
TwoMonthReport1$YearMonth = as.character(TwoMonthReport1$YearMonth)
TwoMonthReport1[1,1] = 'Relative'

TwoMonthReport2 <- (TwoMonthReport[2,] - TwoMonthReport[1,])
TwoMonthReport2$YearMonth = as.character(TwoMonthReport2$YearMonth)
TwoMonthReport2[1,1] = 'Absolute'



TwoMonthReport$YearMonth = as.character(TwoMonthReport$YearMonth)


TwoMonthReport <- TwoMonthReport %>% 
  rbind(TwoMonthReport1) %>% 
  rbind(TwoMonthReport2) 
  

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#CreateXLSX files with two worksheets with openxlsx
wb<-createWorkbook("GrecoWB")
#add worksheets
addWorksheet(wb, "AggregateMeans", gridLines = FALSE)
addWorksheet(wb, "MOM", gridLines = FALSE)

#write data to worksheets
writeData(wb, sheet = 1, aggMeans, rowNames=FALSE)
writeData(wb, sheet = 2, TwoMonthReport, rowNames=FALSE)

## create and add a style to the column headers
headerStyle <- createStyle(
  fontSize = 10, fontColour = "#FFFFFF", halign = "center",
  fgFill = "#4F81BD", border = "TopBottom", borderColour = "#4F81BD"
)
addStyle(wb, sheet = 1, headerStyle, rows = 1, cols = 1:6, gridExpand = TRUE)
## style for body
bodyStyle <- createStyle(border = "TopBottom", borderColour = "#4F81BD")
addStyle(wb, sheet = 1, bodyStyle, rows = 2:37, cols = 1:6, gridExpand = TRUE)
setColWidths(wb, 1, cols = 1, widths = 20) 

## create and add a style to the column headers
headerStyle <- createStyle(
  fontSize = 10, fontColour = "#FFFFFF", halign = "center",
  fgFill = "#4F81BD", border = "TopBottom", borderColour = "#4F81BD"
)
addStyle(wb, sheet = 2, headerStyle, rows = 1, cols = 1:6, gridExpand = TRUE)
addStyle(wb, sheet = 2, rows = 4, cols = 2:6, style = createStyle(numFmt= 'PERCENTAGE'))

## style for body
bodyStyle <- createStyle(border = "TopBottom", borderColour = "#4F81BD")
addStyle(wb, sheet = 2, bodyStyle, rows = 2:5, cols = 1:6, gridExpand = TRUE)
setColWidths(wb, 2, cols = 1:6, widths = 20) 



saveWorkbook(wb, "GrecoWB.xlsx", overwrite = TRUE)   

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#PLOTS

#Quantity by Device Bar Chart
ggplot(aggSUM, aes(x=Device, y=QTY)) + 
  geom_bar(stat="identity", width=.5, fill="tomato3") + 
  labs(title="Quantity by Device", 
       subtitle="Total Sum for Year", 
       caption="source: Google") + 
  theme(axis.text.x = element_text(angle=65, vjust=0.6))

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#Quantity by Browser Bar Chart:
#Sum QTY by browser
Count2<-Counts %>%                                        
  group_by(dim_browser) %>%                         
  summarise_at(vars(QTY),              
               list(name = sum))     
# Rename
Count2<-Count2%>% rename(Quantity = name)
Count2<-Count2%>% rename(Browser = dim_browser)
#Descending Order
Count2 <- Count2 %>%                                     
  arrange(desc(Quantity)) 
#top 5
Count2<-head(Count2, 5)
#plot
ggplot(Count2, aes(x=Browser, y=Quantity)) + 
  geom_bar(stat="identity", width=.5, fill="Blue") + 
  labs(title="Top 5 Most Freqent Browsers", 
       subtitle="Quantity by Browser", 
       caption="source: Google") + 
  theme(axis.text.x = element_text(angle=65, vjust=0.6))

#++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

#prep
#Summarize metrics 
Counts1 <- Counts %>%
  group_by(dim_date) %>%
  summarize(
    DailySessions = sum(sessions),
    DailyTransactions = sum(transactions),
    DailyQTY = sum(QTY) ) %>%
  arrange(dim_date)

#add monthly ECR
Counts1<-Counts1%>% mutate(DailyECR= DailyTransactions / DailySessions)
#rename
Counts1<-Counts1%>% rename(Month = dim_date)


# plot
ggplot(Counts1, aes(x=Month)) + 
  geom_line(aes(y=DailyECR)) + 
  labs(title="Daily Time Series", 
       subtitle="Ecommerce Conversion Rate", 
       caption="Google", 
       y="ECR %") +  # title and caption
  scale_x_date(date_breaks = "1 month", date_minor_breaks = "1 week",
               date_labels = "%m")
  theme(axis.text.x = element_text(angle = 90, vjust=0.5),  # rotate x axis text
        panel.grid.minor = element_blank())  # turn off minor grid



