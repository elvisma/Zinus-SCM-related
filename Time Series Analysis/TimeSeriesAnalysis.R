library(xts)
library(zoo)
library(lubridate)

## Regression Model
# create holiday/promotinoal effect variable
v.date<-as.Date(c("2014-02-09","2015-02-08","2016-02-07"))   #need to add c(), otherwise the total length would be 1 instead of 3
valentine <- as.xts(rep(1,3),order.by=v.date)

dates_train <- seq(as.Date("2014-01-19"),length=154, by="weeks")
valentine <-merge(valentine,dates_train,fill=0)



###### Case Study 1: Flight Data #####
### get a feel of data
str(flights)
head(flights, n=5)
class(flights$date) ## check the format of the date column
### if it is character, change it to date
flights$date <- as.Date(flights$date)

# Load the xts package
library(xts)

# Convert date column to a time-based class
flights$date <- as.Date(flights$date)

# Convert flights to an xts object using as.xts
flights_xts <- as.xts(flights[, -5], order.by = flights$date)

# Check the class of flights_xts
class(flights_xts)

# Examine the first five lines of flights_xts
head(flights_xts, n=5)


# Use plot.xts() to view total monthly flights into BOS over time
plot.xts(flights_xts$total_flights)

# Use plot.xts() to view monthly delayed flights into BOS over time
plot.xts(flights_xts$delay_flights)

# Use plot.zoo() to view all four columns of data in their own panels
plot.zoo(flights_xts, plot.type = "multiple", ylab = labels)

# Use plot.zoo() to view all four columns of data in one panel
plot.zoo(flights_xts, plot.type = "single", lty = lty)
legend("right", lty = lty, legend = labels)


#### write & read xts objects
# Save your xts object to rds file using saveRDS
saveRDS(object = flights_xts, file = "flights_xts.rds")

# Read your flights_xts data from the rds file
flights_xts2 <- readRDS("flights_xts.rds")

# Check the class of your new flights_xts2 object
class(flights_xts2)

# Examine the first five rows of your new flights_xts2 object
head(flights_xts2,n=5)

# Export your xts object to a csv file using write.zoo
write.zoo(flights_xts, file = "flights_xts.csv", sep = ",")

# Open your saved object using read.zoo
flights2 <- read.zoo("flights_xts.csv", sep = ",", FUN = as.Date, header = TRUE, index.column = 1)

# Encode your new object back into xts
flights_xts2 <- as.xts(flights2)

# Examine the first five rows of your new flights_xts2 object
head(flights_xts2,n=5)


#### Case Study 2: Weather data ####
install.packages("weatherData")

# Confirm that the date column in each object is a time-based class
class(temps_1)
class(temps_2)

# Encode your two temperature data frames as xts objects
temps_1_xts <- as.xts(temps_1[, -4], order.by = temps_1$date)
temps_2_xts <- as.xts(temps_2[, -4], order.by = temps_2$date)

# View the first few lines of each new xts object to confirm they are properly formatted
head(temps_1_xts)
head(temps_2_xts)

# Use rbind to merge your new xts objects
temps_xts <- rbind(temps_1_xts,temps_2_xts)

# View data for the first 3 days of the last month of the first year in temps_xts
first(last(first(temps_xts, "1 year"), "1 month"), "3 days")

# Identify the periodicity of temps_xts
periodicity(temps_xts)

# Generate a plot of mean Boston temperature for the duration of your data
plot.xts(temps_xts$mean)

# Generate a plot of mean Boston temperature from November 2010 through April 2011
plot.xts(temps_xts$mean["2010-11/2011-04"])

# Use plot.zoo to generate a single plot showing mean, max, and min temperatures during the same period 
plot.zoo(temps_xts["2010-11/2011-04"], plot.type = "single", lty = lty)
legend('right',lty = lty,legend=labels)

### Workflow for merging
##1. Encoding all time series objects to xts
data_1_xts <- as.xts(as.numeric(data_1), order.by=index)

##2. Examine and adjust periodicity 
periodicity(data_1_xts)
to.period(data_1_xts, period = "years")
##3. Merge xts objects
merge_data <- merge(data_1_xts, data_2_xts)


###CASE 3: economics
## handling missingness
na.locf(data)  ### fill NAs with Last Observation carried forward
na.locf(data, fromLast=TRUE) ### fill NAs with next observation carried backward
na.approx(data)   ### fill NAs with linear interpolation


### generating moving indicators:  lagging & differencing
# Generate monthly difference in unemployment
unemployment$us_monthlydiff <- diff(unemployment$us, lag = 1, differences = 1)  

# Generate yearly difference in unemployment
unemployment$us_yearlydiff <- diff(unemployment$us, lag = 12, differences = 1)

# Plot US unemployment and yearly difference
par(mfrow = c(2,1))
plot.xts(unemployment$us)
plot.xts(unemployment$us_yearlydiff, type = "h")

# rolling functions --> rolling indicators

## discrete windows  (SPLIT LAPPLY RBIND pattern)
# split the data according to period
unemployment_yrs <-split(unemployment, f="years")

# Apply function within period
unemployment_yrs <-lapply(unemployment_yrs, cummax)

# Bind new data into xts object
unemployment-ytd <- do.call(rbind, unemployment_yrs)


## rolling windows
# rollapply() applies a function to a rolling window
unemployment_avg <- rollapply(unemployment, width=12, FUN=mean)

## example
# Add a quarterly difference in gdp
gdp$quarterly_diff <- diff(gdp$gdp, lag = 1, differences = 1)

# Split gdp$quarterly_diff into years
gdpchange_years <- split(gdp$quarterly_diff, f = "years")

# Use lapply to calculate the cumsum each year
gdpchange_ytd <- lapply(gdpchange_years, FUN = cumsum)

# Use do.call to rbind the results
gdpchange_xts <- do.call(rbind, gdpchange_ytd)

# Plot cumulative year-to-date change in GDP
plot.xts(gdpchange_xts, type = "h")

### xts/zoo command from LAG, DIFF to ROLLING INDICATOR

# Add a one-year lag of MA unemployment
unemployment$ma_yearlag <- lag(unemployment$ma,k=12)

# Add a six-month difference of MA unemployment
unemployment$ma_sixmonthdiff <- diff(unemployment$ma, lag=6, difference=1)

# Add a six-month rolling average of MA unemployment
unemployment$ma_sixmonthavg <- rollapply(unemployment$ma, width=6, FUN=mean)

# Add a yearly rolling maximum of MA unemployment
unemployment$ma_yearmax <- rollapply(unemployment$ma, width=12, FUN=max)

# View the last year of unemployment data
tail(unemployment, n=12)

#### combination of endpoints() & period.apply() 
# Generate a new variable coding for red sox wins
redsox_xts$win_loss <- ifelse(redsox_xts$boston_score > redsox_xts$opponent_score, 1, 0)

# Identify the date of the last game each season
close <- endpoints(redsox_xts, on = "years")

# Calculate average win/loss record at the end of each season
period.apply(redsox_xts[, "win_loss"], INDEX=close, FUN=mean)


##### generate closing average
endpoints() + period.apply(data, INDEX=endpoints(), FUN)

##### generate cumulative average
split() + lapply() + do.call(rbind,list)

##### generate rolling average(moving average)
rollapply(data, width=10, FUN)

##### indexing commands
.indexwday() 
sunday=0 
monday=1  
tuesday=2 ### and so on

sunday_games<-which(.indexwday(sports)==0)

# Extract the day of the week of each observation
weekday <- .indexwday(sports)
head(weekday)

# Generate an index of weekend dates
weekend <- which(.indexwday(sports) == 6 | .indexwday(sports) == 0)

# Subset only weekend games
weekend_games <- sports[weekend]
head(weekend_games)