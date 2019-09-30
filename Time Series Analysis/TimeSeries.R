library(xts)
library(zoo)

##### xts = matrix + index

# Create the object data using 5 random numbers
data <- rnorm(5)

# Create dates as a Date class object starting from 2016-01-01
dates <- seq(as.Date("2016-01-01"), length = 5, by = "days")

# Use xts() to create smith
smith <- xts(x = data, order.by = dates)

# Create bday (1899-05-08) using a POSIXct date class object
bday <- as.POSIXct("1899-05-08")

# Create hayek and add a new attribute called born
hayek <- xts(x = data, order.by = dates, born = bday)
str(hayek)

# Extract the core data of hayek
hayek_core=coredata(hayek)

# View the class of hayek_core ==> matrix
class(hayek_core)

# Extract the index of hayek
hayek_index=index(hayek)

# View the class of hayek_index ==> date
class(hayek_index)

# Create dates
dates <- as.Date("2016-01-01") + seq(0,28,7)

# Create ts_a
ts_a <- xts(x = 1:5, order.by = dates)

# Create ts_b
ts_b <- xts(x = 1:5, order.by = as.POSIXct(dates))

# Extract the rows of ts_a using the index of ts_b
ts_a[index(ts_b)]

# Extract the rows of ts_b using the index of ts_a
ts_b[index(ts_a)]


#### import and export and converting time series
### assume austres is "ts" object
au <-as.xts(austres)  ### class(au) -> xts / zoo
am <- as.matrix(au)
am2<-as.matrix(as.xts(austres))

library(xlsx)
setwd("C:/Users/Elvis Ma/Desktop")
ts_data <- read.xlsx("timeSeriesTesting.xlsx",sheetIndex = 1)
xts_data <-xts(ts_data[,c(2,3)], order.by = ts_data[,1])

m_data <-as.matrix(xts_data)
df_data <-as.data.frame(xts_data)

### get the temporary file name
tmp <- tempfile()

###write the xts object using zoo to tmp
write.zoo(xts_data, sep=',', file=tmp)

# Read the tmp file. FUN = as.yearmon converts strings such as Jan 1749 into a proper time class
new <- read.zoo(tmp, sep = ",", FUN = as.yearmon)

# Convert sun into xts. Save this as sun_xts
new_xts <- as.xts(new)

#### Basic Date queries

# Extract all data from irreg between 8AM and 10AM
morn_2010 <- irreg['T08:00/T10:00']

# Extract the observations in morn_2010 for January 13th, 2010
morn_2010["2010-01-13"]


#### subsetting and indexing
# Subset x using the vector dates
dates=as.Date(c("2016-01-02","2016-01-04"))
x[dates]

# Subset x using dates as POSIXct
x[as.POSIXct(c("2016-06-04","2016-06-08"))]

# Replace the values in x contained in the dates vector with NA
x[dates] <- NA

# Replace all values in x for dates starting June 9, 2016 with 0
### ISO 8601
x["2016-06-09/"] <- 0
### INDEX
x[index(x)>"2016-06-09"] <- 0

# Verify that the value in x for June 11, 2016 is now indeed 0
x["2016-06-11"]

## use first() and last() to filter xts data  
# Create lastweek using the last 1 week of temps
lastweek <- last(temps, "1 week")

# Print the last 2 observations in lastweek
last(lastweek, n=2)

# Extract all but the first two days of lastweek
first(lastweek, "-2 days")

# Extract the first three days of the second week of temps
first(last(first(temps, "2 weeks"), "1 week"), "3 days")


a <-xts(c(1,1,1),order.by = as.Date(c("2015-01-24","2015-01-25","2015-01-26")))
b <-xts(2, order.by = as.Date("2015-01-24"))
# Add a and b
a+b

# Add a with the numeric value of b
a+as.numeric(b)

# Add a to b, and fill all missing rows of b with 0
a + merge(b, index(a), fill = 0)

# Add a to b and fill NAs with the last observation
a + merge(b, index(a), fill = na.locf)

### deal with missing values

### last observation carried forward
na.locf(x)
### next observatoin carried backford
na.locf(x,fromLast=TRUE)

### combine a leading and lagging time series
xts_data_lead <-lag(xts_data,k=-1)
xts_data_lag <- lag(xts_data, k=1)
xts_data_combine <- merge(xts_data_lead,xts_data,xts_data_lag)

### calculate first difference
diff_xts_data1 <-xts_data-lag(xts_data,k=1)   # method 1
diff_xts_data2<-diff(xts_data,lag=1)          # method 2

merge(head(diff_xts_data1),head(diff_xts_data2))  # make sure 2 methods are equal

wkly_diff_xts_data<-diff(xts_data,lag=52,differences=1)

# These are the same
diff(x, differences = 2)
diff(diff(x))

# Find intervals by time in xts
# Locate the weeks
endpoints(temps, on = "weeks")

# Locate every two weeks
endpoints(temps, on = "weeks", k = 2)

# Calculate the weekly endpoints
ep <- endpoints(temps, on = "weeks")

# Now calculate the weekly mean and display the results
period.apply(temps[, "Temp.Mean"], INDEX = ep, FUN = mean)

# Split temps by week
temps_weekly <- split(temps, f = "weeks")

# Create a list of weekly means, temps_avg, and print this list
temps_avg <- lapply(X = temps_weekly, FUN = mean)
temps_avg

# Use the proper combination of split, lapply and rbind
temps_1 <- do.call(rbind, lapply(split(temps, "weeks"), function(w) last(w, n = "1 day")))

# Create last_day_of_weeks using endpoints()
last_day_of_weeks <- endpoints(temps,on="weeks")

# Subset temps using last_day_of_weeks 
temps_2 <- temps[last_day_of_weeks,]

# Convert univariate series to OHLC data
# Convert usd_eur to weekly and assign to usd_eur_weekly
usd_eur_weekly <- to.period(usd_eur, period = "weeks")

# Convert usd_eur to monthly and assign to usd_eur_monthly
usd_eur_monthly <- to.period(usd_eur, period = "months")

# Convert usd_eur to yearly univariate and assign to usd_eur_yearly
usd_eur_yearly <- to.period(usd_eur, period = "years", OHLC = FALSE)

# Convert eq_mkt to quarterly OHLC
mkt_quarterly <- to.period(eq_mkt, period = "quarters")

# Convert eq_mkt to quarterly using shortcut function
mkt_quarterly2 <- to.quarterly(eq_mkt, name = "edhec_equity", indexAt = "firstof")

# Split edhec into years
edhec_years <- split(edhec , f = "years")

# Use lapply to calculate the cumsum for each year in edhec_years
edhec_ytd <- lapply(edhec_years, FUN = cumsum)

# Use do.call to rbind the results
edhec_xts <- do.call(rbind, edhec_ytd)

# Use rollapply to calculate the rolling 3 period sd of eq_mkt
eq_sd <- rollapply(eq_mkt, 3, FUN = sd,na.rm=TRUE)

#ExerciseExercise Class attributes - tclass, tzone, and tformat
# View the first three indexes of temps
index(temps)[1:3]

# Get the index class of temps
indexClass(temps)

# Get the timezone of temps
indexTZ(temps)

# Change the format of the time display
indexFormat(temps) <- "%b-%d-%Y"

# View the new format
head(temps)

# Construct times_xts with tzone set to America/Chicago
times_xts <- xts(1:10, order.by = times, tzone = 'America/Chicago')

# Change the time zone of times_xts to Asia/Hong_Kong
tzone(times_xts) <-'Asia/Hong_Kong' 

# Extract the current time zone of times_xts
tzone(times_xts)

# Construct times_xts with tzone set to America/Chicago
times_xts <- xts(1:10, order.by = times, tzone = 'America/Chicago')

# Change the time zone of times_xts to Asia/Hong_Kong
tzone(times_xts) <-'Asia/Hong_Kong' 

# Extract the current time zone of times_xts
tzone(times_xts)

# Count the months
nmonths(edhec)

# Count the quarters
nquarters(edhec)

# Count the years
nyears(edhec)


# Explore underlying units of temps in two commands: .index() and .indexwday()

.index(temps)
.indexwday(temps)

# Create an index of weekend days using which()
# .indexwday(), range from 0-6, sunday equal to 0
index <- which(.indexwday(temps) == 0 | .indexwday(temps) == 6)

# Select the index
temps[index,]

# Make z have unique timestamps
z_unique <- make.index.unique(z, eps = 1e-4)

# Remove duplicate times in z
z_dup <- make.index.unique(z, drop = TRUE)

# Round observations in z to the next hour
z_round <- align.time(z, n = 3600)

