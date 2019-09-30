library(xts)
library(zoo)

data_vector=c(1,2,3,34,5)
plot(ts(data_vector))
plot(ts(data_vector, start = 1999, frequency = 1))


eu_stocks<-EuStockMarkets  ## R default time series database
is.ts(eu_stocks)
start(eu_stocks)
end(eu_stocks)
frequency(eu_stocks)

# Generate a simple plot of eu_stocks
plot(eu_stocks)

# Use ts.plot with eu_stocks
ts.plot(eu_stocks, col = 1:4, xlab = "Year", ylab = "Index Value", main = "Major European Stock Indices, 1991-1998")

# Add a legend to your ts.plot
legend("topleft", colnames(eu_stocks), lty = 1, col = 1:4, bty = "n")

