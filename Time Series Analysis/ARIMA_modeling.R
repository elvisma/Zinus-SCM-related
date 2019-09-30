library(xts)
library(astsa)

### utilize built-in AirPassengers soi, DJIA DataSET
help(AirPassengers)
help(djia$Close)
help(soi)

plot(AirPassengers)
plot(djia$Close)
plot(soi)


# Generate and plot white noise
WN <- arima.sim(model = list(order = c(0, 0, 0)), n = 200)
plot(WN)

# Generate and plot an MA(1) with parameter .9 by filtering the noise
MA <- arima.sim(model = list(order = c(0, 0, 1), ma = .9), n = 200)  
plot(MA)

# Generate and plot an AR(1) with parameters 1.5 and -.75
### p=2 --> ar is a vector, not just one numeric value
AR <- arima.sim(model = list(order = c(2, 0, 0), ar = c(1.5, -0.75)), n = 200) 
plot(AR)


#### AR and MA modeling
#simulate model 
x <-arima.sim(model=list(order=c(2,0,0),ar=c(1.5,-0.75)),n=200)

#plotx
plot(x)

plot(acf2(x))
plot(pacf(x))

## fit an AR(2) to the data and examine the t-table
x_fit <- astsa::sarima(x,p=2,d=0,q=0)
x_fit$ttable

#### ARMA model choise (smallest AIC/BIC)
###From the P/ACF pair, apparent that correlation are small and returns are nealy noise
### BUT it could be both ACF and PACF are tailing off
###in that case, ARMA(1,1) is suggested

# Calculate approximate oil returns
oil_returns <-diff(log(oil))

# Plot oil_returns. Notice the outliers.
plot(oil_returns)

# Plot the P/ACF pair for oil_returns
acf2(oil_returns)
pacf(oil_returns)
# Assuming both P/ACF are tailing, fit a model to oil_returns
sarima(oil_returns,1,0,1)



## if differenced data seems ARMA, then data seems like ARIMA
# Plot the sample P/ACF pair of the differenced data 
acf2(diff(globtemp))

pacf(diff(globtemp))

# Fit an ARIMA(1,1,1) model to globtemp
sarima(globtemp, 1,1,1)

# Fit an ARIMA(0,1,2) model to globtemp. Which model is better?
sarima(globtemp,0,1,2)


# Plot P/ACF pair of differenced data 
acf2(diff(x))
pacf(diff(x))
# Fit model - check t-table and diagnostics
sarima(x,1,1,0)


# Forecast the data 20 time periods ahead
sarima.for(x, n.ahead = 20, p = 1, d = 1, q = 0) 
lines(y)  





####### Pure seasonal model

# Plot sample P/ACF to lag 60 and compare to the true values

acf2(x, max.lag = 60)

# Fit the seasonal model to x
sarima(x, p = 0, d = 0, q = 0, P = 1, D = 0, Q = 1, S = 12)   #### the format is important  pdq--> non-seasonal components PDQS --> seasonal components


####### Mixed seasonal model

# Plot unemp 
plot(unemp)

# Difference your data and plot
d_unemp <- diff(unemp)
plot(d_unemp)

# Plot seasonal differenced diff_unemp
dd_unemp <- diff(d_unemp, lag = 12)   
plot(dd_unemp)


# Plot P/ACF pair of the fully differenced data to lag 60
dd_unemp <- diff(diff(unemp), lag = 12)

#### ACF analysis is for dd_unemp
acf2(dd_unemp,max.lag=60)
pacf(dd_unemp,max.lag=60)

# Fit an appropriate model
#### Fit is for unemp
sarima(unemp, p = 2, d = 1, q = 0, P = 0, D = 1, Q = 1, S = 12)


# Plot P/ACF to lag 60 of differenced data
d_birth <- diff(birth)
acf2(d_birth,max.lag=60)
pacf(d_birth, max.lag=60)

# Plot P/ACF to lag 60 of seasonal differenced data
dd_birth <- diff(d_birth, lag = 12)
acf2(dd_birth, max.lag=60)
pacf(dd_birth, max.lag=60)
# Fit SARIMA(0,1,1)x(0,1,1)_12. What happens?
sarima(birth,0,1,1,0,1,1,12)

# Add AR term and conclude
sarima(birth,1,1,1,0,1,1,12)

# Fit the chicken model again and check diagnostics
sarima(chicken,2,1,0,1,0,0,12)

# Forecast the chicken data 5 years into the future
sarima.for(chicken, n.ahead=60, 2,1,0,1,0,0,12)

