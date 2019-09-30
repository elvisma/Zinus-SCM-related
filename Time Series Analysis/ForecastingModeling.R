library(ggplot2)
library(fpp2)
library(forecast)

autoplot(oil)
train<-window(oil, end=2000)
#### if p value is > 0.05 residual is white noise, normally distribution
train%>%naive()%>%checkresiduals()

### two forecasting methods

naive_fc <-naive(train,h=13)
snaive_fc <-snaive(train, h=13)
mean_fc <-meanf(train,h=13)

### use accuracy() to compute RMSE statiscs
accuracy(naive_fc, oil)['Test set',c('RMSE','MAPE')]
accuracy(snaive_fc, oil)['Test set',c('RMSE','MAPE')]
accuracy(mean_fc, oil)['Test set',c('RMSE','MAPE')]


##### WHAT IS A GOOD FORECASTING MODEL?
#####  low RMSE on a test set
#####  white noise residual

### simple exponential smoothing
### ses() function
oildata <-window(oil, start=1996)
fc_ses <-ses(oildata, h=5)
summary(fc_ses)

autoplot(fc_ses)+ylab("oil (million of tonnes)")+xlab("Year")
### add one step forecast of the observations
autoplot(fc_ses)+autolayer(fitted(fc_ses))

checkresiduals(fc_ses)  # p value >0.05 is white noise


#### exponential smoothing with trend
fc_holt <-holt(oildata, h=5)
summary(fc_holt)
### a bad example for trending
autoplot(fc_holt)
autoplot(fc_holt)+autolayer(fitted(fc_holt))
checkresiduals(fc_holt)  # P value <0.05 not white noise

#### exponential smoothing with trend and seasonality
autoplot(gas)  ###  frequency 52--> weekly data
head(gas)
fc_hw <-hw(gas, seasonal = 'additive',h=7)
autoplot(fc_hw)

##### exponential smoothing --> ETS model choose the function for you ###
fitoil <-ets(oilprice)
fitgas <-ets(gas)
fc_ets_gas <-forecast(fitgas)
autoplot(fc_ets_gas)

##pipe data

## ets model will fail when long term cyclic time series
autoplot(oilprice)
oilprice%>%
  ets()%>%
  forecast(h=10)%>%
  autoplot
##### compare ETS model and seasonal naive model
## define function
fets <-function(y,h){
  forecast(ets(y),h=h)
}

e1 <-tsCV(oil, fets, h=4)
e2 <-tsCV(oil, snaive, h=4)

# compute MSE of resulting errors 
mean(e1^2, na.rm=TRUE)
mean(e2^2, na.rm = TRUE)

# complex is not always better!!
bestmse <-mean(e2^2, na.rm=TRUE)

##### data transformation
### get the lambda value
forecast::BoxCox.lambda(oil)
forecast::BoxCox.lambda(gas)
### get the model 
gas_ets <-ets(gas, lambda=0.0826)
autoplot(gas_ets)
f_transformed_gas <-forecast(gas_ets,h=60)
autoplot(f_transformed_gas)+autolayer(fitted(f_transformed_gas))

##### arima model



##### Advanced forecasting methods


### dynamic harnomic regression
#### important: particularly useful for short period seasonality, like weekly data, daily data, and sub-daily data

# fourier(x, K, h = NULL)
## example
fit <- auto.arima(cafe, xreg = fourier(cafe, K = 6),
                    seasonal = FALSE, lambda = 0)
fit %>%
  forecast(xreg = fourier(cafe, K = 6, h = 24)) %>%
  autoplot() + ylim(1.6, 5.1)

##
# Set up harmonic regressors of order 13, which has been chosen to minimize the AICc
harmonics <- fourier(gasoline, K = 13)

# Fit regression model with ARIMA errors
# here the ARIMA error is not seasonal, because seasonality is handled by regressors
fit <- auto.arima(gasoline, xreg = harmonics, seasonal = FALSE)

# Forecasts next 3 years
newharmonics <- fourier(gasoline, K = 13, h = 156)
fc <- forecast(fit, xreg = newharmonics)

# Plot forecasts fc
autoplot(fc)


####### harmonic regression for multiple seasonality 

### assume taylor contains half-hourly electricity demand in England over a few months
### 2 types of seasonality periods are 1:  48 (daily), 2: 7*48 =336(weekly)
### auto.arima() will take a long time to fit a long time series
### instead you fit a standard regression model with Fourier terms to capture seasonalities : tslm()
# Fit a harmonic regression using order 10 for each type of seasonality
fit <- tslm(taylor ~ fourier(taylor, K = c(10, 10)))

# Forecast 20 working days ahead
fc <- forecast(fit, newdata = data.frame(fourier(taylor, K = c(10, 10), h = 20 * 48)))

# Plot the forecasts
autoplot(fc)

# Check the residuals of fit
checkresiduals(fit)

####### Forecasting call bookings 
###
# Plot the calls data
autoplot(calls)

# Set up the xreg matrix:  using order 10 for daily seasonality and 0 for weekly seasonality.
xreg <- fourier(calls, K = c(10, 0))

# Fit a dynamic regression model
fit <- auto.arima(calls, xreg = xreg, seasonal = FALSE, stationary = TRUE)

# Check the residuals of the fitted model.
checkresiduals(fit)

# Plot forecast for 10 working days ahead
fc <- forecast(fit, xreg = fourier(calls, c(10, 0), h = 12*24*10))
autoplot(fc)

