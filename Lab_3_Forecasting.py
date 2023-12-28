import numpy as np
import pandas as pd
from tkinter.messagebox import showerror, showwarning, showinfo
import matplotlib.pyplot as plt 
import matplotlib.dates as mdates

from datetime import date, datetime
from dateutil.relativedelta import relativedelta

from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.holtwinters import SimpleExpSmoothing
from statsmodels.tsa.holtwinters import Holt
from statsmodels.tsa.holtwinters import ExponentialSmoothing 

from sklearn.metrics import mean_absolute_error
from sklearn.metrics import mean_squared_error
from sklearn.metrics import mean_squared_error
from math import sqrt

# Autoregressive (AR) forecasting
def forecast_AR(data, order1, order2, order3):
    # data (Pandas Data Series):  historical data for past 12 month
    # model parameters: order1, order2, order3
    # My_model = ARIMA(data, order=(2, 0, 0))
    # Return forecast for 12 months
    My_hist = data.to_numpy()
    My_model = ARIMA(My_hist, order=(order1, order2, order3))
    My_model_fit = My_model.fit()
    My_forecast = My_model_fit.forecast(12)
    My_forecast = My_forecast.astype(int)
    return My_forecast   

# Autoregressive Moving-Average (ARMA) forecasting 
def forecast_ARIMA(data, order1, order2, order3):
    # data (Pandas Data Series):  historical data for past 12 month
    # model parameters: order1, order2, order3
    # My_model = ARIMA(data, order=(2, 0, 1))
    # Return forecast for 12 months
    My_hist = data.to_numpy()
    My_model = ARIMA(My_hist, order=(order1, order2, order3))
    My_model_fit = My_model.fit()
    My_forecast = My_model_fit.forecast(12)
    My_forecast = My_forecast.astype(int)
    return My_forecast    

# Single Exponential Smoothing (SES) forecasting 
def forecast_SExpSmoothing(data, my_alpha):
    # data (Pandas Data Series)-- historical data for past 12 month
    # my_alpha -- smoothing level from 0 to 1    
    # my_beta -- smoothing trend from 0 to 1
    # my_gamma -- smoothing seasonal from 0 to 1
    # return forecast for next 12 months 
    My_hist = data.to_numpy()
    My_model = SimpleExpSmoothing(My_hist)
    My_model_fit = My_model.fit(my_alpha)
    My_forecast = My_model_fit.forecast(12)
    My_forecast = My_forecast.astype(int)
    return My_forecast  

# Holt Winter's Exponential Smoothing (HWES) forecasting 
def forecast_HWExpSmoothing(data, my_alpha, my_beta, my_gamma):
    # data (Pandas Data Series)-- historical data for past 12 month
    # my_alpha -- smoothing level from 0 to 1    
    # my_beta -- smoothing trend from 0 to 1
    # my_gamma -- smoothing seasonal from 0 to 1
    # return forecast for next 12 months 
    My_hist = data.to_numpy()
    My_model = ExponentialSmoothing(My_hist)
    # My_model_fit = My_model.fit(smoothing_level=my_alpha,smoothing_trend=my_beta, smoothing_seasonal=my_gamma)
    My_model_fit = My_model.fit(my_alpha, my_beta, my_gamma)
    My_forecast = My_model_fit.forecast(12)
    My_forecast = My_forecast.astype(int)
    return My_forecast    

# Forecast performance checking 
# MAE : Mean absolute Error
def MAE_value(expected_value, forecasted_value):
    # expected_value []
    # forecasted_value []
    mae = mean_absolute_error(expected_value, forecasted_value)
    return mae

# RMSE : Root Mean Squared Error
def RMSE_value(expected_value, forecasted_value):
    # expected_value []
    # forecasted_value []
    rmse =  mean_squared_error(expected_value, forecasted_value)
    rmse = sqrt (rmse)
    return rmse

# MSE : Mean Squared Error
def MSE_value(expected_value, forecasted_value):
    # expected_value []
    # forecasted_value []
    mse =  mean_squared_error(expected_value, forecasted_value)
    return mse

# MASE : Mean Absolute Squared Error
def MASE_value(expected_value, forecasted_value):
    # expected_value []
    # forecasted_value []
    mae =  mean_absolute_error(expected_value, forecasted_value)
    naive_value = np.roll(expected_value, 1)
    naive_mae = mean_absolute_error(expected_value[1:], naive_value[1:])
    mase = mae / naive_mae
    return mase

def main():
    InputFileName ="DataSet1.xlsx"
    ExcelTab = "RawData"
    try: 
        ExcelData = pd.ExcelFile(InputFileName)
    except:
        showinfo(title="File Error", \
                 message = f"{InputFileName} opening error.")
    else: 
        Data_df1 = pd.read_excel(ExcelData, ExcelTab)    
    
    # LED study 
    # Forecasting variables
    p=5
    d=0
    q=2
    alpha = 0.1
    beta = 0.6
    gamma = 0
    # Historical data    
    # My_data_list = Data_df1["LEDLight"].array
    My_data_list = Data_df1["Monitor"].array
    # My_data_list = Data_df1["Laptop"].array
    # My_data_list = Data_df1["Mobile"].array
    My_month_list = Data_df1["Month"].array

    LED_Y2021 = pd.Series(data = My_data_list[0:12], index = My_month_list[0:12])  
    LED_Y2022 = pd.Series(data = My_data_list[12:], index = My_month_list[12:])
    # Forecast data
    AR_forecast = forecast_ARIMA(LED_Y2022, p, 0, 0)   
    MA_forecast = forecast_ARIMA(LED_Y2022, 0, 0, q)
    ARMA_forecast = forecast_ARIMA(LED_Y2022, p, 0, q)
    
    ExpSmoothing_forecast1 = forecast_HWExpSmoothing (LED_Y2022, alpha, 0, 0)
    ExpSmoothing_forecast2 = forecast_HWExpSmoothing (LED_Y2022, alpha, beta, 0)
    ExpSmoothing_forecast3 = forecast_HWExpSmoothing (LED_Y2022, alpha, beta, gamma)
    
    # plotting line chart
    plt.figure(figsize=(16,7))

    My_13_month= LED_Y2022.index[11] + relativedelta(months=+1)
    My_forecast_index = pd.date_range(My_13_month, periods=12, freq='m')

    # plotting historical data and forecast 
    plt.plot(LED_Y2022.index, LED_Y2022 , color = 'black', label ="Historical")
    plt.plot(My_forecast_index, AR_forecast, marker='o',label ='AR')
    plt.plot(My_forecast_index, MA_forecast, marker='o', label ='AM')
    plt.plot(My_forecast_index, ARMA_forecast, marker='o', label = "ARAM")
    plt.plot(My_forecast_index, ExpSmoothing_forecast1, marker='D', label = "SES")
    plt.plot(My_forecast_index, ExpSmoothing_forecast2, marker='D', label = "Holt")
    
    plt.title("LED")
    My_x_index = pd.date_range(LED_Y2022.index[0], periods=24, freq="m") 
    
    plt.xticks (My_x_index)
    plt.tick_params(labelrotation=45)
    plt.legend ()
    plt.show ()
       
    # check forecast performance for LED Light
    expected_LED_2022 = LED_Y2022.array
    mae_ARMA_LED = MAE_value (expected_LED_2022, ARMA_forecast)
    print ('MAE for LED Light: %f' %mae_ARMA_LED)
    
    rmse_ARMA_LED = RMSE_value (expected_LED_2022, ARMA_forecast)    
    print ('RMSE for LED Light: %f' %rmse_ARMA_LED)    

if __name__ == "__main__":
    main()