# Import python modules
import tkinter as tk
from tkinter import TclError, ttk, Text
from tkinter import *
from tkinter.messagebox import showerror, showwarning, showinfo
import pandas as pd
from pandas.tseries.offsets import DateOffset
import numpy as np
from tkinter import filedialog
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import statsmodels.api as sm
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import date, datetime
from dateutil.relativedelta import relativedelta
import Lab_3_Forecasting as lab_integration

# Block 0: create the main GUI window
MainWindow = tk.Tk()
MainWindow.title ('Demand Forecasting')
MainWindow.geometry("910x710")
MainWindow.resizable(False,True)

MainMenu = tk.Menu(MainWindow)
MainWindow.config(menu=MainMenu)
FileMenu = tk.Menu(MainMenu)
MainMenu.add_cascade(label='File', menu=FileMenu)
FileMenu.add_command(label='Open from')
FileMenu.add_command(label='Save to')
FileMenu.add_separator()
FileMenu.add_command(label='Exit', command=MainWindow.destroy)
HelpMenu = tk.Menu(MainMenu)
MainMenu.add_cascade(label='Help', menu=HelpMenu)
HelpMenu.add_command(label='About')

# Block 1: show author name
AuthorFrame = ttk.Frame(MainWindow)
AuthorFrame.grid(column=0, row=0)
ttk.Label(AuthorFrame, text="Author : Vicki").grid(
    column=0, row=0, sticky=tk.W)
ttk.Label(AuthorFrame, text="Group : Solo").grid(column=0, row=1, sticky=tk.W)

# Block 2: upload historical data from Excel file
MonthHist = pd.Series(24)
LEDHist = pd.Series(24)
MonitorHist = pd.Series(24)
LaptopHist = pd.Series(24)
MobileHist = pd.Series(24)
InputFileName = tk.StringVar()
File_label = ttk.Label(MainWindow, text="Input File:")
File_label.grid(column=1, row=0, sticky=tk.E, padx=5, pady=5)
File_entry = tk.Label(text='', bg='white', fg='black',
                      width=39, anchor="w", justify="left")
File_entry.grid(column=1, row=0, sticky='NW',
                padx=(262, 0), pady=(10, 0), columnspan=4)

def Upload_data():
    global MonthHist
    global LEDHist
    global MonitorHist
    global LaptopHist
    global MobileHist

    """ upload data from excel file """
    InputFileName = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")])

    if InputFileName == "":
        showinfo(title='Data uploading',
                 message='Please type in a valid file name.')
    else:
        try:
            ExcelFileInput = pd.ExcelFile(InputFileName)
            File_entry['text'] = InputFileName.split(
                '/')[len(InputFileName.split('/'))-1]
        except:
            showinfo(title="File open",
                     message="Error is found for file openning.")
        else:
            df1 = pd.read_excel(ExcelFileInput, "RawData")

            MonthHist = df1["Month"]
            LEDHist = df1["LEDLight"]
            MonitorHist = df1["Monitor"]
            LaptopHist = df1["Laptop"]
            MobileHist = df1["Mobile"]

            showinfo(title='Data uploading',
                     message='The historical data has been uploaded successfully!')

Upload_button = ttk.Button(MainWindow, text="Upload Data", command=Upload_data)
Upload_button.grid(column=1, row=0)

# Block 3: select a product from combobox and display historical data using treeview
Selected_item = ""
ItemSelection_cb = ttk.Combobox(MainWindow, textvariable=Selected_item)
ItemSelection_cb['values'] = ["LEDLight", "Monitor", "Laptop", "Mobile"]
ItemSelection_cb.set("")
ItemSelection_cb['state'] = 'readonly'    
ItemSelection_cb.grid(column=0, row= 2, sticky = tk.E)    

SeperatorFrame = ttk.Frame(MainWindow, width=800, height=50)
SeperatorFrame.grid (column = 0, row=1, columnspan =4, sticky=tk.EW)

HistoricalDataLabel = ttk.Label(MainWindow, text = "Historical data for:" )
HistoricalDataLabel.grid(column = 0, row = 2, sticky=tk.W)

HistColumns = ('SerialNo', 'HistMonth', 'HistSales' )
HistView = ttk.Treeview(MainWindow, columns=HistColumns, height=12, show='headings')

HistView.column("SerialNo", width=50 )
HistView.column("HistMonth", width=100 )
HistView.column("HistSales", width=100)

HistView.heading('SerialNo', text='No')
HistView.heading('HistMonth', text='Date')
HistView.heading('HistSales', text='Sales')

for i in range(12):
    HistView.insert('', tk.END, iid = i+1, text = "", values=(i+1, "" , ""))
   
HistView.grid(column = 0, row=3)

HistView2 = ttk.Treeview(MainWindow, columns=HistColumns, height=12, show='headings')
HistView2.column("SerialNo", width=50 )
HistView2.column("HistMonth", width=100 )
HistView2.column("HistSales", width=100)

HistView2.heading('SerialNo', text='No')
HistView2.heading('HistMonth', text='Date')
HistView2.heading('HistSales', text='Sales')

for i in range(12):
    HistView2.insert('', tk.END, iid = i+1, text = "", values=(12+i+1, "" , ""))
   
HistView2.grid(column = 1, row=3)

def Item_changed(event):
    """ handle the month changed event """
    
    ItemName=ItemSelection_cb.get()
    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        match ItemName:
            case 'LEDLight':
                for i in range(12):
                    HistView.item(i+1, text = LEDHist[i], values=(i+1, MonthHist[i].date(), LEDHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = LEDHist[12+i], values=(12+i+1, MonthHist[12+i].date(), LEDHist[12+i]))
            case 'Monitor': 
                for i in range(12):
                    HistView.item(i+1, text = MonitorHist[i], values=(i+1, MonthHist[i].date(), MonitorHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = MonitorHist[12+i], values=(12+i+1, MonthHist[12+i].date(), MonitorHist[12+i]))

            case 'Laptop' :
                for i in range(12):
                    HistView.item(i+1, text = LaptopHist[i], values=(i+1, MonthHist[i].date(), LaptopHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = LaptopHist[12+i], values=(12+i+1, MonthHist[12+i].date(), LaptopHist[12+i]))

            case 'Mobile' :
                for i in range(12):
                    HistView.item(i+1, text = MobileHist[i], values=(i+1, MonthHist[i].date(), MobileHist[i]))
                for i in range(12):
                    HistView2.item(i+1, text = MobileHist[12+i], values=(12+i+1, MonthHist[12+i].date(), MobileHist[12+i]))

ItemSelection_cb.bind('<<ComboboxSelected>>', Item_changed)
    
# Block 4: display forecasting parameters
p_string =tk.StringVar()
d_string = tk.StringVar()
q_string = tk.StringVar()
alpha_string = tk.StringVar()
beta_string = tk.StringVar()
gamma_string = tk.StringVar()

# set default paramters
p_string = "2"
d_string = "0"
q_string = "1"
alpha_string = "0.2"
beta_string = "0.3"
gamma_string = "0.0"

p_value = int()
d_value = int()
q_value = int()
alpha_value = float()
beta_value = float()
gamma_value = float()

ParameterFrame = ttk.Frame(MainWindow, width=253, height=70)
# ParameterFrame['padding'] = (2,2,2,2)
ParameterFrame['borderwidth'] = 5
ParameterFrame['relief'] ="groove"
ParameterFrame.grid_propagate(False)
ParameterFrame.grid(column = 0, row = 4)

ttk.Label(ParameterFrame, text = "Forecasting parameters:").grid( column = 0, row =0, columnspan= 4, sticky=tk.W)
ttk.Label(ParameterFrame, text = "p:").grid (column = 0, row =1, sticky=tk.E)
ttk.Label(ParameterFrame, text = "d:").grid (column = 2, row =1, sticky=tk.E)
ttk.Label(ParameterFrame, text = "q:").grid (column = 4, row =1, sticky=tk.E)

p_entry = ttk.Entry(ParameterFrame,   width = 4)
d_entry = ttk.Entry(ParameterFrame,   width = 4)
q_entry = ttk.Entry(ParameterFrame,   width = 4)

p_entry.grid(column=1,row =1)
d_entry.grid(column=3, row =1)
q_entry.grid(column=5, row =1)

p_entry.insert(0, p_string)
d_entry.insert(0, d_string)
q_entry.insert(0, q_string)

ttk.Label(ParameterFrame, text = "level:").grid (column = 0, row =2, sticky=tk.E)
ttk.Label(ParameterFrame, text = "trend:").grid (column = 2, row =2, sticky=tk.E)
ttk.Label(ParameterFrame, text = "season:").grid (column = 4, row =2, sticky=tk.E)
alpha_entry = ttk.Entry(ParameterFrame,  width = 4)
beta_entry = ttk.Entry (ParameterFrame,  width = 4)
gamma_entry = ttk.Entry (ParameterFrame, width = 4)
alpha_entry.grid(column=1,row =2)
beta_entry.grid(column=3, row =2)
gamma_entry.grid(column=5, row =2)
alpha_entry.insert(0, alpha_string)
beta_entry.insert(0, beta_string)
gamma_entry.insert(0, gamma_string)

# Block 5: display empty forecast result in treeview
ResultLabel = ttk.Label(MainWindow, text = " Demand Forecast " )
ResultLabel.grid(column = 3, row = 2)

ResultColumns = ('SerialNo', 'Month', 'Forecast' )
ResultView = ttk.Treeview(MainWindow, columns=ResultColumns, height=12, show='headings')

ResultView.column("SerialNo", width=50 )
ResultView.column("Month", width=100 )
ResultView.column("Forecast", width=100)

ResultView.heading('SerialNo', text='No')
ResultView.heading('Month', text='Month')
ResultView.heading('Forecast', text='Forecasts')

for i in range(12):
    ResultView.insert('', tk.END, iid = i+1, text = "", values=(i+1, "" , ""))
   
ResultView.grid(column = 3, row=3)    

# Block 6: algorithm Radiobutton for selection
ttk.Label(MainWindow,text = "Algorithms").grid( column = 2, row = 2 )

RadioFrame = ttk.Frame(MainWindow, width=150, height=120)
RadioFrame['padding'] = (10,10,10,10)
RadioFrame['borderwidth'] = 5
RadioFrame['relief'] = 'sunken'
RadioFrame.grid_propagate(False)
RadioFrame.grid(column = 2, row = 3, sticky=tk.NW)

AlgoSelected = tk.IntVar()
AlgoButton1 = ttk.Radiobutton(RadioFrame, text='Algo 1: AR', variable=AlgoSelected, value=1)
AlgoButton1.grid(column=1, row=0, padx=2, pady=2, sticky=tk.W)

AlgoButton2 = ttk.Radiobutton(RadioFrame, text='Algo 2: ARMA', variable=AlgoSelected, value=2)
AlgoButton2.grid(column=1, row=1, padx=2, pady=2, sticky=tk.W)

AlgoButton3 = ttk.Radiobutton(RadioFrame, text='Algo 3: SES', variable=AlgoSelected, value=3)
AlgoButton3.grid(column=1, row=2, padx=2, pady=2, sticky=tk.W)

AlgoButton3 = ttk.Radiobutton(RadioFrame, text='Algo 4: HWES', variable=AlgoSelected, value=4)
AlgoButton3.grid(column=1, row=3, padx=2, pady=2, sticky=tk.W)

AlgoSelected.set(2)

# Block 7: accuracy checkboxs for selections
MAE_var = tk.StringVar()
MSE_var = tk.StringVar()
RMSE_var = tk.StringVar()
MASE_var = tk.StringVar()

AccuracyFrame = ttk.Frame(MainWindow, width=150, height=120)
AccuracyFrame['padding'] = (10,10,10,10)
AccuracyFrame['borderwidth'] = 5
AccuracyFrame['relief'] = 'sunken'
AccuracyFrame.grid_propagate(False)
AccuracyFrame.grid(column = 2, row = 3, columnspan = 2, sticky=tk.SW)

# display MAE error 
def Display_MAE_error ():
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    AR_forecast = lab_integration.forecast_AR(My_Y2021, p_value, 0, 0)
                    mae_ARMA = lab_integration.MAE_value(My_Y2022, AR_forecast)
                    MAEResult_label.config(text = '%.2f'%mae_ARMA)
                case 2:
                    # ARMA algo is selected
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2021, p_value, 0, q_value)
                    mae_ARMA = lab_integration.MAE_value(My_Y2022, ARMA_forecast)
                    MAEResult_label.config(text = '%.2f'%mae_ARMA)
                case 3:
                   # SES algo is selected
                   SES_forecast = lab_integration.forecast_SExpSmoothing(My_Y2021, alpha_value)
                   mae_SES = lab_integration.MAE_value(My_Y2022, SES_forecast)
                   MAEResult_label.config(text = '%.2f'%mae_SES)                
                case 4:
                    # Holt-Winters algo is selected
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2021, alpha_value, beta_value, gamma_value)
                    mae_HWES = lab_integration.MAE_value(My_Y2022, HWES_forecast)
                    MAEResult_label.config(text = '%.2f'%mae_HWES)                
                case _: 
                    print ("Error in algo or accuracy measure selection.")
    

# display RMSE error 
def Display_RMSE_error ():    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    print("AR")
                    AR_forecast = lab_integration.forecast_AR(My_Y2021, p_value, 0, 0)
                    print("prediction for Y2022: ")
                    print(AR_forecast)
                    rmse_AR = lab_integration.RMSE_value(My_Y2022, AR_forecast)
                    print ('RMSE for selected product and algo: %f' %rmse_AR)
                    RMSEResult_label.config(text = '%.2f'%rmse_AR)
                case 2:
                    # ARMA algo is selected
                    print("ARMA")
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2021, p_value, 0, q_value)
                    print("prediction for Y2022: ")
                    print(ARMA_forecast)
                    rmse_ARMA = lab_integration.RMSE_value(My_Y2022, ARMA_forecast)
                    print ('RMSE for selected product and algo: %f' %rmse_ARMA)
                    RMSEResult_label.config(text = '%.2f'%rmse_ARMA)
                case 3:
                   # SES algo is selected
                    print("Single")
                    SES_forecast = lab_integration.forecast_SExpSmoothing(My_Y2021, alpha_value)
                    print("prediction for Y2022")
                    print(SES_forecast)
                    rmse_SES = lab_integration.RMSE_value(My_Y2022, SES_forecast)
                    print('RMSE for selected product and algo: %f' %rmse_SES)
                    RMSEResult_label.config(text = '%.2f'%rmse_SES)  
                case 4:
                    # Holt-Winters algo is selected
                    print("Holt Winters")
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2021, alpha_value, beta_value, gamma_value)
                    print("prediction for Y2022")
                    print(HWES_forecast)
                    rmse_HWES = lab_integration.RMSE_value(My_Y2022, HWES_forecast)
                    print('RMSE for selected product and algo: %f' %rmse_HWES)
                    RMSEResult_label.config(text = '%.2f'%rmse_HWES)                
                case _: 
                    print("Error in algo or accuracy measure selection.")

# display MSE error
def Display_MSE_error ():    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    print("AR")
                    AR_forecast = lab_integration.forecast_AR(My_Y2021, p_value, 0, 0)
                    print("prediction for Y2022: ")
                    print(AR_forecast)
                    mse_AR = lab_integration.MSE_value(My_Y2022, AR_forecast)
                    print ('MSE for selected product and algo: %f' %mse_AR)
                    MSEResult_label.config(text = '%.2f'%mse_AR)                
                case 2:
                    # ARMA algo is selected
                    print("ARMA")
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2021, p_value, 0, q_value)
                    print("prediction for Y2022: ")
                    print(ARMA_forecast)
                    mse_ARMA = lab_integration.MSE_value(My_Y2022, ARMA_forecast)
                    print ('MSE for selected product and algo: %f' %mse_ARMA)
                    MSEResult_label.config(text = '%.2f'%mse_ARMA)                
                case 3:
                   # SES algo is selected
                    print("Single")
                    SES_forecast = lab_integration.forecast_SExpSmoothing(My_Y2021, alpha_value)
                    print("prediction for Y2022")
                    print(SES_forecast)
                    mse_SES = lab_integration.MSE_value(My_Y2022, SES_forecast)
                    print('MSE for selected product and algo: %f' %mse_SES)
                    MSEResult_label.config(text = '%.2f'%mse_SES)  
                case 4:
                    # Holt-Winters algo is selected
                    print("Holt Winters")
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2021, alpha_value, beta_value, gamma_value)
                    print("prediction for Y2022")
                    print(HWES_forecast)
                    mse_HWES = lab_integration.MSE_value(My_Y2022, HWES_forecast)
                    print('MSE for selected product and algo: %f' %mse_HWES)
                    MSEResult_label.config(text = '%.2f'%mse_HWES)  
                case _: 
                    print("Error in algo or accuracy measure selection.") 

# display MASE error
def Display_MASE_error ():    
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    print("AR")
                    AR_forecast = lab_integration.forecast_AR(My_Y2021, p_value, 0, 0)
                    print("prediction for Y2022: ")
                    print(AR_forecast)
                    mase_AR = lab_integration.MASE_value(My_Y2022, AR_forecast)
                    print ('MASE for selected product and algo: %f' %mase_AR)
                    MASEResult_label.config(text = '%.2f'%mase_AR)  
                case 2:
                    # ARMA algo is selected
                    print("ARMA")
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2021, p_value, 0, q_value)
                    print("prediction for Y2022: ")
                    print(ARMA_forecast)
                    mase_ARMA = lab_integration.MASE_value(My_Y2022, ARMA_forecast)
                    print ('MASE for selected product and algo: %f' %mase_ARMA)
                    MASEResult_label.config(text = '%.2f'%mase_ARMA)  
                case 3:
                   # SES algo is selected
                    print("Single")
                    SES_forecast = lab_integration.forecast_SExpSmoothing(My_Y2021, alpha_value)
                    print("prediction for Y2022")
                    print(SES_forecast)
                    mase_SES = lab_integration.MASE_value(My_Y2022, SES_forecast)
                    print('MASE for selected product and algo: %f' %mase_SES)
                    MASEResult_label.config(text = '%.2f'%mase_SES)  
                case 4:
                    # Holt-Winters algo is selected
                    print("Holt Winters")
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2021, alpha_value, beta_value, gamma_value)
                    print("prediction for Y2022")
                    print(HWES_forecast)
                    mase_HWES = lab_integration.MASE_value(My_Y2022, HWES_forecast)
                    print('MASE for selected product and algo: %f' %mase_HWES)
                    MASEResult_label.config(text = '%.2f'%mase_HWES)  
                case _: 
                    print("Error in algo or accuracy measure selection.") 

def MAE_changed():
    if MAE_var.get() == "selected":
        Display_MAE_error()
    else: 
        if MAE_var.get() == "unselected":
            MAEResult_label.config(text = "")    
        else:
            showinfo(title="err", message=f" MAE var = {MAE_var.get()}, !")

def RMSE_changed():
    if RMSE_var.get() == "selected":
        Display_RMSE_error()
    else:    
        if RMSE_var.get() == "unselected":
            RMSEResult_label.config(text = "")
        else:
            showinfo(title="err", message=f" RMAE var = {RMSE_var.get()}, !")
    
def MSE_changed ():
    if MSE_var.get() == "selected":
        Display_MSE_error()
    else:
        if MSE_var.get() == "unselected":
            MSEResult_label.config(text="")
        else: 
            showinfo(title="err", message=f" MSE var = {MSE_var.get()}, !")
        
def MASE_changed ():
    if MASE_var.get() == "selected":
        Display_MASE_error()
    else:
        if MASE_var.get() == "unselected":
            MASEResult_label.config(text="")
        else: 
            showinfo(title="err", message=f" MASE var = {MASE_var.get()}, !")

MAE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='MAE',
                command=MAE_changed,
                variable=MAE_var,
                onvalue="selected",
                offvalue="unselected")
MAE_checkbox.grid(column=0, row = 1, sticky=tk.W)

RMSE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='RMSE',
                command=RMSE_changed,
                variable=RMSE_var,
                onvalue='selected',
                offvalue='unselected')
RMSE_checkbox.grid(column=0, row = 2, sticky=tk.W)

MSE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='MSE',
                command=MSE_changed,
                variable=MSE_var,
                onvalue='selected',
                offvalue='unselected')
MSE_checkbox.grid(column=0, row = 3, sticky=tk.W)

MASE_checkbox = ttk.Checkbutton(AccuracyFrame,
                text='MASE',
                command=MASE_changed,
                variable=MASE_var,
                onvalue='selected',
                offvalue='unselected')
MASE_checkbox.grid(column=0, row = 4, sticky=tk.W)

MAEResult_label = ttk.Label(AccuracyFrame, relief='groove', width = 5, padding = 2)
MAEResult_label.config(text = "1020")
MAEResult_label.grid(column=1, row=1, sticky=tk.E)

RMSEResult_label = ttk.Label(AccuracyFrame, relief='groove', width = 5, padding=2)
RMSEResult_label.config(text = "3333")
RMSEResult_label.grid(column=1, row=2, sticky=tk.E)

MSEResult_label = ttk.Label(AccuracyFrame, relief='groove', width = 5, padding=2)
MSEResult_label.config(text = "4050")
MSEResult_label.grid(column=1, row=3, sticky=tk.E)

MASEResult_label = ttk.Label(AccuracyFrame, relief='groove', width = 5, padding=2)
MASEResult_label.config(text = "6666")
MASEResult_label.grid(column=1, row=4, sticky=tk.E)

# Block 8: forecast generation and accuracy checking
def plot_graph1(actual_values1):
    fig = plt.figure(figsize=(2, 2)) 
    plt.plot(actual_values1, marker='.', color='red')
    plt.tight_layout(pad=0.1)
    canvas = FigureCanvasTkAgg(fig, master=Graph1Frame)
    canvas.get_tk_widget().grid(column=0, row=5)
    canvas.draw()

def plot_graph2(actual_values2):
    fig = plt.figure(figsize=(2, 2)) 
    plt.plot(actual_values2, marker='.', color='blue')
    plt.tight_layout(pad=0.1)
    canvas = FigureCanvasTkAgg(fig, master=Graph2Frame)
    canvas.get_tk_widget().grid(column=1, row=5)
    canvas.draw()

def plot_graph3(forecasted_values):
    fig = plt.figure(figsize=(2, 2)) 
    plt.plot(forecasted_values, marker='.', color='purple')
    plt.tight_layout(pad=0.1)
    canvas = FigureCanvasTkAgg(fig, master=Graph3Frame)
    canvas.get_tk_widget().grid(column=3, row=5)
    canvas.draw()

def Forecasting_clicked ():
    if len(LEDHist) < 24:
        showinfo (title="Error", message="Please upload data first.")
    else:
        # 1. get historical data for selected product  
        current_product = ItemSelection_cb.get()
        match current_product:
            case "LEDLight":
                my_hist = LEDHist
            case "Monitor":
                my_hist = MonitorHist
            case "Laptop":
                my_hist = LaptopHist
            case "Mobile":
                my_hist = MobileHist

        My_Y2021 = my_hist[0:12]
        My_Y2022 = my_hist[12:]   
        
        # 2. checking forecasting parameters 
        try: 
            p_value = int(p_entry.get())
            d_value = int(d_entry.get())
            q_value = int(q_entry.get())
            alpha_value = float(alpha_entry.get())
            beta_value = float(beta_entry.get())
            gamma_value = float(gamma_entry.get())
        except:     
            showinfo(title="Parameter error", \
                     message = "p, d, and q parameters must be integers; alpha, beta, and gamma must be decimals.")
        else:
            # 3. read forecasting algo
            match AlgoSelected.get():
                case 1:
                    # AR algo is selected
                    AR_forecast = lab_integration.forecast_AR(My_Y2022, p_value, 0 ,0)
                    # display forecast on treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i+1, values=(i+1, New_month[i].date(), AR_forecast[i]))
                        plot_graph1(My_Y2021)
                        plot_graph2(My_Y2022)
                        plot_graph3(AR_forecast)
                case 2:
                    # ARMA algo is selected
                    ARMA_forecast = lab_integration.forecast_ARIMA(My_Y2022, p_value, 0, q_value)
                    # display forecast on treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i+1, values=(i+1, New_month[i].date(), ARMA_forecast[i]))
                        plot_graph1(My_Y2021)
                        plot_graph2(My_Y2022)
                        plot_graph3(ARMA_forecast)
                case 3:
                   # SES algo is selected
                    SES_forecast = lab_integration.forecast_SExpSmoothing(My_Y2022, alpha_value)
                    # display forecast on treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i+1, values=(i+1, New_month[i].date(), SES_forecast[i]))
                        plot_graph1(My_Y2021)
                        plot_graph2(My_Y2022)
                        plot_graph3(SES_forecast)
                case 4:
                    # Holt-Winters algo is selected
                    HWES_forecast = lab_integration.forecast_HWExpSmoothing(My_Y2022, alpha_value, beta_value, gamma_value)
                    # display forecast on treeview
                    for i in range(12):
                        New_month = MonthHist + DateOffset(years=2)
                        ResultView.item(i+1, values=(i+1, New_month[i].date(), HWES_forecast[i]))
                        plot_graph1(My_Y2021)
                        plot_graph2(My_Y2022)
                        plot_graph3(HWES_forecast)
                case _: 
                    print ("Error in algo or accuracy measure selection.")

ForecastGenerationbutton = tk.Button(MainWindow, text = 'Forecasting >', \
                                     fg ='black', command=Forecasting_clicked)
ForecastGenerationbutton.grid( column = 2, row = 3)

# Block 9: result and comments by student using label
CommentFrame = ttk.Frame(MainWindow, width=456, height=70)
CommentFrame['padding'] = (2,2,2,2)
CommentFrame['borderwidth'] = 5
CommentFrame['relief'] ="groove"
CommentFrame.grid_propagate(False)
CommentFrame.grid(column=1, row=4, columnspan=5, sticky=tk.NSEW)
ttk.Label(CommentFrame, text="Result & comment by student: By the MAE, RMSE, MSE and MASE values, the smaller and closer to 1 of the naive forecast").grid(column=0, row=0)
ttk.Label(CommentFrame, text="that the values are, the more accurate the method in predicting the inventory demand. Hence, for monitor and laptop,").grid(column=0, row=1)
ttk.Label(CommentFrame, text=" ARMA is the most suitable method, AR method is most suitable for mobile, and SES/HWES is most suitable for LED Lights.").grid(column=0, row=2)

# Block 10: plot graph
Graph1Frame = ttk.Frame(MainWindow, width=256, height=266, style='new.TFrame')
Graph1Frame['borderwidth'] = 5
Graph1Frame['relief'] = 'sunken'
Graph1Frame.grid_propagate(False)
Graph1Frame.grid(column = 0, row = 5, columnspan = 2, sticky=tk.SW)

Graph2Frame = ttk.Frame(MainWindow, width=256, height=266, style='new.TFrame')
Graph2Frame['borderwidth'] = 5
Graph2Frame['relief'] = 'sunken'
Graph2Frame.grid_propagate(False)
Graph2Frame.grid(column = 1, row = 5, columnspan = 2, sticky=tk.SW)

Graph3Frame = ttk.Frame(MainWindow, width=256, height=266, style='new.TFrame')
Graph3Frame['borderwidth'] = 5
Graph3Frame['relief'] = 'sunken'
Graph3Frame.grid_propagate(False)
Graph3Frame.grid(column = 3, row = 5, columnspan = 2, sticky=tk.SW)

s = ttk.Style()
s.configure('new.TFrame', background='white')

# Block 11: main loop
MainWindow.mainloop()