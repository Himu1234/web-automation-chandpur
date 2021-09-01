import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import WebDriverWait # Required for explicit wait
from selenium.webdriver.support import expected_conditions as ec # Required for explicit wait
from selenium.webdriver.common.by import By # Required for explicit wait

excel_file_main = 'Main.xlsx'
excel_file_test = 'Test.xlsx'
excel_file_out = 'Output.xlsx'
driver_exe = 'chromedriver.exe'
df1 = pd.read_excel(os.path.join(os.getcwd(),excel_file_main))
df2 = pd.read_excel(os.path.join(os.getcwd(),excel_file_test))
df3 = pd.read_excel(os.path.join(os.getcwd(),excel_file_out))
df_t = []

for x in range(len(df2.index)):

    df_t = df1[df1["New_Meter_No"] == int(df2.iat[x, 0])]
    df3 = df3.append(df_t)
    
df3.to_excel(os.path.join(os.getcwd(),excel_file_out), sheet_name = 'Sheet_3')
