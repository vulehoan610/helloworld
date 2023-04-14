import pandas as pd
import win32com.client as win32
import datetime 
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

import sqlite3
import numpy as np

# connect to database

sheet_month = input('Nhập sheet name:')    
# đọc file weekly report
week_activity = pd.read_excel(r"D:\\Nordson\\report\\weekly report\\FY23\\Weekly Report (Vu).xlsx",sheet_name = sheet_month,skiprows = 5,usecols='A:L', index_col=1)
week_activity.drop(columns=['Unnamed: 0','Contact Name / Position','Unnamed: 4'],axis=1,inplace=True)

# sửa thông tin
week_activity.dropna(subset=['Time'],inplace=True)
week_activity['Date '].fillna(method='ffill',axis=0,inplace=True)
week_activity.rename({'Date ': 'job_date','Time': 'time_in','Unnamed: 5': 'time_out', 'Company Name': 'customer_name','Location / Province':'Location','Reason for visit':'reason_visit',' Visiting result':'visiting_result','Next Action / Follow up ':'next_action'}, axis = "columns", inplace = True)  
# lưu vào database
conn = sqlite3.connect('icsdata.db')
week_activity.to_sql('work_history',conn, if_exists='append', index = True)
conn.commit()
conn.close()
print('Đã lưu vào database')
print('Test git clone')
