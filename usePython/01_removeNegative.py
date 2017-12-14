# -*- coding: utf-8 -*-
"""
Created on Sun Sep 17 20:54:47 2017

@author: JZL
"""

#import numpy as np
import pandas as pd
import winreg
def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]

desk = get_desktop()
path = desk + '\\Book1.xlsx'

#raw = pd.read_csv(path, thousands=',', encoding = 'utf-8')
raw = pd.read_excel(path)
negatives = raw['Amount in doc. curr.'][raw['Amount in doc. curr.'] < 0].unique()

for negative in negatives:
    negRows = raw[raw['Amount in doc. curr.'] == negative]
    posRows = raw[raw['Amount in doc. curr.'] == -negative]
#    filtered = raw[(raw['Amount in doc. curr.'] == negative) | (raw['Amount in doc. curr.']==-negative)]
    for index, row in posRows.iterrows():
        if row['G/L'] == negRows.at[negRows.index[0], 'G/L']:
            if row['Profit Ctr'] == negRows.at[negRows.index[0],'Profit Ctr']:
                raw.loc[row.name,'FA']= 'Y'
                raw.loc[negRows.index[0], 'FA'] = 'Y'
                break

#raw.to_csv('C:\Users\Zhen Liu\Documents\Python\Book2.csv', encoding = 'utf-8')
raw.to_excel(desk + '\\Book2.xlsx', sheet_name = 'Sheet1', index=False)
