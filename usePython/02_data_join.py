# -*- coding: utf-8 -*-
"""
Created on Tue Oct 10 19:44:41 2017

@author: JZL
"""

import pandas as pd
import winreg

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]

desk = get_desktop()
path = desk + '\\Book1.xlsx'

sheet1 = pd.read_excel(path) # read sheet1 as dataframe
sheet2 = pd.read_excel(path, sheetname = 1) # read sheet2 as dataframe


parts = sheet1['Part no.'].unique() # get unique part no.

for part in parts: # loop each part in parts sequence
    diff = 0
    i = 0
    qty1 = sheet1['Deflict qty'][sheet1['Part no.'] == part]
    qty2= sheet2['Order Quantity'][sheet2['Part no.'] == part]
    
    for y, qt in enumerate(qty1): # loop qty1
        
        if len(qty2) == 0: # if qty2 is empty, break the loop
            break;
        diff = qty2[qty2.index[i]] - qt # get the diff between the first work need and first PO quantity
        if diff == 0:
            sheet1.at[qty1.index[y], 'PO'] = sheet2.at[qty2.index[i], 'Purchasing Doc Number']
            sheet1.at[qty1.index[y], 'residual QTY'] = 0
            i = i+1 # manually find next qty2 item
        elif diff > 0:
            sheet1.at[qty1.index[y], 'PO'] = sheet2.at[qty2.index[i], 'Purchasing Doc Number']
            sheet1.at[qty1.index[y], 'residual QTY'] = qty2[qty2.index[i]]-qt
            qty2[qty2.index[i]] = sheet1.at[qty1.index[y], 'residual QTY'] # put deduction into current PO quantity
            # Don't find next item
        elif diff < 0:
            try: # in order to avoid qty2.index[i+1] is out of bounds
                sheet1.at[qty1.index[y], 'PO'] = sheet2.at[qty2.index[i], 'Purchasing Doc Number'] +'+' 
                + sheet2.at[qty2.index[i+1], 'Purchasing Doc Number'] # put two PO into the work order
                sheet1.at[qty1.index[y], 'residual QTY'] = sheet2.at[qty2.index[i], 'Purchasing Doc Number'] 
                + sheet2.at[qty2.index[i+1], 'Purchasing Doc Number'] - qt
                qty2[qty2.index[i+1]] = sheet1.at[qty1.index[y], 'residual QTY'] # put remains into next qty2 item
                i = i+1
            except: # if try fails, run this part
                sheet1.at[qty1.index[y], 'PO'] = sheet2.at[qty2.index[i], 'Purchasing Doc Number']
                sheet1.at[qty1.index[y], 'residual QTY'] = qty2[qty2.index[i]] - qt
                break;
            # qty2 PO run out 
sheet1.to_excel(desk + '\\Book2.xlsx', index=False)
