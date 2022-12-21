import CreateService
import pandas as pd
from datetime import date
import json
import datetime as dt
import csv
from google.oauth2 import service_account
import gspread
import pprint as pp
import numpy as np
from gspread_dataframe import get_as_dataframe, set_with_dataframe
import time
from gspread_formatting import *


# Get MC Pack Events from Todays Date
df = pd.read_html('https://rfa.tile.hiveplatform.org/mc_pack_event?cell_id=&page_size=25000&offset=0',parse_dates=['mc_pack_event_timestamp'])[0]
df['date'] = pd.to_datetime(df['mc_pack_event_timestamp'], format="%H:%M")
#print("Enter Todays date")
#start_date=input("'yyyy-mm-dd': ")
today = date.today()
start_date = today.strftime("%Y/%m/%d")
#start_date = input("Type in Date: ")
mask = (df['date'] > start_date) 
df = df.loc[mask]
outFileName = ("ProdRecap_")
title = outFileName + start_date

# Group the time by Hour
df_group = df.groupby([pd.Grouper(key='date', freq='H'),df.line_id]).size().reset_index(name='Units/Hour')
df_date = df_group['date'].drop_duplicates()
df_date = pd.DataFrame(df_date)
# Remove the 'YY-MM-DD' to create values for start time
df_date['date'] = pd.to_datetime(df_date['date'],format).apply(lambda x: x.time())

#Get google account authorization 
gc = gspread.oauth()
sh = gc.open("TileProd")

#Grabbing the Copy of Template worksheet and renaming it 
spreadsheetId = '1SMANrp8dgZO5V6RDGiIbc7qBzvRIguO_gPQXKt3NJls'
wsh = gc.open_by_key(spreadsheetId).worksheet("Copy of Template")
# update your worksheet name:

wsh.update_title(title)

worksheet_title = title

#look for or create new sheet
try:
    worksheet = sh.worksheet(worksheet_title)
except gspread.WorksheetNotFound:
    worksheet = sh.add_worksheet(title=worksheet_title, rows=100, cols=100)


def countdown(t):
    
    while t:
        mins, secs = divmod(t, 60)
        timer = '{:02d}:{:02d}'.format(mins, secs)
        print(timer, end="\r")
        time.sleep(1)
        t -= 1
      
    
  
  
# input time in seconds
t = 30
  

def line1Kit():
    wait = time.sleep(5)
     #group line 1 by hour
    group = df.groupby([pd.Grouper(key='date', freq='H'),df.line_id]).size().reset_index(name='Units/Hour')
    df1 = group[group.line_id == 'KITTILELINE1']
    #format datetime to 00:00
    df1['date'] = df1['date'].dt.strftime('%H:%M')
    #Find Individual Times
  
    yo = df1.loc[df1['date']=='00:00']
    yo1 = df1.loc[df1['date']=='05:00']
    yo2 = df1.loc[df1['date']=='06:00']
    yo3 = df1.loc[df1['date']=='07:00']
    yo4 = df1.loc[df1['date']=='08:00']
    yo5 = df1.loc[df1['date']=='09:00']
    yo6 = df1.loc[df1['date']=='10:00']
    yo7 = df1.loc[df1['date']=='11:00']
    yo8 = df1.loc[df1['date']=='12:00']
    yo9 = df1.loc[df1['date']=='13:00']

        #drop date and lin_id columns
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)
    
    
    set_with_dataframe(worksheet, bo, row=3, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo1, row=4, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo2, row=5, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo3, row=6, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo4, row=7, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo5, row=8, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo6, row=9, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo7, row=10, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo8, row=11, col=4, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo9, row=12, col=4, include_column_header=False)

def missed1():
    df1 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line1-qc&page_size=2000&offset=0')[0]
    df2 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line2-qc&page_size=2000&offset=0')[0]
    df3 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line3-qc&page_size=2000&offset=0')[0]
    df4 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line4-qc&page_size=2000&offset=0')[0]
    df5 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line5-qc&page_size=2000&offset=0')[0]
    df6 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line6-qc&page_size=2000&offset=0')[0]
    
    df1['date'] = pd.to_datetime(df1['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    df2['date'] = pd.to_datetime(df2['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    df3['date'] = pd.to_datetime(df3['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    df4['date'] = pd.to_datetime(df4['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    df5['date'] = pd.to_datetime(df5['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    df6['date'] = pd.to_datetime(df6['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    
    
    mask1 = (df1['date'] > start_date)
    mask2 = (df2['date'] > start_date)
    mask3 = (df3['date'] > start_date)
    mask4 = (df4['date'] > start_date)
    mask5 = (df5['date'] > start_date)
    mask6 = (df6['date'] > start_date)

    df1 = df1.loc[mask1]
    df2 = df2.loc[mask2]
    df3 = df3.loc[mask3]
    df4 = df4.loc[mask4]
    df5 = df5.loc[mask5]
    df6 = df6.loc[mask6]
    
    l1 = df1[df1['barcode_data'].str.startswith('F0', na=False)]
    g1 = l1.groupby([pd.Grouper(key='date', freq='H'),l1.line_id]).size().reset_index(name='Missed/Hour')
    
    yo = g1.loc[g1['date']=='00:00']
    yo1 = g1.loc[g1['date']=='05:00']
    yo2 = g1.loc[g1['date']=='06:00']
    yo3 = g1.loc[g1['date']=='07:00']
    yo4 = g1.loc[g1['date']=='08:00']
    yo5 = g1.loc[g1['date']=='09:00']
    yo6 = g1.loc[g1['date']=='10:00']
    yo7 = g1.loc[g1['date']=='11:00']
    yo8 = g1.loc[g1['date']=='12:00']
    yo9 = g1.loc[g1['date']=='13:00']
  
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    try:
        bo = bo['Units/Hour'].item()
        worksheet.update("D3", bo, raw=False)
    except:
        set_with_dataframe(worksheet, bo, row=3, col=5, include_column_header=False)
    try:
        bo1 = bo1['Units/Hour'].item()
        worksheet.update("D4", bo1, raw=False)
    except:
        set_with_dataframe(worksheet, bo1, row=4, col=5, include_column_header=False)
    try:
        bo2 = bo2['Units/Hour'].item()
        worksheet.update("D5", bo2, raw=False)
    except:
        set_with_dataframe(worksheet, bo2, row=5, col=5, include_column_header=False)
    try:
        bo3 = bo3['Units/Hour'].item()
        worksheet.update("D6", bo3, raw=False)
    except:
        set_with_dataframe(worksheet, bo3, row=6, col=5, include_column_header=False)
    try:
        bo4 = bo4['Units/Hour'].item()
        worksheet.update("D7", bo4, raw=False)
    except:
        set_with_dataframe(worksheet, bo4, row=7, col=5, include_column_header=False)
    try:
        bo5 = bo5['Units/Hour'].item()
        worksheet.update("D8", bo5, raw=False)
    except:
        set_with_dataframe(worksheet, bo5, row=8, col=5, include_column_header=False)
    try:
        bo6 = bo6['Units/Hour'].item()
        worksheet.update("D8", bo6, raw=False)
    except:
        set_with_dataframe(worksheet, bo6, row=9, col=5, include_column_header=False)
    try:
        bo7 = bo7['Units/Hour'].item()
        worksheet.update("D10", bo7, raw=False)
    except:
        set_with_dataframe(worksheet, bo7, row=10, col=5, include_column_header=False)
    try:
        bo8 = bo8['Units/Hour'].item()
        worksheet.update("D11", bo8, raw=False)
    except:
        set_with_dataframe(worksheet, bo8, row=11, col=5, include_column_header=False)
    try:
        bo9 = bo9['Units/Hour'].item()
        worksheet.update("D12", bo9, raw=False)
    except:
        set_with_dataframe(worksheet, bo9, row=12, col=5, include_column_header=False)

def line2Kit():
     #group line 2 by hour
    group = df.groupby([pd.Grouper(key='date', freq='H'),df.line_id]).size().reset_index(name='Units/Hour')
    df1 = group[group.line_id == 'KITTILELINE2']
    #format datetime to 00:00
    df1['date'] = df1['date'].dt.strftime('%H:%M')
    #Find Individual Times
    yo = df1.loc[df1['date']=='00:00']
    yo1 = df1.loc[df1['date']=='05:00']
    yo2 = df1.loc[df1['date']=='06:00']
    yo3 = df1.loc[df1['date']=='07:00']
    yo4 = df1.loc[df1['date']=='08:00']
    yo5 = df1.loc[df1['date']=='09:00']
    yo6 = df1.loc[df1['date']=='10:00']
    yo7 = df1.loc[df1['date']=='11:00']
    yo8 = df1.loc[df1['date']=='12:00']
    yo9 = df1.loc[df1['date']=='13:00']

    #drop date and lin_id columns
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    set_with_dataframe(worksheet, bo, row=3, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo1, row=4, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo2, row=5, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo3, row=6, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo4, row=7, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo5, row=8, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo6, row=9, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo7, row=10, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo8, row=11, col=6, include_column_header=False)
    set_with_dataframe(worksheet, bo9, row=12, col=6, include_column_header=False)

def missed2():
    df2 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line2-qc&page_size=2000&offset=0')[0]
    
    df2['date'] = pd.to_datetime(df2['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    
    
    mask2 = (df2['date'] > start_date)
    df2 = df2.loc[mask2]
    
    l1 = df2[df2['barcode_data'].str.startswith('F0', na=False)]
    g1 = l1.groupby([pd.Grouper(key='date', freq='H'),l1.line_id]).size().reset_index(name='Missed/Hour')
    
    yo = g1.loc[g1['date']=='00:00']
    yo1 = g1.loc[g1['date']=='05:00']
    yo2 = g1.loc[g1['date']=='06:00']
    yo3 = g1.loc[g1['date']=='07:00']
    yo4 = g1.loc[g1['date']=='08:00']
    yo5 = g1.loc[g1['date']=='09:00']
    yo6 = g1.loc[g1['date']=='10:00']
    yo7 = g1.loc[g1['date']=='11:00']
    yo8 = g1.loc[g1['date']=='12:00']
    yo9 = g1.loc[g1['date']=='13:00']

  
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    
    try:
        bo = bo['Units/Hour'].item()
        worksheet.update("F3", bo, raw=False)
    except:
        set_with_dataframe(worksheet, bo, row=3, col=7, include_column_header=False)
    try:
        bo1 = bo1['Units/Hour'].item()
        worksheet.update("F4", bo1, raw=False)
    except:
        set_with_dataframe(worksheet, bo1, row=4, col=7, include_column_header=False)
    try:
        bo2 = bo2['Units/Hour'].item()
        worksheet.update("F5", bo2, raw=False)
    except:
        set_with_dataframe(worksheet, bo2, row=5, col=7, include_column_header=False)
    try:
        bo3 = bo3['Units/Hour'].item()
        worksheet.update("F6", bo3, raw=False)
    except:
        set_with_dataframe(worksheet, bo3, row=6, col=7, include_column_header=False)
    try:
        bo4 = bo4['Units/Hour'].item()
        worksheet.update("F7", bo4, raw=False)
    except:
        set_with_dataframe(worksheet, bo4, row=7, col=7, include_column_header=False)
    try:
        bo5 = bo5['Units/Hour'].item()
        worksheet.update("F8", bo5, raw=False)
    except:
        set_with_dataframe(worksheet, bo5, row=8, col=7, include_column_header=False)
    try:
        bo6 = bo6['Units/Hour'].item()
        worksheet.update("F8", bo6, raw=False)
    except:
        set_with_dataframe(worksheet, bo6, row=9, col=7, include_column_header=False)
    try:
        bo7 = bo7['Units/Hour'].item()
        worksheet.update("F10", bo7, raw=False)
    except:
        set_with_dataframe(worksheet, bo7, row=10, col=7, include_column_header=False)
    try:
        bo8 = bo8['Units/Hour'].item()
        worksheet.update("F11", bo8, raw=False)
    except:
        set_with_dataframe(worksheet, bo8, row=11, col=7, include_column_header=False)
    try:
        bo9 = bo9['Units/Hour'].item()
        worksheet.update("F12", bo9, raw=False)
    except:
        set_with_dataframe(worksheet, bo9, row=12, col=7, include_column_header=False)

def line3Kit():
    wait = time.sleep(5)
     #group line 6 by hour
    group = df.groupby([pd.Grouper(key='date', freq='H'),df.line_id]).size().reset_index(name='Units/Hour')
    df1 = group[group.line_id == 'KITTILELINE3']
    #format datetime to 00:00
    df1['date'] = df1['date'].dt.strftime('%H:%M')
    #Find Individual Times
    yo = df1.loc[df1['date']=='00:00']
    yo1 = df1.loc[df1['date']=='05:00']
    yo2 = df1.loc[df1['date']=='06:00']
    yo3 = df1.loc[df1['date']=='07:00']
    yo4 = df1.loc[df1['date']=='08:00']
    yo5 = df1.loc[df1['date']=='09:00']
    yo6 = df1.loc[df1['date']=='10:00']
    yo7 = df1.loc[df1['date']=='11:00']
    yo8 = df1.loc[df1['date']=='12:00']
    yo9 = df1.loc[df1['date']=='13:00']

    #drop date and lin_id columns
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    set_with_dataframe(worksheet, bo, row=3, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo1, row=4, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo2, row=5, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo3, row=6, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo4, row=7, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo5, row=8, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo6, row=9, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo7, row=10, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo8, row=11, col=8, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo9, row=12, col=8, include_column_header=False)

def missed3():
    df3 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line3-qc&page_size=2000&offset=0')[0]
    
    df3['date'] = pd.to_datetime(df3['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    
    
    mask3 = (df3['date'] > start_date)
    df3 = df3.loc[mask3]
    
    l1 = df3[df3['barcode_data'].str.startswith('F0', na=False)]
    g1 = l1.groupby([pd.Grouper(key='date', freq='H'),l1.line_id]).size().reset_index(name='Missed/Hour')
    
    yo = g1.loc[g1['date']=='00:00']
    yo1 = g1.loc[g1['date']=='05:00']
    yo2 = g1.loc[g1['date']=='06:00']
    yo3 = g1.loc[g1['date']=='07:00']
    yo4 = g1.loc[g1['date']=='08:00']
    yo5 = g1.loc[g1['date']=='09:00']
    yo6 = g1.loc[g1['date']=='10:00']
    yo7 = g1.loc[g1['date']=='11:00']
    yo8 = g1.loc[g1['date']=='12:00']
    yo9 = g1.loc[g1['date']=='13:00']

  
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    
    try:
        bo = bo['Units/Hour'].item()
        worksheet.update("H3", bo, raw=False)
    except:
        set_with_dataframe(worksheet, bo, row=3, col=9, include_column_header=False)
    try:
        bo1 = bo1['Units/Hour'].item()
        worksheet.update("H4", bo1, raw=False)
    except:
        set_with_dataframe(worksheet, bo1, row=4, col=9, include_column_header=False)
    try:
        bo2 = bo2['Units/Hour'].item()
        worksheet.update("H5", bo2, raw=False)
    except:
        set_with_dataframe(worksheet, bo2, row=5, col=9, include_column_header=False)
    try:
        bo3 = bo3['Units/Hour'].item()
        worksheet.update("H6", bo3, raw=False)
    except:
        set_with_dataframe(worksheet, bo3, row=6, col=9, include_column_header=False)
    try:
        bo4 = bo4['Units/Hour'].item()
        worksheet.update("H7", bo4, raw=False)
    except:
        set_with_dataframe(worksheet, bo4, row=7, col=9, include_column_header=False)
    try:
        bo5 = bo5['Units/Hour'].item()
        worksheet.update("H8", bo5, raw=False)
    except:
        set_with_dataframe(worksheet, bo5, row=8, col=9, include_column_header=False)
    try:
        bo6 = bo6['Units/Hour'].item()
        worksheet.update("H8", bo6, raw=False)
    except:
        set_with_dataframe(worksheet, bo6, row=9, col=9, include_column_header=False)
    try:
        bo7 = bo7['Units/Hour'].item()
        worksheet.update("H10", bo7, raw=False)
    except:
        set_with_dataframe(worksheet, bo7, row=10, col=9, include_column_header=False)
    try:
        bo8 = bo8['Units/Hour'].item()
        worksheet.update("H11", bo8, raw=False)
    except:
        set_with_dataframe(worksheet, bo8, row=11, col=9, include_column_header=False)
    try:
        bo9 = bo9['Units/Hour'].item()
        worksheet.update("H12", bo9, raw=False)
    except:
        set_with_dataframe(worksheet, bo9, row=12, col=9, include_column_header=False)

def line4Kit():
    wait = time.sleep(5)
     #group line 6 by hour
    group = df.groupby([pd.Grouper(key='date', freq='H'),df.line_id]).size().reset_index(name='Units/Hour')
    df1 = group[group.line_id == 'KITTILELINE4']
    #format datetime to 00:00
    df1['date'] = df1['date'].dt.strftime('%H:%M')
    #Find Individual Times
    yo = df1.loc[df1['date']=='00:00']
    yo1 = df1.loc[df1['date']=='05:00']
    yo2 = df1.loc[df1['date']=='06:00']
    yo3 = df1.loc[df1['date']=='07:00']
    yo4 = df1.loc[df1['date']=='08:00']
    yo5 = df1.loc[df1['date']=='09:00']
    yo6 = df1.loc[df1['date']=='10:00']
    yo7 = df1.loc[df1['date']=='11:00']
    yo8 = df1.loc[df1['date']=='12:00']
    yo9 = df1.loc[df1['date']=='13:00']

    #drop date and lin_id columns
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    set_with_dataframe(worksheet, bo, row=3, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo1, row=4, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo2, row=5, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo3, row=6, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo4, row=7, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo5, row=8, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo6, row=9, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo7, row=10, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo8, row=11, col=10, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo9, row=12, col=10, include_column_header=False)

def missed4():
    df4 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line4-qc&page_size=2000&offset=0')[0]
    
    df4['date'] = pd.to_datetime(df4['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    
    
    mask4 = (df4['date'] > start_date)
    df4 = df4.loc[mask4]
    
    l1 = df4[df4['barcode_data'].str.startswith('F0', na=False)]
    g1 = l1.groupby([pd.Grouper(key='date', freq='H'),l1.line_id]).size().reset_index(name='Missed/Hour')
    
    yo = g1.loc[g1['date']=='00:00']
    yo1 = g1.loc[g1['date']=='05:00']
    yo2 = g1.loc[g1['date']=='06:00']
    yo3 = g1.loc[g1['date']=='07:00']
    yo4 = g1.loc[g1['date']=='08:00']
    yo5 = g1.loc[g1['date']=='09:00']
    yo6 = g1.loc[g1['date']=='10:00']
    yo7 = g1.loc[g1['date']=='11:00']
    yo8 = g1.loc[g1['date']=='12:00']
    yo9 = g1.loc[g1['date']=='13:00']

  
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    
    try:
        bo = bo['Units/Hour'].item()
        worksheet.update("J3", bo, raw=False)
    except:
        set_with_dataframe(worksheet, bo, row=3, col=11, include_column_header=False)
    try:
        bo1 = bo1['Units/Hour'].item()
        worksheet.update("J4", bo1, raw=False)
    except:
        set_with_dataframe(worksheet, bo1, row=4, col=11, include_column_header=False)
    try:
        bo2 = bo2['Units/Hour'].item()
        worksheet.update("J5", bo2, raw=False)
    except:
        set_with_dataframe(worksheet, bo2, row=5, col=11, include_column_header=False)
    try:
        bo3 = bo3['Units/Hour'].item()
        worksheet.update("J6", bo3, raw=False)
    except:
        set_with_dataframe(worksheet, bo3, row=6, col=11, include_column_header=False)
    try:
        bo4 = bo4['Units/Hour'].item()
        worksheet.update("J7", bo4, raw=False)
    except:
        set_with_dataframe(worksheet, bo4, row=7, col=11, include_column_header=False)
    try:
        bo5 = bo5['Units/Hour'].item()
        worksheet.update("J8", bo5, raw=False)
    except:
        set_with_dataframe(worksheet, bo5, row=8, col=11, include_column_header=False)
    try:
        bo6 = bo6['Units/Hour'].item()
        worksheet.update("J8", bo6, raw=False)
    except:
        set_with_dataframe(worksheet, bo6, row=9, col=11, include_column_header=False)
    try:
        bo7 = bo7['Units/Hour'].item()
        worksheet.update("J10", bo7, raw=False)
    except:
        set_with_dataframe(worksheet, bo7, row=10, col=11, include_column_header=False)
    try:
        bo8 = bo8['Units/Hour'].item()
        worksheet.update("J11", bo8, raw=False)
    except:
        set_with_dataframe(worksheet, bo8, row=11, col=11, include_column_header=False)
    try:
        bo9 = bo9['Units/Hour'].item()
        worksheet.update("J12", bo9, raw=False)
    except:
        set_with_dataframe(worksheet, bo9, row=12, col=11, include_column_header=False)

def line5Kit():
    wait = time.sleep(5)
     #group line 6 by hour
    group = df.groupby([pd.Grouper(key='date', freq='H'),df.line_id]).size().reset_index(name='Units/Hour')
    df1 = group[group.line_id == 'KITTILELINE5']
    #format datetime to 00:00
    df1['date'] = df1['date'].dt.strftime('%H:%M')
    #Find Individual Times
    yo = df1.loc[df1['date']=='00:00']
    yo1 = df1.loc[df1['date']=='05:00']
    yo2 = df1.loc[df1['date']=='06:00']
    yo3 = df1.loc[df1['date']=='07:00']
    yo4 = df1.loc[df1['date']=='08:00']
    yo5 = df1.loc[df1['date']=='09:00']
    yo6 = df1.loc[df1['date']=='10:00']
    yo7 = df1.loc[df1['date']=='11:00']
    yo8 = df1.loc[df1['date']=='12:00']
    yo9 = df1.loc[df1['date']=='13:00']

    #drop date and lin_id columns
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    set_with_dataframe(worksheet, bo, row=3, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo1, row=4, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo2, row=5, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo3, row=6, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo4, row=7, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo5, row=8, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo6, row=9, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo7, row=10, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo8, row=11, col=12, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo9, row=12, col=12, include_column_header=False)

def missed5():
    df5 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line5-qc&page_size=2000&offset=0')[0]
    
    df5['date'] = pd.to_datetime(df5['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    
    
    mask5 = (df5['date'] > start_date)
    df5 = df5.loc[mask5]
    
    l1 = df5[df5['barcode_data'].str.startswith('F0', na=False)]
    g1 = l1.groupby([pd.Grouper(key='date', freq='H'),l1.line_id]).size().reset_index(name='Missed/Hour')
    
    yo = g1.loc[g1['date']=='00:00']
    yo1 = g1.loc[g1['date']=='05:00']
    yo2 = g1.loc[g1['date']=='06:00']
    yo3 = g1.loc[g1['date']=='07:00']
    yo4 = g1.loc[g1['date']=='08:00']
    yo5 = g1.loc[g1['date']=='09:00']
    yo6 = g1.loc[g1['date']=='10:00']
    yo7 = g1.loc[g1['date']=='11:00']
    yo8 = g1.loc[g1['date']=='12:00']
    yo9 = g1.loc[g1['date']=='13:00']

  
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    
    try:
        bo = bo['Units/Hour'].item()
        worksheet.update("K3", bo, raw=False)
    except:
        set_with_dataframe(worksheet, bo, row=3, col=13, include_column_header=False)
    try:
        bo1 = bo1['Units/Hour'].item()
        worksheet.update("K4", bo1, raw=False)
    except:
        set_with_dataframe(worksheet, bo1, row=4, col=13, include_column_header=False)
    try:
        bo2 = bo2['Units/Hour'].item()
        worksheet.update("K5", bo2, raw=False)
    except:
        set_with_dataframe(worksheet, bo2, row=5, col=13, include_column_header=False)
    try:
        bo3 = bo3['Units/Hour'].item()
        worksheet.update("K6", bo3, raw=False)
    except:
        set_with_dataframe(worksheet, bo3, row=6, col=13, include_column_header=False)
    try:
        bo4 = bo4['Units/Hour'].item()
        worksheet.update("K7", bo4, raw=False)
    except:
        set_with_dataframe(worksheet, bo4, row=7, col=13, include_column_header=False)
    try:
        bo5 = bo5['Units/Hour'].item()
        worksheet.update("K8", bo5, raw=False)
    except:
        set_with_dataframe(worksheet, bo5, row=8, col=13, include_column_header=False)
    try:
        bo6 = bo6['Units/Hour'].item()
        worksheet.update("K8", bo6, raw=False)
    except:
        set_with_dataframe(worksheet, bo6, row=9, col=13, include_column_header=False)
    try:
        bo7 = bo7['Units/Hour'].item()
        worksheet.update("K10", bo7, raw=False)
    except:
        set_with_dataframe(worksheet, bo7, row=10, col=13, include_column_header=False)
    try:
        bo8 = bo8['Units/Hour'].item()
        worksheet.update("K11", bo8, raw=False)
    except:
        set_with_dataframe(worksheet, bo8, row=11, col=13, include_column_header=False)
    try:
        bo9 = bo9['Units/Hour'].item()
        worksheet.update("K12", bo9, raw=False)
    except:
        set_with_dataframe(worksheet, bo9, row=12, col=13, include_column_header=False)

def line6Kit():
    wait = time.sleep(5)
     #group line 6 by hour
    group = df.groupby([pd.Grouper(key='date', freq='H'),df.line_id]).size().reset_index(name='Units/Hour')
    df1 = group[group.line_id == 'KITTILELINE6']
    #format datetime to 00:00
    df1['date'] = df1['date'].dt.strftime('%H:%M')
    #Find Individual Times
    yo = df1.loc[df1['date']=='00:00']
    yo1 = df1.loc[df1['date']=='05:00']
    yo2 = df1.loc[df1['date']=='06:00']
    yo3 = df1.loc[df1['date']=='07:00']
    yo4 = df1.loc[df1['date']=='08:00']
    yo5 = df1.loc[df1['date']=='09:00']
    yo6 = df1.loc[df1['date']=='10:00']
    yo7 = df1.loc[df1['date']=='11:00']
    yo8 = df1.loc[df1['date']=='12:00']
    yo9 = df1.loc[df1['date']=='13:00']

    #drop date and lin_id columns
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    set_with_dataframe(worksheet, bo, row=3, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo1, row=4, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo2, row=5, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo3, row=6, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo4, row=7, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo5, row=8, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo6, row=9, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo7, row=10, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo8, row=11, col=14, include_column_header=False)
    wait
    set_with_dataframe(worksheet, bo9, row=12, col=14, include_column_header=False)
    wait

def missed6():
    df6 = pd.read_html('https://rfa.tile.hiveplatform.org/bc_scan_event?cell_id=tile-line6-qc&page_size=2000&offset=0')[0]
    
    df6['date'] = pd.to_datetime(df6['scan_event_timestamp'], format="%Y-%m-%d %H:%M:%S")
    
    
    mask6 = (df6['date'] > start_date)
    df6 = df6.loc[mask6]
    
    l1 = df6[df6['barcode_data'].str.startswith('F0', na=False)]
    g1 = l1.groupby([pd.Grouper(key='date', freq='H'),l1.line_id]).size().reset_index(name='Missed/Hour')
    
    yo = g1.loc[g1['date']=='00:00']
    yo1 = g1.loc[g1['date']=='05:00']
    yo2 = g1.loc[g1['date']=='06:00']
    yo3 = g1.loc[g1['date']=='07:00']
    yo4 = g1.loc[g1['date']=='08:00']
    yo5 = g1.loc[g1['date']=='09:00']
    yo6 = g1.loc[g1['date']=='10:00']
    yo7 = g1.loc[g1['date']=='11:00']
    yo8 = g1.loc[g1['date']=='12:00']
    yo9 = g1.loc[g1['date']=='13:00']

  
    bo = yo.drop(['date','line_id'], axis=1)
    bo1 = yo1.drop(['date','line_id'], axis=1)
    bo2 = yo2.drop(['date','line_id'], axis=1)
    bo3 = yo3.drop(['date','line_id'], axis=1)
    bo4 = yo4.drop(['date','line_id'], axis=1)
    bo5 = yo5.drop(['date','line_id'], axis=1)
    bo6 = yo6.drop(['date','line_id'], axis=1)
    bo7 = yo7.drop(['date','line_id'], axis=1)
    bo8 = yo8.drop(['date','line_id'], axis=1)
    bo9 = yo9.drop(['date','line_id'], axis=1)

    
    
    try:
        bo = bo['Units/Hour'].item()
        worksheet.update("N3", bo, raw=False)
    except:
        set_with_dataframe(worksheet, bo, row=3, col=15, include_column_header=False)
    try:
        bo1 = bo1['Units/Hour'].item()
        worksheet.update("N4", bo1, raw=False)
    except:
        set_with_dataframe(worksheet, bo1, row=4, col=15, include_column_header=False)
    try:
        bo2 = bo2['Units/Hour'].item()
        worksheet.update("N5", bo2, raw=False)
    except:
        set_with_dataframe(worksheet, bo2, row=5, col=15, include_column_header=False)
    try:
        bo3 = bo3['Units/Hour'].item()
        worksheet.update("N6", bo3, raw=False)
    except:
        set_with_dataframe(worksheet, bo3, row=6, col=15, include_column_header=False)
    try:
        bo4 = bo4['Units/Hour'].item()
        worksheet.update("N7", bo4, raw=False)
    except:
        set_with_dataframe(worksheet, bo4, row=7, col=15, include_column_header=False)
    try:
        bo5 = bo5['Units/Hour'].item()
        worksheet.update("N8", bo5, raw=False)
    except:
        set_with_dataframe(worksheet, bo5, row=8, col=15, include_column_header=False)
    try:
        bo6 = bo6['Units/Hour'].item()
        worksheet.update("N9", bo6, raw=False)
    except:
        set_with_dataframe(worksheet, bo6, row=9, col=15, include_column_header=False)
    try:
        bo7 = bo7['Units/Hour'].item()
        worksheet.update("N10", bo7, raw=False)
    except:
        set_with_dataframe(worksheet, bo7, row=10, col=15, include_column_header=False)
    try:
        bo8 = bo8['Units/Hour'].item()
        worksheet.update("N11", bo8, raw=False)
    except:
        set_with_dataframe(worksheet, bo8, row=11, col=15, include_column_header=False)
    try:
        bo9 = bo9['Units/Hour'].item()
        worksheet.update("N12", bo9, raw=False)
    except:
        set_with_dataframe(worksheet, bo9, row=12, col=15, include_column_header=False)

def copy():
    chart3 = ['B16:B25','D16:D25', 'F16:F25', 'H16:H25', 'J16:J25', 'L16:L25']
    chart2 = ['D3:D12','F3:F12', 'H3:H12', 'J3:J12', 'L3:L12', 'N3:N12']
    for i,u in zip(chart2, chart3):
        s_range = sh.worksheet(title).get(i)
        range_copy = s_range.copy()
        worksheet.update(u, range_copy)


def findSumKit():
    sum_range = ["D13",'F13','H13','J13','L13','N13']
    sum_list = ['(D3:D12)', '(F3:F12)', '(H3:H12)', '(J3:J12)', '(L3:L12)', '(N3:N12)']
    for i, j in zip(sum_range,sum_list):
        b=[]
        b = ("=SUM"+j)
        worksheet.update(i,b, raw=False)

def findSumQC():
    sum_range = ["E13",'G13','I13','K13','M13','O13']
    sum_list = ['(E3:E12)', '(G3:G12)', '(I3:I12)', '(K3:K12)', '(M3:M12)', '(O3:O12)']
    for i, j in zip(sum_range,sum_list):
        b=[]
        b = ("=SUM"+j)
        worksheet.update(i,b, raw=False)


def findAvg():
    #Write formula for Average Kit
    worksheet.update('B3', "=Iferror(AVERAGE(D3,F3,H3,J3,L3,N3))", raw=False)
    worksheet.update('B4', "=Iferror(AVERAGE(D4,F4,H4,J4,L4,N4))", raw=False)
    worksheet.update('B5', "=Iferror(AVERAGE(D5,F5,H5,J5,L5,N5))", raw=False)
    worksheet.update('B6', "=Iferror(AVERAGE(D6,F6,H6,J6,L6,N6))", raw=False)
    worksheet.update('B7', "=Iferror(AVERAGE(D7,F7,H7,J7,L7,N7))", raw=False)
    worksheet.update('B8', "=Iferror(AVERAGE(D8,F8,H8,J8,L8,N8))", raw=False)
    worksheet.update('B9', "=Iferror(AVERAGE(D9,F3,H9,J9,L9,N9))", raw=False)
    worksheet.update('B10', "=Iferror(AVERAGE(D10,F10,H10,J10,L10,N10))", raw=False)
    worksheet.update('B11', "=Iferror(AVERAGE(D11,F11,H11,J11,L11,N11))", raw=False)
    worksheet.update('B12', "=Iferror(AVERAGE(D12,F12,H12,J12,L12,N12))", raw=False)

    g = ['B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12']
    for k in g:
        worksheet.format(k, { "numberFormat": { "type": "NUMBER","pattern": "##.##" }})
def findQCperc():
    #Write formula for Average Kit
    worksheet.update('C3', "=Iferror(SUM(E3,G3,I3,K3,M3,O3)/B3)", raw=False)
    worksheet.update('C4', "=Iferror(SUM(E4,G4,I4,K4,M4,O4)/B4)", raw=False)
    worksheet.update('C5', "=Iferror(SUM(E5,G5,I5,K5,M5,O5)/B5)", raw=False)
    worksheet.update('C6', "=Iferror(SUM(E6,G6,I6,K6,M6,O6)/B6)", raw=False)
    worksheet.update('C7', "=Iferror(SUM(E7,G7,I7,K7,M7,O7)/B7)", raw=False)
    worksheet.update('C8', "=Iferror(SUM(E8,G8,I8,K8,M8,O8)/B8)", raw=False)
    worksheet.update('C9', "=Iferror(SUM(E9,G3,I9,K9,M9,O9)/B9)", raw=False)
    worksheet.update('C10', "=Iferror(SUM(E10,G10,I10,K10,M10,O10)/B10)", raw=False)
    worksheet.update('C11', "=Iferror(SUM(E11,G11,I11,K11,M11,O11)/B11)", raw=False)
    worksheet.update('C12', "=Iferror(SUM(E12,G12,I12,K12,M12,O12)/B12)", raw=False)

    g = ['C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12']
    for k in g:
        worksheet.format(k, { "numberFormat": { "type": "PERCENT","pattern": "##.##%" }})


def findQCtwo():
    a = ["C16:C25","E16:E25","G16:G25","I16:I25","K16:K25","M16:M25"]

    b = ["C16","C17", "C18", "C19","C20", "C21", "C22", "C23", "C24", "C25",
            "E16","E17", "E18", "E19","E20", "E21", "E22", "E23", "E24", "E25",
            "G16","G17", "G18", "G19","G20", "G21", "G22", "G23", "G24", "G25",
            "I16","I17", "I18", "I19","I20", "I21", "I22", "I23", "I24", "I25",
            "K16","K17", "K18", "K19","K20", "K21", "K22", "K23", "K24", "K25",
            "M16","M17", "M18", "M19","M20", "M21", "M22", "M23", "M24", "M25"]

    c =["=Iferror(E3/D3)", "=Iferror(E4/D4)", "=Iferror(E5/D5)", "=Iferror(E6/D6)", "=Iferror(E7/D7)","=Iferror(E8/D8)", "=Iferror(E9/D9)", "=Iferror(E10/D10)", "=Iferror(E11/D11)", "=Iferror(E12/D12)",
        "=Iferror(G3/F3)", "=Iferror(G4/F4)", "=Iferror(G5/F5)", "=Iferror(G6/F6)", "=Iferror(G7/F7)","=Iferror(G8/F8)", "=Iferror(G9/F9)", "=Iferror(G10/F10)", "=Iferror(G11/F11)", "=Iferror(G12/F12)",
        "=Iferror(I3/H3)", "=Iferror(I4/H4)", "=Iferror(I5/H5)", "=Iferror(I6/H6)", "=Iferror(I7/H7)","=Iferror(I8/H8)", "=Iferror(I9/H9)", "=Iferror(I10/H10)", "=Iferror(I11/H11)", "=Iferror(I12/H12)",
        "=Iferror(K3/J3)", "=Iferror(K4/J4)", "=Iferror(K5/J5)", "=Iferror(K6/J6)", "=Iferror(K7/J7)","=Iferror(K8/J8)", "=Iferror(K9/J9)", "=Iferror(K10/J10)", "=Iferror(K11/J11)", "=Iferror(K12/J12)",
        "=Iferror(M3/L3)", "=Iferror(M4/L4)", "=Iferror(M5/L5)", "=Iferror(M6/L6)", "=Iferror(M7/L7)","=Iferror(M8/L8)", "=Iferror(M9/L9)", "=Iferror(M10/L10)", "=Iferror(M11/L11)", "=Iferror(M12/L12)",
        "=Iferror(O3/N3)", "=Iferror(O4/N4)", "=Iferror(O5/N5)", "=Iferror(O6/N6)", "=Iferror(O7/N7)","=Iferror(O8/N8)", "=Iferror(O9/N9)", "=Iferror(O10/N10)", "=Iferror(O11/N11)", "=Iferror(O12/N12)"]

    time.sleep(30)
    for i, u in zip(b,c):
        worksheet.update(i, u, raw=False)
    time.sleep(45)
    for j in (a):
        worksheet.format(a, { "numberFormat": { "type": "PERCENT","pattern": "##.##%" }})

print("Writing Sum Kit Formula for Chart 1")
findSumKit()
print("Done")
# function call
print("Writing Sum QC Formula for Chart 1 in: ")
countdown(int(t))
findSumQC()
print("Done")
# function call
print("Writing Average Formula for Chart 1 in: ")
countdown(int(t))
findAvg()
print("Done")
# function call
print("Writing QC Percent Formula for Chart 1 in: ")
countdown(int(t))
findQCperc()
print("Done")
# function call
print("Writing QC Percent Formula for Chart 2 in: ")
countdown(int(t))
findQCtwo()
print("Done")
print("Writing Line 1 Kit Data in: ")
countdown(int(t))
line1Kit()
print("Done")
print("Writing Line 1 QC Data in: ")
countdown(int(t))
missed1()
print("Done")
print("Writing Line 2 Kit Data in: ")
countdown(int(t))
line2Kit()
print("Done")
print("Writing Line 2 QC Data in: ")
countdown(int(t))
missed2()
print("Done")
print("Writing Line 3 Kit Data in: ")
countdown(int(t))
line3Kit()
print("Done")
print("Writing Line 3 QC Data in: ")
countdown(int(t))
missed3()
print("Done")
print("Writing Line 4 Kit Data in: ")
countdown(int(t))
line4Kit()
print("Done")
print("Writing Line 4 QC Data in: ")
countdown(int(t))
missed4()
print("Done")
print("Writing Line 5 Kit Data in: ")
countdown(int(t))
line5Kit()
print("Done")
print("Writing Line 5 QC Data in: ")
countdown(int(t))
missed5()
print("Done")
print("Writing Line 6 Kit Data in: ")
countdown(int(t))
line6Kit()
print("Done")
print("Writing Line 6 QC Data in: ")
countdown(int(t))
missed6()
print("Done")
print("Copying data from Chart 1 to Chart 2")
countdown(int(t))
copy()
print("Done")
print("ProductionRecap has Finished")






