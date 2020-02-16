#!/usr/bin/env python
# coding: utf-8
# By Grand Chang. in 2020 CNY. Use for download BT_Billing data .xls (but HTML format)
# Combine Sold data and Sales name to povit and charting data to out :
# Daily shipped Qty, Top sell models, Top Sales, country, models.
# Download file by selenium and save to path //daily_billing 
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import matplotlib.dates as mdates
import numpy as np
import datetime
import sys
import pathlib
import os
# download bt_billing and save to daily_billing


from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from time import sleep


options = Options()
options.add_experimental_option("prefs",{"download.default_directory": r"D:\coding\daily_billing"})

driver = webdriver.Chrome(chrome_options=options)
# launch chrome to open the following URL 
url = 'http://mpserver6/mrpbt/BtBillingQueryByRef.aspx'
# url = "http://mpserver/MrpNew/BoQueryTwn.aspx" 
driver.get(url)
# input the 88 in Item/PI
# item_pi = driver.find_element_by_name("txtItem")
# item_pi.send_keys('88')

# Click Display button to show latest data 
driver.find_element_by_name('btnDisplay').click()

# Timeout to wait for query 
sleep(5)

# 43~45 Click “Display in Excel” to save the query result as BoQueryTW.aspx
driver.find_element_by_name('btnExport').click() 
sleep(10)
driver.close()
# finish download

# run script with filename that data input (49~56)
# try:
#     filename = sys.argv[1]
# except:
#     print ("\n- Fail, you must pass in excel file name xxxx.xlsx")
#     sys.exit()


# filename = '%s' % filename
filename =input('input file name just downloaded: ')
currentPath= pathlib.Path().absolute()

bt_bill=pd.read_html(filename)
billdata=pd.DataFrame(bt_bill[0])
# print(billdata)
# Read html file (downlad named as xxx.xls but actually is HTML). Not necessary -Skip first row as excel file but use "[0]". 
# Read excel file either xls or xlsx. Skip first row to get correct title. 

salesName = pd.DataFrame(pd.read_excel('SalesList12312019.xlsx')).drop(['Team','Sales Forecast Y/N','Group leader','Sales','Head count  by team','Head count by group','in SJ office','Location','Current month Hire'],axis=1)
itemno_cat= pd.DataFrame(pd.read_excel('itemno_cat.xlsx'))
solddata=billdata.drop(['Order/DN Num','Invoice','Customer','ZDSR','Order Type','Type'], axis=1)

# Write dataframe to new .xls file in new dir(currenet+newfilename)

newfile = 'GC%s' %filename
# print("New File Name: ",newfile)
newDirName= filename[:-4]
# print("New Dir Name:",newDirName)

currPath=pathlib.Path().absolute()

print("Current Path Name: ",currPath)

newPath=f"{currPath}/{newDirName}/{newfile}"
# newPath=f"{newDirName}/{newfile}"

print("New Path and file name is: ",newPath)

# Create New Dir for store new file in new Path.
newDir=f"{currPath}/{newDirName}"

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

createFolder(newDir)

writer = pd.ExcelWriter(newPath)
billdata.to_excel(writer, sheet_name= 'Raw_Data',na_rep=False,index=False,header=True)

solddata= solddata.merge(salesName,how='left',left_on='Sales',right_on='Rep code')
# Merge 2 tables for Sales ID (code) to find sales name that easy to know. (May also use .join method)
solddata = solddata.merge(itemno_cat, how='left', left_on='Itemno', right_on='Itemno')

solddata['Date']= pd.to_datetime(solddata['Date'].astype(str), format='%Y%m%d')
solddata['Date']= solddata['Date'].apply(lambda x:x.strftime('%Y-%m-%d'))
# Covert Date type from object to datetime and Change Date format to be 2020-01-01

# Write new dataframe to sheet name= New_DF;

solddata.to_excel(writer, sheet_name= 'New_DF',na_rep=False,index=True,header=True)

# Summary each day shipment amount;
# solddata.groupby(['Date']).sum().plot(kind='bar',x='Date', figsize=(12,6), fontsize= 7, rot=45, title='Daily Shipment')
# solddata.groupby(['Date']).agg(sum).reset_index().plot(kind='bar',x='Date', figsize=(8,6), fontsize= 7, rot=45, title='Daily Shipment')

dailyShip=solddata.groupby(['Date']).agg(sum).reset_index()
temp=dailyShip[['Qty','Node Qty']]

# dailySum=dailyShip.append(dailyShip.sum(numeric_only=True), ignore_index=True)
dailySum= dailyShip[['Qty','Node Qty']].sum()
dailySum['Date']= 'UptoDate Total'
dailyShip = dailyShip.append(dailySum,ignore_index=True)
dailyShip.to_excel(writer, sheet_name= 'by_Date', index=False,header=True)

# .plot(kind='bar',x='Date', figsize=(8,6), fontsize= 7, rot=75, title='Daily Shipment')


# plt.subplots_adjust(bottom = 0.2)

dailyShip.plot(kind='bar',x='Date', figsize=(8,6), fontsize= 7, rot=75, title='Daily Shipment')
plt.tight_layout()
plt.savefig(f'{newDir}/dailyship%s.png' %filename)

plt.show()

# print(solddata.groupby('Itemno')['Qty','Node Qty'].agg(sum).nlargest(30,'Qty'))

# List Model name by_Itemno and sum up sold Qty (System) and Node count;
solddata.groupby('Itemno')['Qty','Node Qty'].agg(sum).nlargest(20,'Qty').plot(kind='barh',title='Top 20 Models', figsize=(8,6),fontsize=7)
plt.tight_layout()
plt.savefig(f'{newDir}/Top20Model_%s.png' %filename)
plt.show()


# plot (by_ShipTo)
# solddata.groupby('Ship To')['Qty','Node Qty'].agg(sum).nlargest(10,'Qty').plot.pie(autopct='%.2f', fontsize=8, figsize=(10,6), legend=False, subplots=True)
by_ShipTo=solddata.groupby('Ship To')['Qty','Node Qty'].agg(sum).sort_values(by='Qty',ascending=True)
# print(by_ShipTo)

by_ShipTo['Qty'].plot.pie(autopct='%.2f', fontsize=8, figsize=(6,4), legend=False)
plt.title('By Country Sold Amount',fontsize=18, fontweight='bold')
plt.ylabel('By System Count', fontsize=12)
plt.savefig(f'{newDir}/ShipTo_%s_bySys.png' % filename)
# plt.show()

by_ShipTo['Node Qty'].plot.pie(autopct='%.2f', fontsize=8, figsize=(6,4), legend=False, subplots=False)
plt.title('By Country Sold Amount',fontsize=18, fontweight='bold')
plt.savefig(f'{newDir}/ShipTo_%s_byNode.png' % filename)
plt.ylabel('By Node Count', fontsize=12)
# plt.show()



# List Sales name and sum up sold Qty (System) and Node count;
solddata.groupby('Name')['Qty','Node Qty'].agg(sum).nlargest(20,'Qty').plot(kind='barh',figsize=(8,5))
plt.title('Top 20 Sales', fontsize=18, fontweight='bold')
plt.tight_layout()
plt.savefig(f'{newDir}/Top Sale_%s.png' %filename)
# plt.show()

#By Country and Sales 
# by_sales=solddata.groupby(['Ship To','Sales']).Qty.sum()
by_sales_ship_item=solddata.groupby(['Name','Ship To','Itemno']).agg(sum)
# print(by_sales_ship_item)

# List Sales Name and what Itemno they sold.
# print(solddata.groupby(['Name','Itemno']).agg(sum))
sales_item=solddata.groupby(['Name','Itemno']).agg(sum)
sales_item.to_excel(writer, sheet_name= 'by_Sales_Item', index=True,header=True)

# Pivot method which list by Ship to country and Sales Name and summary them.
pv_contrySale = solddata.pivot_table(index=['Ship To','Name'],aggfunc=[np.sum])

# print(pv_contrySale)
pv_contrySale.to_excel(writer, sheet_name= 'by_Country_Sales', index=True,header=True)

# Test povit with 3 columns:
pv_ConSalIte = solddata.pivot_table(index=['Ship To','Name','Itemno'],aggfunc=[np.sum])
pv_IteConSal = solddata.pivot_table(index=['Itemno','Ship To','Name'],aggfunc=[np.sum])
pv_CatIteConSal = solddata.pivot_table(index=['Cat','Itemno','Ship To','Name'],aggfunc=[np.sum])

# print(pv_ConSalIte)
pv_ConSalIte.to_excel(writer, sheet_name= 'by_Country_Sales_Item', index=True,header=True)
pv_CatIteConSal.to_excel(writer, sheet_name= 'by_Cat_Item_Country_Sales', index=True,header=True)


writer.close()


pv_contrySale.plot(kind='bar',figsize=(8,6),subplots=True,rot=270, fontsize=6)
plt.subplots_adjust(bottom = 0.3)
plt.tight_layout()
plt.savefig(f'{newDir}/Sale_Country_%s.png' %filename)
# plt.show()

print("=============================================")
print("Up to date total sold Systems Qty: ", solddata['Qty'].sum())
print("Up to date total sold Nodes total: ", solddata['Node Qty'].sum())
print("=============================================")

# Write a file for daily update;
solddata.to_excel('daily_output.xlsx')


