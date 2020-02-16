#!/usr/bin/env python
# coding: utf-8
# By Grand Chang. in 2020 CNY. Use for download BT_Billing data .xls (but HTML format)
# Combine Sold data and Sales name to povit and charting data to out :
# Daily shipped Qty, Top sell models, Top Sales, country, models.
# Download file by selenium and save to path //daily_billing 
# Some new items are not listed on itemno_cat.xlsx
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
import matplotlib.dates as mdates
import numpy as np
import datetime
import sys
import pathlib
import os

# filename = '%s' % filename
filename =input('input file name just downloaded: ')
# currentPath= pathlib.Path().absolute()

bt_bill=pd.read_html(filename)
billdata=pd.DataFrame(bt_bill[0])

# salesName = pd.DataFrame(pd.read_excel('SalesList12312019.xlsx')).drop(['Team','Sales Forecast Y/N','Group leader','Sales','Head count  by team','Head count by group','in SJ office','Location','Current month Hire'],axis=1)
itemno_cat= pd.DataFrame(pd.read_excel('itemno_cat.xlsx'))
solddata=billdata.drop(['Order/DN Num','Invoice','Customer','ZDSR','Order Type','Type'], axis=1)

# Merge 2 tables for Sales ID (code) to find sales name that easy to know. (May also use .join method)
solddata = solddata.merge(itemno_cat, how='left', left_on='Itemno', right_on='Itemno')

naCat = solddata[solddata['Cat'].isnull()]
naCat = naCat.drop(['Date','Qty','Node Qty','Sales','Ship To'], axis=1)

print(naCat)

print('=============================================')
print("Up to date total sold Systems Qty: ", solddata['Qty'].sum())
print("Up to date total sold Nodes total: ", solddata['Node Qty'].sum())
print("==============================================")

naCat.to_excel('itemno_new.xlsx')

