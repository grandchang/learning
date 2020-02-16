#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime
import sys
import pathlib
import os

try:
    filename = sys.argv[1]
except:
    print ("\n- Fail, you must pass in excel file name xxxx.xlsx")
    sys.exit()


filename = '%s' % filename
currentPath= pathlib.Path().absolute()

# In[304]:


# billdata=pd.DataFrame(pd.read_html(filename))
bt_bill=pd.read_html(filename)
billdata=pd.DataFrame(bt_bill[0])
print(billdata)
# Read excel file either xls or xlsx. Skip first row to get correct title. 


# In[305]:


salesName = pd.DataFrame(pd.read_excel('SalesList12312019.xlsx')).drop(['Team','Sales Forecast Y/N','Group leader','Sales','Head count  by team','Head count by group','in SJ office','Location','Current month Hire'],axis=1)


# In[306]:

# In[307]:


solddata=billdata.drop(['Order/DN Num','Invoice','Customer','ZDSR','Order Type','Type'], axis=1)

# Write dataframe to new .xls file in new dir(currenet+newfilename)

newfile = 'GC%s' %filename
print("New File Name: ",newfile)
newDirName= filename[:-4]
print("New Dir Name:",newDirName)

currPath=pathlib.Path().absolute()

print("Current Path Name: ",currPath)

newPath=f"{currPath}/{newDirName}/{newfile}"
# newPath=f"{newDirName}/{newfile}"

print("New Path and file name: ",newPath)

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
billdata.to_excel(writer, sheet_name= 'Raw_Data',na_rep=False,index=True,header=True)

solddata= solddata.merge(salesName,how='left',left_on='Sales',right_on='Rep code')

solddata['Date']= pd.to_datetime(solddata['Date'].astype(str), format='%Y%m%d')
# Covert Date type from object to datetime and Change Date format to be 2020-01-01

print('-------------------------')
print('shape:', solddata.shape)
print("--------------------------")
print('data type: \n',solddata.dtypes)
print('---------------------------')


# In[308]:


solddata['Date']= solddata['Date'].apply(lambda x:x.strftime('%Y-%m-%d'))


# In[330]:


# Write new dataframe to sheet name= New_DF;

solddata.to_excel(writer, sheet_name= 'New_DF',na_rep=False,index=True,header=True)


# In[1]:


# solddata.groupby(['Date']).sum().plot(kind='bar',x='Date', figsize=(12,6), fontsize= 7, rot=45, title='Daily Shipment')
solddata.groupby(['Date']).agg(sum).reset_index().plot(kind='bar',x='Date', figsize=(10,7), fontsize= 7, rot=45, title='Daily Shipment')
plt.subplots_adjust(bottom = 0.2)
plt.tight_layout()
# plt.xticks(pd.date_range(start='2020-01-01', end='2020-01-01'))

plt.savefig(f'{newDir}/dailyship%s.png' %filename)
plt.show()

# In[277]:


print(solddata.groupby('Itemno')['Qty','Node Qty'].agg(sum).nlargest(30,'Qty'))

# In[275]:


# List Model name by_Itemno and sum up sold Qty (System) and Node count;
solddata.groupby('Itemno')['Qty','Node Qty'].agg(sum).nlargest(20,'Qty').plot(kind='barh',title='Top 20 Models', figsize=(14,7),fontsize=7)
plt.tight_layout()

# In[231]:

plt.savefig(f'{newDir}/Top20Model_%s.png' %filename)
plt.show()



# In[251]:


# plot (by_ShipTo)
# solddata.groupby('Ship To')['Qty','Node Qty'].agg(sum).nlargest(10,'Qty').plot.pie(autopct='%.2f', fontsize=8, figsize=(10,6), legend=False, subplots=True)
by_ShipTo=solddata.groupby('Ship To')['Qty','Node Qty'].agg(sum).sort_values(by='Qty',ascending=True)
print(by_ShipTo)

by_ShipTo['Qty'].plot.pie(autopct='%.2f', fontsize=8, figsize=(8,6), legend=False)
plt.title('By Country System Sold Amount',fontsize=18, fontweight='bold')
plt.savefig(f'{newDir}/ShipTo_%s_bySys.png' % filename)
plt.ylabel('By System Shipment Count', fontsize=12, fontweight='bold')
plt.tight_layout()
plt.show()

by_ShipTo['Node Qty'].plot.pie(autopct='%.2f', fontsize=8, figsize=(8,6), legend=False, subplots=False)

plt.title('By Country Sold Nodes Amount',fontsize=18, fontweight='bold')
plt.savefig(f'{newDir}/ShipTo_%s_byNode.png' % filename)
plt.ylabel('By Shipment Node Count', fontsize=12, fontweight='bold')
plt.tight_layout()
plt.show()


# In[329]:


# List Sales name and sum up sold Qty (System) and Node count;
solddata.groupby('Name')['Qty','Node Qty'].agg(sum).nlargest(20,'Qty').plot(kind='barh',title='Top 20 Sales',figsize=(14,7))
plt.tight_layout()
plt.savefig(f'{newDir}/Top Sale_%s.png' %filename)
plt.show()
# In[325]:


#By Country and Sales 
# by_sales=solddata.groupby(['Ship To','Sales']).Qty.sum()
by_sales_ship_item=solddata.groupby(['Name','Ship To','Itemno']).agg(sum)
# print(by_sales_ship_item)

# In[328]:


# print(solddata.groupby(['Name','Itemno']).agg(sum))
sales_item=solddata.groupby(['Name','Itemno']).agg(sum)
sales_item.to_excel(writer, sheet_name= 'by_Sales_Item', index=True,header=True)


# In[266]:


pv_contrySale = solddata.pivot_table(index=['Ship To','Name'],aggfunc=[np.sum])


# In[313]:


print(pv_contrySale)
pv_contrySale.to_excel(writer, sheet_name= 'by_Country_Sales', index=True,header=True)

# Test povit with 3 columns:
pv_ConSalIte = solddata.pivot_table(index=['Ship To','Name','Itemno'],aggfunc=[np.sum])


# In[313]:


print(pv_ConSalIte)
pv_ConSalIte.to_excel(writer, sheet_name= 'by_Country_Sales_Item', index=True,header=True)


writer.close()
# In[316]:


pv_contrySale.plot(kind='bar',figsize=(12,8),subplots=True)
plt.subplots_adjust(bottom = 0.4)
plt.savefig(f'{newDir}/Sale_Country_%s.png' %filename)
plt.tight_layout()
plt.show()

# In[ ]:
print(solddata.corr())
print("Upto data total sold Qty",solddata['Qty'].sum())
print("Upto data total sold Qty",solddata['Node Qty'].sum())



