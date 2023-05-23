import pandas as pd
import numpy as np
import re
import datetime
from dateutil.relativedelta import relativedelta
from datetime import date


# Clean up raw data imported
df = pd.read_csv(r'C:/Users/sean/Desktop/Aging Raw Data CSV.csv', engine='python', encoding= 'unicode_escape')
df.columns = ['Item', 'Acq Date', 'Asset Value', 'Qty', 'Disb Date']
df = df.drop(labels=0, axis=0)
df['Item'] = df['Item'].str.split('(').str[0]
df['Item'] = df['Item'].str.rstrip()
df['Item'] = df['Item'].str.replace('*','', regex=True)
df['Item'].ffill(inplace=True)


# Drop all rows that have blanks besides item name and Disb Date, then sort items alphabetically
df.dropna(subset=['Acq Date','Asset Value', 'Qty'], inplace = True)
df= df.sort_values(by=['Item'])

# Convert Dates to proper format for comparisons later
df['Disb Date']= pd.to_datetime(df['Disb Date'])
df['Acq Date']= pd.to_datetime(df['Acq Date'])


# Create a list of last invoice (li) for each item
li = df[['Item','Acq Date']].copy()
li = li.groupby('Item')['Acq Date'].max()
last_invoice = pd.DataFrame(li)
last_invoice.columns = ['Last Invoice']


# Create Aged Inventory table (ai) which isolates all entries that have NaN as Disb Date (haven't been sold yet)
ai = df[df['Disb Date'].isnull()].copy()
ai = pd.DataFrame(ai)


# Get difference between Acq Date and today in days to determine aging
ai['Acq Date']= pd.to_datetime(ai['Acq Date'])
get_today = pd.to_datetime(datetime.date.today())
ai['# of Days'] = get_today - ai['Acq Date']
ai['# of Days'] = ai['# of Days'] / np.timedelta64(1,'D')


# Drop the Disb Date column for Aged Inventory table
ai.drop('Disb Date', axis=1, inplace=True)


# Create aging brackets and assign each row a value
aging_bins = [
    (ai['# of Days'] <= 30),
    (ai['# of Days'] > 30) & (ai['# of Days'] <=60),
    (ai['# of Days'] > 60) & (ai['# of Days'] <=90),
    (ai['# of Days'] > 90) & (ai['# of Days'] <=120),
    (ai['# of Days'] > 120) & (ai['# of Days'] <=180),
    (ai['# of Days'] > 180)
    ]

labels = ['0-30 Days', '31-60 Days', '61-90 Days', '91-120 Days', '121-180 Days', '180+ Days']

ai['Aging'] = np.select(aging_bins, labels)


# Create new dataframe from import of "Open Sales" item report
open_sales = pd.read_csv(r'C:/Users/sean/Desktop/Item Info List.csv', engine='python', encoding= 'unicode_escape')
open_sales.columns = ['Type', 'Item', 'On Hand', 'On Purch', 'On Sales', 'Price']
open_sales['Item'] = open_sales['Item'].str.split('(').str[0]
open_sales['Item'] = open_sales['Item'].str.rstrip()
open_sales['Item'] = open_sales['Item'].str.replace('*','', regex=True)


# Drop all rows that are in the types listed, only looking for Inventory Part and Inventory Assembly
open_sales = open_sales.drop(open_sales.index[open_sales['Type'].isin(['Service', 'Non-inventory Part', 'Other Charge'])])


# Drop the Type column since no longer needed
open_sales.drop('Type', axis=1, inplace=True)

# Create a notes table to merge into final table as well
notes = pd.read_csv(r'C:/Users/sean/Desktop/Notes.csv', engine='python', encoding= 'unicode_escape')
notes.columns = ['Item', 'Notes']
notes['Item'] = notes['Item'].str.split('(').str[0]
notes['Item'] = notes['Item'].str.rstrip()
notes['Item'] = notes['Item'].str.replace('*','', regex=True)
notes['Notes'] = notes['Notes'].fillna('')
notes = pd.DataFrame(notes)


#Merge the dataframes based on the item name to form Final Table (ft)
frst = open_sales.merge(notes, on='Item')
temp = last_invoice.merge(frst, on='Item')
ft = ai.merge(temp, on='Item')


# Add retail value column and cbfilter for Chris & Brooke customers
ft['Retail Value'] = ft['Qty'] * ft['Price']
ft['cbfilter'] = ft['Item'].str.split('-').str[0]


# Format all columns to display correctly
# This causes a future warning but seems like a bug from all reports online
ft.loc[:, 'Acq Date'] = ft['Acq Date'].dt.strftime('%m/%d/%Y')
ft.loc[:, 'Qty'] = ft['Qty'].map('{:,.0f}'.format)
ft.loc[:, 'Asset Value'] ='$'+ ft['Asset Value'].map('{:,.2f}'.format)
ft.loc[:, 'Retail Value'] ='$'+ ft['Retail Value'].map('{:,.2f}'.format)
ft.loc[:, '# of Days'] = ft['# of Days'].map('{:,.0f}'.format)
ft.loc[:, 'Last Invoice'] = ft['Last Invoice'].dt.strftime('%m/%d/%Y')
ft.loc[:, 'On Hand'] = ft['On Hand'].map('{:,.0f}'.format)
ft.loc[:, 'On Purch'] = ft['On Purch'].map('{:,.0f}'.format)
ft.loc[:, 'On Sales'] = ft['On Sales'].map('{:,.0f}'.format)
ft.loc[:, 'Price'] ='$'+ ft['Price'].map('{:,.4f}'.format)


# Get the three final tables, dana, cb, and qty
dana = ft[['Item', 'Acq Date', 'Qty', 'Asset Value', 'Retail Value', '# of Days', 'Aging']]

cb_temp = ft[ft['cbfilter'].isin(['21ST', 'AE', 'AEP', 'AKRN', 'ALT', 'AM', 'AMN', 'AMR', 'ANT',
                            'ASCP', 'AVA', 'BIO', 'BOTT', 'CON', 'FAL', 'FRA', 'GEM', 'LNK',
                            'LRM', 'NP', 'NV', 'PIO', 'PLD', 'PREM', 'SAPT', 'SCI', 'SHA', 'SN',
                            'SRI', 'STOCK', 'SVT', 'TCL', 'TISH', 'VIT', 'VW', 'WLNG'])]

cb = cb_temp[['Item', 'Acq Date', 'Qty', 'Asset Value', 'Retail Value', '# of Days', 'Aging']]


temp_qty = ft.pivot_table(index=['Item'], columns=['Aging'], values=['Qty'], aggfunc='sum', fill_value=0,)
qty =  temp_qty.reindex(['0-30 Days', '31-60 Days', '61-90 Days', '91-120 Days', '121-180 Days', '180+ Days'], axis=1, level=1)

#print(qty)
with pd.ExcelWriter(r'C:/Users/sean/Desktop/scriptexport.xlsx', engine='xlsxwriter') as writer:
    dana.to_excel(writer, sheet_name='Dana')
    cb.to_excel(writer, sheet_name='C&B')
workbook = writer.book
worksheet = writer.sheets['Dana']

cellFormat = workbook.add_format({'num_format': '#,##'})
worksheet.set_column('C:C', 10, cellFormat)

    #qty.to_excel(writer, sheet_name='Qty')



#print(cb.head(10))
#print(ai.tail(25))
