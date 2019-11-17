# -*- coding: utf-8 -*-
"""
Created on Tue Aug 28 10:22:29 2018

@author: Rohit Nagpal
"""
import sys
sys.setrecursionlimit(1500)
from fuzzywuzzy import fuzz
import pandas as pd
import datetime

#creating date stamp
now = str(datetime.datetime.now())[:10]

#reading input file
df = pd.read_excel('Input_Data.xlsx')

# creating new fields
df['Match %'] = 0
df['Field3'] = df['Field1'].isnull()
for rows in df.index:
    if df['Field3'][rows] == True:
        df['Field1'][rows]=''

df['Field4'] = df['Field2'].isnull()
for rows in df.index:
    if df['Field4'][rows] == True:
        df['Field2'][rows]=''

#calculating average match %
for rows in df.index:
    
    df['Match %'][rows] = (fuzz.ratio(df['Field1'][rows], df['Field2'][rows]) + fuzz.partial_ratio(df['Field1'][rows], df['Field2'][rows]) + fuzz.token_sort_ratio(df['Field1'][rows], df['Field2'][rows])+fuzz.token_set_ratio(df['Field1'][rows], df['Field2'][rows]) )/4
    

#deleting extra columns
del df['Field3']
del df['Field4']

# Create a Pandas Excel writer using XlsxWriter as the engine.
name =  'Fuzzy_Match_Output_' + now + '.xlsx'
#df.to_excel(name)

writer = pd.ExcelWriter(name, engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Add a format. Light red fill with dark red text
format1 = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

lrow = str(df.index.max() + 2)
worksheet.conditional_format('E2:E'+lrow, {'type':     'cell',
                                        'criteria': '<',
                                        'value':    60,
                                        'format':   format1})
    
# Add a format. Green fill with dark green text
format2 = workbook.add_format({'bg_color': '#C6EFCE',
                               'font_color': '#006100'})
    
worksheet.conditional_format('E2:E'+lrow, {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    60,
                                        'format':   format2})
    
# Close the Pandas Excel writer and output the Excel file.
writer.save()
