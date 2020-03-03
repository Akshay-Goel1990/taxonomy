# -*- coding: utf-8 -*-
"""
Created on Mon Feb 24 10:26:45 2020

@author: aksgoel
"""

import re
import pandas as pd
import csv
import numpy as np
from scipy import stats
import math



data = pd.read_csv("C:/Users/aksgoel/Desktop/CM_Standard_file.csv")
Campaign = data['Campaign']

df = pd.DataFrame(columns=['Field','Line Item','Total Attribute','Maximum','Mean','Median','Mode','Count','Deviation','Taxonomy Sustained'])
df1 = pd.DataFrame(columns=['Line Item','Difference','Count_Rows'])
b=0
for i in Campaign:    
    Field = 'Campaign'
    Line_Item = i
    try:
        attribute = re.findall(r'[_]',i)
    except Exception as e:
        attribute = 0
    try:
        attribute_count = len(attribute)
    except Exception as e:
        attribute_count = 0
    
    df.loc[b] = [Field,Line_Item,attribute_count,0,0,0,0,0,0,0]
    b=b+1
    print (df[['Line Item', 'Total Attribute']])

df['Maximum'] = max(df['Total Attribute'])
df['Mean'] = np.mean(df['Total Attribute'])
df['Median'] = np.median(df['Total Attribute'])
Mode_Calculate = stats.mode(df['Total Attribute'])
Mode_value = Mode_Calculate[0]
df['Mode'] = Mode_value[0]
total_rows = len(df)
df['Count'] = total_rows

k = 0
for i1, j in df.iterrows():
    Placement = j['Line Item']
    Attribute = j['Total Attribute']
    Median1 = j['Median']
    Total_rows = j['Count']
    difference = abs(Attribute - Median1)
    df1.loc[k] = [Placement, difference,Total_rows]
    k = k+1
df1
df['Deviation'] = (df1['Difference']*df1['Count_Rows'])/sum(df1['Difference'])
df['Deviation'] = df['Deviation'].apply(lambda x:round(x,2))
df['Deviation']
df['Taxonomy Sustained'] = np.where(df['Deviation'] > 3, "Outlier", "With In Range")

df = df.drop(['Mean','Median','Mode','Count'], axis=1)
from StyleFrame import StyleFrame, Styler
from django.utils.formats import number_format

sf = StyleFrame(df)

sf.apply_column_style(cols_to_style='Taxonomy Sustained', styler_obj=Styler(font_size=12), width = 20)
sf.apply_style_by_indexes(sf[sf['Taxonomy Sustained'] == 'With In Range'], cols_to_style='Taxonomy Sustained',
                          styler_obj=Styler(bg_color='green'), height = 15)
sf.apply_style_by_indexes(sf[sf['Taxonomy Sustained'] == 'Outlier'], cols_to_style='Taxonomy Sustained',
                          styler_obj=Styler(bg_color='red'), height = 15)
sf.apply_column_style(cols_to_style='Line Item', styler_obj=Styler(font_size=12), width = 75)

sf.apply_column_style(cols_to_style='Field', styler_obj=Styler(font_size=12), width = 20)
sf.apply_column_style(cols_to_style='Deviation', styler_obj=Styler(font_size=12), width = 20)
sf.apply_column_style(cols_to_style='Total Attribute', styler_obj=Styler(font_size=12), width = 20)
sf.apply_column_style(cols_to_style='Maximum', styler_obj=Styler(font_size=12), width = 20)

sf.to_excel(r'C:/Users/aksgoel/Desktop/Taxonomy Overview.xlsx',index=False).save()
print ("Completed")
