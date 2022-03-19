#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import tabula
import re
import glob

table = tabula.read_pdf(glob.glob("*.pdf")[0], output_format= 'dataframe', pages='3-10', multiple_tables=False, pandas_options={'names': ["Description", 0, 1, 2, 3, 4, 5, 6, 'Unit Price', 7, 8, 9, 10, 11, 12]})

table = pd.DataFrame(table[0])

table = table.dropna(subset=['Description'])

table = table[table['Description'].str.fullmatch("\d{3}-\d{3}-\d{4}\s.+")]

table = table[['Description', 'Unit Price']].reset_index(drop=True)


table[['number', 'name']] = table['Description'].str.split(' ', 1, expand=True)


output = pd.read_excel(glob.glob("*.xlsx")[0], skiprows=4)


output[['number', 'name']] = output['Description'].str.split(' ', 1, expand=True)


merge = pd.merge(left=table,
        right=output,
        on=['number'],
        how='right',
        suffixes =['_new', '_old'])

merge['Unit Price_new'] = merge['Unit Price_new'].fillna(0)


merge = merge[['Description_old', 'Unit Price_new', 'Company ID', 'Account', 'Sub Account', 'Quantity']]


merge = merge.rename(columns = {'Description_old': 'Description', 'Unit Price_new': 'Unit Price'})


merge = merge[['Company ID', 'Description', 'Account', 'Sub Account', 'Unit Price', 'Quantity']]


merge.to_csv('New Invoice.csv')
