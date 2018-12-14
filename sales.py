import numpy as np
import pandas as pd
df = pd.read_excel('sales.xls')
df = df[2:5384]
print (df)
df = df.fillna(method='backfill')
df = df.drop_duplicates(subset = 'Sales by customer as of current month', keep = 'last')

#print (df)
writer = pd.ExcelWriter('output.xlsx')
df.to_excel(writer,'Sheet1')
writer.save()