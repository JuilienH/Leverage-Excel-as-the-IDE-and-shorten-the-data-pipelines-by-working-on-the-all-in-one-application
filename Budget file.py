
from pickle import FALSE
import pandas as pd
import numpy as np
from datetime import datetime as dt



## Scenario:  BUDGET daily volumes
df_example = pd.read_excel('Essbase output file.xlsx', 'SME', skiprows = 32, nrows=366,  usecols= 'BW:CO',header=None)

#Period column padded in
df_example['Period'] = pd.date_range(start='1/1/2025', end='12/31/2025', freq='D')


#Common fields
df_example['New Column']='2025_this_year'

df_example.to_csv("daily_view_Budget_CSV.csv",index=False)

# no of csv files with row size in the output csv files
k = 3
size = 10000
 
for i in range(k):
    df = df_example[size*i:size*(i+1)]
    df.to_csv(f'daily_view_Budget_CSV_{i+1}.csv', index=False)
 

