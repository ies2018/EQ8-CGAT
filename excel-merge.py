# -*- coding: utf-8 -*-
"""
Created on Tue Jul 17 10:03:35 2018
@author: Brian

Utility to merge Excel templates into one file (output_file)
1. finds all Excel file in the given directory
2. Forms them into a list (data_files)
3. Reads each file into dataframe
4. Drops the first row
5. Converts empty cells to 0
6. Saves the combined dataframe as a new Excel worksheet file
"""

import os
import glob
import pandas as pd
import numpy as np

#### 1. finds all Excel file in the given directory         
data_files = glob.glob(r'C:\TARS\AAAActive\EQ8_Gap_Analysis_2018\NewRegs\*.xlsx')         

### 2. Forms them into a list (data_files)
print (" ".join(map(str,data_files)))

print("\nTotal files found is ",len(data_files))

### 3. Reads each file into dataframe
all_data = pd.DataFrame()
for f in data_files:
    df = pd.read_excel(f)
    
    ### 4. Drops the first row
    df = df.drop([0])
    all_data = all_data.append(df,ignore_index=True)

### 5. Converts empty cells to 0    
all_data.fillna(0)

### 6. Saves the combined dataframe as a new Excel worksheet file
output_file = 'master_reg.xlsx'
os.remove(output_file)    
writer = pd.ExcelWriter(output_file)
all_data.to_excel(writer)
writer.save()