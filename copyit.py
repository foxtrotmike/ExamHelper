# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 16:03:12 2024

@author: fayya
"""

import pandas as pd
bdir = r'C:\Users\fayya\OneDrive - University of Warwick\Desktop\CS429/'
# Read the first file
file1 = pd.read_excel(bdir+'CS429-Assignment-2.xlsx')

# Determine the final marks based on 'Adjustment Mark' or 'Feedback Mark'
file1['Final Mark'] = file1.apply(
    lambda row: row['Adjustment Mark'] if pd.notna(row['Adjustment Mark']) else row['Feedback Mark'], axis=1)

# Read the second file, setting the header as the first two rows
file2 = pd.read_excel(bdir+'CS429-15.xlsx', header=[0, 1])

# Flatten the MultiIndex columns to single level
file2.columns = [' '.join(col).strip() for col in file2.columns.values]

# Locate the ID column in file2 and the target column
id_column = 'University ID'
target_column_prefix = 'Assignment 2'

# Extract the relevant columns for easier handling
file2_ids = file2[[col for col in file2.columns if id_column in col]]
file2_target = file2[[col for col in file2.columns if target_column_prefix in col]]

# Flatten the MultiIndex columns for easier comparison
file2_ids.columns = [id_column]
file2_target.columns = [target_column_prefix]

# Merge the final marks from file1 into file2 based on IDs
merged_df = file2_ids.merge(file1[['Student University Id', 'Final Mark']], left_on=id_column, right_on='Student University Id', how='left')

# Add the 'Final Mark' to the target column in file2
file2[target_column_prefix] = merged_df['Final Mark']

# Save the updated file2
file2.to_excel(bdir+'CS429-15_updated.xlsx', index=False)
