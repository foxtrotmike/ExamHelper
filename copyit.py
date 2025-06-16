# -*- coding: utf-8 -*-
"""
CS429 Assignment 2 Marks Integration Script
===========================================

Purpose:
--------
This script integrates final marks from the feedback sheet (`CS429-Assignment-2.xlsx`) into
the main CS429 gradebook (`CS429-15.xlsx`) by computing the final mark for each student,
favouring the 'Adjustment Mark' if available, otherwise using the 'Feedback Mark'.

Steps Performed:
----------------
1. Load the feedback file and calculate a 'Final Mark' per student.
2. Load the main CS429 gradebook with a two-row header and flatten it.
3. Merge final marks into the appropriate student rows using University ID.
4. Overwrite the 'Assignment 2' column in the gradebook with the computed final marks.
5. Save the result as `CS429-15_updated.xlsx`.

Next Step (Manual):
-------------------
After running this script, open `CS429-15_updated.xlsx` and **copy the updated
'Assignment 2' column** into the original gradebook file (`CS429-15.xlsx`) to
finalise the marks.

This manual copy step ensures that no other formatting or formulas in the original
gradebook are unintentionally altered.

Author:
-------
Fayyaz Minhas
"""

import pandas as pd
import os

# ----------------------------
# Step 1: Define base directory
# ----------------------------
# Use a relative or absolute path to the folder containing the Excel files
base_dir = r'C:\Users\fayya\OneDrive - University of Warwick\Desktop\CS429'

# ----------------------------
# Step 2: Load feedback marks
# ----------------------------
file_feedback = os.path.join(base_dir, 'CS429-Assignment-2.xlsx')
df_feedback = pd.read_excel(file_feedback)

# Compute the final mark: use 'Adjustment Mark' if present, else 'Feedback Mark'
df_feedback['Final Mark'] = df_feedback.apply(
    lambda row: row['Adjustment Mark'] if pd.notna(row['Adjustment Mark']) else row['Feedback Mark'],
    axis=1
)

# ----------------------------
# Step 3: Load main gradebook
# ----------------------------
file_gradebook = os.path.join(base_dir, 'CS429-15.xlsx')
df_gradebook = pd.read_excel(file_gradebook, header=[0, 1])  # Two-row header

# Flatten MultiIndex columns into single-line strings
df_gradebook.columns = [' '.join(col).strip() for col in df_gradebook.columns.values]

# ----------------------------
# Step 4: Extract relevant columns
# ----------------------------
id_column = 'University ID'
assignment_col = 'Assignment 2'

# Identify columns by partial match to avoid exact column name dependence
df_ids = df_gradebook[[col for col in df_gradebook.columns if id_column in col]]
df_scores = df_gradebook[[col for col in df_gradebook.columns if assignment_col in col]]

# Standardise column names for merging
df_ids.columns = [id_column]
df_scores.columns = [assignment_col]

# ----------------------------
# Step 5: Merge and update marks
# ----------------------------
# Match final marks from feedback sheet to gradebook using University ID
df_merged = df_ids.merge(
    df_feedback[['Student University Id', 'Final Mark']],
    left_on=id_column,
    right_on='Student University Id',
    how='left'
)

# Overwrite assignment scores with final marks
df_gradebook[assignment_col] = df_merged['Final Mark']

# ----------------------------
# Step 6: Save updated gradebook
# ----------------------------
output_file = os.path.join(base_dir, 'CS429-15_updated.xlsx')
df_gradebook.to_excel(output_file, index=False)
