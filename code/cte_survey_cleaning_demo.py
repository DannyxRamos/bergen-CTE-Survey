"""
cte_survey_cleaning_demo.py

This script demonstrates how to automate cleaning of raw SurveyMonkey survey exports,
which often contain two header rows and inconsistent formatting, especially for checkbox/matrix-style questions.

Key features:
- Automatically merges and formats header rows
- Removes irrelevant columns
- Prepares clean Excel file for analysis or visualization
- Combines grouped checkbox responses into unified columns

No actual student or institutional data is included in this demo version.
"""

import pandas as pd

# ---------------------------
# STEP 1: Load raw data with messy two-row header
# ---------------------------
df = pd.read_excel("data.xlsx", header=None)

# Drop first 9 metadata columns (e.g., timestamps, IDs)
df = df.iloc[:, 9:]

# Extract header rows
row0 = df.iloc[0].copy()
row1 = df.iloc[1].copy()

# ---------------------------
# STEP 2: Merge header rows into a single clean row
# ---------------------------
for col in df.columns:
    val1 = str(row1[col]).strip() if pd.notna(row1[col]) else ""
    if val1 in ["Response", "Open-Ended Response"]:
        continue
    elif val1 == "Other (please specify)":
        row0[col] = val1

filled_row0 = row0.copy()
cols = list(df.columns)

for i, col in enumerate(cols):
    val0 = str(row0[col]).strip() if pd.notna(row0[col]) else ""
    val1 = str(row1[col]).strip() if pd.notna(row1[col]) else ""
    if val0 == "" and val1 != "":
        filled_row0[col] = val1
        if i > 0:
            prev_col = cols[i - 1]
            val1_prev = str(row1[prev_col]).strip() if pd.notna(row1[prev_col]) else ""
            if val1_prev != "":
                filled_row0[prev_col] = val1_prev

# Apply new headers and drop original two rows
df.columns = filled_row0
df = df.drop(index=[0, 1]).reset_index(drop=True)

# Save cleaned survey data
df.to_excel("data_modified.xlsx", index=False)
print("Cleaned survey data saved to 'data_modified.xlsx'")

# ---------------------------
# STEP 3: Combine checkbox question blocks into unified columns (optional)
# ---------------------------

# Load cleaned file
df = pd.read_excel("data_modified2.xlsx")  # This assumes manual review/editing in between

# Helper function to stack grouped checkbox columns into single columns
def combine_columns_by_index(start_idx, end_idx, col_name):
    block = df.iloc[:, start_idx:end_idx + 1]
    combined = block.stack().reset_index(drop=True).to_frame(name=col_name)
    return combined

# Combine several grouped sections (index ranges reflect a SurveyMonkey export format)
q1 = combine_columns_by_index(25, 39, "Q1")
q2 = combine_columns_by_index(49, 67, "Q2")
q3 = combine_columns_by_index(68, 86, "Q3")
q4 = combine_columns_by_index(96, 107, "Q4")

# Save combined outputs
with pd.ExcelWriter("data_combined.xlsx", engine="openpyxl") as writer:
    q1.to_excel(writer, index=False, sheet_name="Q1")
    q2.to_excel(writer, index=False, sheet_name="Q2")
    q3.to_excel(writer, index=False, sheet_name="Q3")
    q4.to_excel(writer, index=False, sheet_name="Q4")

print("Combined grouped responses saved to 'data_combined.xlsx'")
