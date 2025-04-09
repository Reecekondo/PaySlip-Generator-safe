# Step 1: Read Excel file with pandas
# Step 2: For each employee:
#         - Calculate net salary
#         - Generate a PDF payslip
#         - Send the payslip to their email

import pandas as pd

try:
    df = pd.read_excel("PYTHON PROJECTS/Wolf-Fuels.xlsx")  # Make sure the file exists
    print("Pandas loaded successfully! Sample data:")
    print(df.head())  # Show first few rows
except Exception as e:
    print("Error loading file:", e)
