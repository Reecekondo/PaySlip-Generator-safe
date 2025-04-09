# Step 1: Read Excel file with pandas
# Step 2: For each employee:
#         - Calculate net salary
#         - Generate a PDF payslip
#         - Send the payslip to their email

import pandas as pd

try:
    df = pd.read_excel("./Wolf-Fuels.xlsx")  # Make sure the file exists
    print("Pandas loaded successfully! Sample data:")
    print(df.head())  # Show first few rows
except Exception as e:
    print("Error loading file:", e)



import pandas as pd
import pandas as pd

# Load the Excel file
df = pd.read_excel("./Wolf-Fuels.xlsx")

# Remove currency symbols and convert to float
df["Basic Salary"] = df["Basic Salary"].replace({'\$': ''}, regex=True).astype(float)
df["Allowances"] = df["Allowances"].astype(float)
df["Deductions"] = df["Deductions"].astype(float)

# Calculate Net Salary
df["Net Salary"] = df["Basic Salary"] + df["Allowances"] - df["Deductions"]

# Display results
print(df.head())  # Show the first few rows with Net Salaryoutput 




from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Test generating a single payslip
c = canvas.Canvas("test_payslip.pdf", pagesize=letter)
c.drawString(100, 750, "Payslip Generator Test")
c.drawString(100, 730, "Employee Name: John Doe")
c.drawString(100, 710, "Net Salary: $1800")
c.save()

print("Test PDF generated successfully!")

