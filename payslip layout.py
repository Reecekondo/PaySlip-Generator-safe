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

import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
import os

# Load employee data
df = pd.read_excel("Wolf-Fuels.xlsx")

# Create output folder
os.makedirs("payslips", exist_ok=True)

for index, row in df.iterrows():
    employee_id = row["Employee ID"]
    name = row["Name"]
    email = row["Email"]
    basic_salary = row["Basic Salary"]
    allowances = row["Allowances"]
    deductions = row["Deductions"]
    net_salary = basic_salary + allowances - deductions

    # PDF setup
    filename = f"payslips/{employee_id}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=letter)
    styles = getSampleStyleSheet()
    
    elements = []

    # Title
    elements.append(Paragraph("<b>Wolf-Fuels Payslip</b>", styles["Title"]))
    elements.append(Spacer(1, 12))

    # Employee Details
    elements.append(Paragraph(f"Employee ID {employee_id}", styles["Normal"]))
    elements.append(Paragraph(f"Name: {name}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    # Salary Breakdown
    elements.append(Paragraph(f"Basic Salary: ${basic_salary}", styles["Normal"]))
    elements.append(Paragraph(f"Allowances: ${allowances}", styles["Normal"]))
    elements.append(Paragraph(f"Deductions: ${deductions}", styles["Normal"]))
    elements.append(Paragraph(f"<b>Net Salary: ${net_salary}</b>", styles["Normal"]))

    # Generate PDF
    doc.build(elements)

    print(f"Payslip generated for {name} ({employee_id}) at {filename}")