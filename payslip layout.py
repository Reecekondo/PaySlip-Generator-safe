# Step 1: Read Excel file with pandas
# Step 2: For each employee:
#         - Calculate net salary
#         - Generate a PDF payslip
#         - Send the payslip to their email


import pandas as pd
import pandas as pd

# Load the Excel file
df = pd.read_excel("./Wolf-Fuels.xlsx")

# Calculate Net Salary
df["Net Salary"] = df["Basic-Salary"] + df["Allowances"] - df["Deductions"]

# Display results
print(df.head())  # Show the first few rows with Net Salaryoutput 



# Step 3: Generate PDF payslips for each employee
import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# Define the output directory properly
output_dir = "payslips"
os.makedirs(output_dir, exist_ok=True)

# Load Excel file
file_path = "Wolf-Fuels.xlsx"  # Change this to your actual file path
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()  # Clean column names

# Function to generate a professional payslip
def generate_payslip(employee):
    employee_id = str(employee["Employee-ID"])
    name = employee["Names"]
    basic_salary = employee["Basic-Salary"]
    allowances = employee["Allowances"]
    deductions = employee["Deductions"]
    net_salary = basic_salary + allowances - deductions  

    filename = f"{output_dir}/{employee_id}.pdf"
    pdf = canvas.Canvas(filename, pagesize=A4)

    # Payslip Design
    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(200, 780, "PAYSLIP")

    pdf.setFont("Helvetica", 12)
    pdf.drawString(50, 740, f"Employee ID: {employee_id}")
    pdf.drawString(50, 720, f"Name: {name}")
    pdf.drawString(50, 690, f"Basic Salary: ${basic_salary:.2f}")
    pdf.drawString(50, 670, f"Allowances: ${allowances:.2f}")
    pdf.drawString(50, 650, f"Deductions: ${deductions:.2f}")
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(50, 620, f"Net Salary: ${net_salary:.2f}")
    pdf.line(50, 610, 400, 610)

    # Save the PDF
    pdf.showPage()
    pdf.save()
    print(f"Payslip generated: {filename}")

# Generate payslips and confirm execution
for _, employee in df.iterrows():
    print(f"Processing payslip for: {employee['Names']}")
    generate_payslip(employee)




# Sending emails with payslips to my employees

import smtplib
import os
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Email Configuration
SMTP_SERVER = "smtp.gmail.com"  # Use your email provider's SMTP server
SMTP_PORT = 587
SENDER_EMAIL = "reecekhalid96@gmail.com"
PASSWORD = "bnamsdoikuzsizgt"  # Use an app password, NOT your regular password!

# Load employee data
file_path = "Wolf-Fuels.xlsx"
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip()  # Clean column names

# Function to send payslip email
def send_payslip(employee):
    employee_email = employee["Email"]
    employee_id = str(employee["Employee-ID"])
    payslip_file = f"payslips/{employee_id}.pdf"

    if not os.path.exists(payslip_file):
        print(f"Payslip not found for {employee['Names']}, skipping...")
        return
    
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = employee_email
    msg["Subject"] = "Your Payslip for This Month"

    # Email body
    msg.attach(MIMEBase("text", "plain"))
    msg.get_payload()[0].set_payload("Dear Employee,\n\nPlease find your payslip for this month attached.\n\nBest regards,\nFinance Team")

    # Attach the payslip PDF
    with open(payslip_file, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{employee_id}.pdf"')
        msg.attach(part)

    # Send Email
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, PASSWORD)
        server.sendmail(SENDER_EMAIL, employee_email, msg.as_string())
        server.quit()
        print(f"Payslip sent to {employee_email}")
    except Exception as e:
        print(f"Error sending to {employee_email}: {e}")

# Send payslip emails
for _, employee in df.iterrows():
    send_payslip(employee)




















# import pandas as pd
# import smtplib
# from email.message import EmailMessage

# # Load and clean Excel data
# df = pd.read_excel("Wolf-Fuels.xlsx")
# df.columns = df.columns.str.strip()  # Remove trailing spaces from column names

# # Login credentials
# sender_email = "Kurangwareece@gmail.com" 
# password = "bnamsdoikuzsizgt"  # Use an app password if using Gmail

# # Set up SMTP server
# server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
# server.login(sender_email, password)

# # Loop through each employee and send email
# for index, row in df.iterrows():
#     name = row['Names & Surname']
#     receiver_email = row['Email']
#     net_salary = row['Net Salary']

#     msg = EmailMessage()
#     msg['Subject'] = 'Your Payslip'
#     msg['From'] = sender_email
#     msg['To'] = receiver_email
#     msg.set_content(f"""
#     Dear {name},

#     Please find your payslip below.
    
#     Net Salary: ${net_salary:.2f}

#     Regards,
#     Payroll Team
#     """)

#     # Optional: attach a PDF (e.g., generated earlier)
#     # with open(f"payslips/{row['Employee ID']}.pdf", 'rb') as f:
#     #     file_data = f.read()
#     #     file_name = f"{row['Employee ID']}_Payslip.pdf"
#     #     msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

#     # Send the email
#     server.send_message(msg)
#     print(f"Email sent to {name} at {receiver_email}")

# # Close the SMTP server
# server.quit()

