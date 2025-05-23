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







# Step 3: Generate PDF Payslips

import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Image, Paragraph
from reportlab.lib.styles import getSampleStyleSheet

# Define the output directory
output_dir = "payslips"
os.makedirs(output_dir, exist_ok=True)

# Load Excel Data
excel_file = "Wolf-Fuels.xlsx"
df = pd.read_excel(excel_file)

# Fix column name inconsistencies
df.columns = df.columns.str.strip()
df.rename(columns={"Employee-ID": "Employee_ID", "Basic-Salary": "Basic_Salary", "Deducations": "Deductions"}, inplace=True)

# Function to generate a professional payslip
def generate_payslip(employee):
    employee_id = str(employee["Employee_ID"])
    name = employee["Names"]
    basic_salary = employee["Basic_Salary"]
    allowances = employee["Allowances"]
    deductions = employee["Deductions"]
    net_salary = basic_salary + allowances - deductions  

    # Create PDF document
    pdf_filename = f"{output_dir}/payslip_{employee_id}.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # **Company Header with Logo (Left-Aligned)**
    logo_path = "Wolf_logo_v3.png"
    company_logo = Image(logo_path, width=100, height=50)
    company_info = [
        [company_logo, Paragraph("<b>Wolf Fuels</b><br/>123 Business Street, Harare, Zimbabwe<br/>Phone: +263 77 123 4567 | Email: contact@wolffuels.com", styles["Normal"])]
    ]
    company_table = Table(company_info, colWidths=[120, 280])
    company_table.setStyle(TableStyle([("ALIGN", (0, 0), (0, 0), "LEFT"), ("ALIGN", (1, 0), (1, 0), "LEFT")]))
    
    elements.append(company_table)
    elements.append(Spacer(1, 15))

    # **Payslip Title with Payroll Info**
    payslip_header = [["Payslip"], [f"Pay Date: April 10, 2025"], [f"Payroll No: {employee_id}"]]
    header_table = Table(payslip_header, colWidths=[400])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ]))
    
    elements.append(header_table)
    elements.append(Spacer(1, 15))

    # **Employee Information Section**
    employee_info = [
        ["Employee Name:", name],
        ["Employee ID:", employee_id],
        ["Email:", employee["Email"]]
    ]
    employee_table = Table(employee_info, colWidths=[200, 200])
    employee_table.setStyle(TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    
    elements.append(employee_table)
    elements.append(Spacer(1, 15))

    # **Earnings Breakdown**
    earnings_details = [
        ["Earnings", "Amount"],
        ["Basic Salary", f"${basic_salary:.2f}"],
        ["Allowances", f"${allowances:.2f}"]
    ]
    earnings_table = Table(earnings_details, colWidths=[200, 200])
    earnings_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(earnings_table)
    elements.append(Spacer(1, 15))

    # **Deductions Breakdown**
    deductions_details = [
        ["Deductions", "Amount"],
        ["Deductions", f"${deductions:.2f}"]
    ]
    deductions_table = Table(deductions_details, colWidths=[200, 200])
    deductions_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkred),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(deductions_table)
    elements.append(Spacer(1, 15))

    # **Net Salary (Highlighted)**
    net_salary_table = Table([["Net Salary", f"${net_salary:.2f}"]], colWidths=[200, 200])
    net_salary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.green),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(net_salary_table)
    elements.append(Spacer(1, 20))

    # **Footer with Contact Details**
    footer_details = Paragraph(
        "<b>For any inquiries, contact:</b><br/>HR Department | Wolf Fuels<br/>Phone: +263 77 123 4567 | Email: hr@wolffuels.com",
        styles["Normal"]
    )
    elements.append(footer_details)

    # Generate PDF for the current employee
    doc.build(elements)
    print(f"Payslip generated: {pdf_filename}")

# Generate payslips for all employees
for _, employee in df.iterrows():
    print(f"Processing payslip for: {employee['Names']}")
    generate_payslip(employee)

print("All payslips generated successfully!")












# Sending emails with payslips to my employees
# Step 4: Send Payslips via Email


import smtplib
import os
import pandas as pd
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Setup logging
logging.basicConfig(filename="logs/email_errors.log", level=logging.ERROR)

# Email Configuration
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.example.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SENDER_EMAIL = os.getenv("SENDER_EMAIL", "your_email@example.com")
PASSWORD = os.getenv("EMAIL_PASSWORD")

# Load employee data
file_path = "data/Wolf-Fuels.xlsx"
try:
    df = pd.read_excel(file_path)
    df.columns = df.columns.str.strip()
except Exception as e:
    logging.error(f"Error loading employee data: {str(e)}")
    print(f"ERROR: Could not load employee data! {str(e)}")
    exit()
except Exception as e:
    logging.error(f"Error loading employee data from {file_path}: {str(e)}")
    
# Function to send payslip email
def send_payslip(employee):
    try:
        employee_email = str(employee["Email"]).strip()
        employee_id = str(employee["Employee_ID"]).strip()
        payslip_file = f"payslips/payslip_{employee_id}.pdf"

        if not os.path.exists(payslip_file):
            error_msg = f"ERROR: Payslip file missing for {employee['Names']}, skipping..."
            logging.error(error_msg)
            print(error_msg)
            return

        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = employee_email
        msg["Subject"] = "Your Payslip for This Month"

        body = f"Dear {employee['Names']},\n\nPlease find your payslip attached.\n\nBest regards,\nFinance Team"
        msg.attach(MIMEText(body, "plain"))

        # Attach the payslip PDF
        with open(payslip_file, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="payslip_{employee_id}.pdf"')
            msg.attach(part)

        # Send Email
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(SENDER_EMAIL, PASSWORD)
        server.sendmail(SENDER_EMAIL, employee_email, msg.as_string())
        server.quit()
        print(f"Payslip sent to {employee_email}")
    except smtplib.SMTPException as e:
        logging.error(f"SMTP Error for {employee_email}: {str(e)}")
        print(f"ERROR: Could not send payslip to {employee_email}.")
    except Exception as e:
        logging.error(f"Unexpected error for {employee_email}: {str(e)}")
        print(f"ERROR: Unknown issue sending payslip to {employee_email}.")

# Send payslip emails
df.apply(send_payslip, axis=1)

print("Payslips emailed successfully!")




