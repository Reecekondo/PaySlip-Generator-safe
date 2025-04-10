**🧾 Payslip Generator Project**

This project is a complete Python-based tool that:

Reads employee data from an Excel file

Calculates net salary using basic salary, allowances, and deductions

Generates payslips as PDF files using ReportLab

Emails those payslips to each employee using Gmail SMTP

*🚀 How to Use*

1. Install Requirements

Make sure you have Python installed, then run:

pip install pandas openpyxl reportlab yagmail

2. Setup Your Excel File

Your Excel file should be named employees.xlsx and include the following columns:

Employee ID

Name

Email

Basic Salary

Allowances

Deductions

Place it in the same folder as your script.

3. Generate a Gmail App Password

Gmail blocks less secure apps. Use an App Password instead:

Visit: https://myaccount.google.com/apppasswords

Choose Mail > Custom name (e.g. PayslipSender) > Generate

Copy the 16-digit password for use in your script

4. Run the Script

Run your Python file in the terminal:

python payslip_generator.py

Enter your Gmail and App Password when prompted

Payslips will be saved in the payslips/ folder

Emails will be sent to employees with the PDF attached

*🛠 Technologies Used*

Python 3.x

pandas – for reading and processing Excel data

reportlab – for generating PDF payslips

smtplib / yagmail – for sending emails

openpyxl – for reading .xlsx files with pandas

*📁 Project Structure*

Payslip Generator/
├── employees.xlsx
├── payslip_generator.py
├── payslips/         # Contains all generated PDF files
└── README.md

✍️ Author

Reece KurangwaStudent @ Uncommon.org BootcampBuilt with guidance, grit, and growth mindset 🚀

