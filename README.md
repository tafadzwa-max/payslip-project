# Payslip Generator

A Python program that generates and emails payslips to employees based on Excel data.

## Requirements

- Python 3.8+
- Required packages:
  * pandas
  * yagmail
  * reportlab
  * python-dotenv

## Setup

1. Install required packages:
   ```bash
pip install pandas yagmail reportlab python-dotenv
```

2. Create a `.env` file with your email configuration:
   ```bash
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
SENDER_EMAIL=your-email@gmail.com
EMAIL_PASSWORD=your-app-password
```

3. Create an `employees.xlsx` file with the following columns:
   - Employee ID
   - Name
   - Email
   - Basic Salary
   - Allowances
   - Deductions

## Usage

1. Run the program:
   ```bash
python payslip_generator.py
```

2. The program will:
   - Read employee data from `employees.xlsx`
   - Generate PDF payslips in the `payslips/` directory
   - Email each payslip to the corresponding employee
   - Log all operations to `payslip_generator.log`

## Troubleshooting

1. Email Issues:
   - Ensure your email account allows less secure apps
   - For Gmail, generate an app password instead of using your regular password
   - Check the log file for specific error messages

2. PDF Issues:
   - Verify the `payslips/` directory is created
   - Check file permissions
   - Review the log file for generation errors

3. Excel Issues:
   - Ensure all required columns exist
   - Verify salary columns contain numeric values
   - Check for any special characters in file paths