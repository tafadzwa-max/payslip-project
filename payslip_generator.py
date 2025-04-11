import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from pathlib import Path
import yagmail
from reportlab.lib import colors
from dotenv import load_dotenv
import logging
from reportlab.platypus import Table, TableStyle
from typing import Dict, Any
from fpdf import FPDF

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("payroll_processing.log"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

# Load environment variables
load_dotenv()

def setup_environment():
    """Ensure all required environment variables are set."""
    required_vars = ["SMTP_SERVER", "SMTP_PORT", "SENDER_EMAIL", "EMAIL_PASSWORD"]
    missing_vars = [var for var in required_vars if os.getenv(var) is None]
    if missing_vars:
        raise ValueError(f"Missing required environment variables: {missing_vars}")

def read_employee_data(file_path: str) -> pd.DataFrame:
    """Read employee data from Excel file."""
    try:
        df = pd.read_excel(file_path)
        required_columns = [
            "Employee ID",
            "Name",
            "Email",
            "Basic Salary",
            "Allowances",
            "Deductions",
        ]
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing required columns: {missing_columns}")

        for col in ["Basic Salary", "Allowances", "Deductions"]:
            if not pd.api.types.is_numeric_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], errors="coerce")
                null_values = df[df[col].isnull()].shape[0]
                if null_values > 0:
                    raise ValueError(f"{col} contains {null_values} invalid values")
        return df
    except Exception as e:
        logger.error(f"Error reading Excel file: {str(e)}", exc_info=True)
        raise

def calculate_net_salary(row: Dict[str, Any]) -> float:
    """Calculate net salary for an employee."""
    return row["Basic Salary"] + row["Allowances"] - row["Deductions"]

def generate_payslip_with_design(employee_data: Dict[str, Any], output_dir: Path) -> str:
    """Generate a payslip PDF for an employee with custom colors and layout."""
    try:
        # Create output directory if it doesn't exist
        output_dir.mkdir(parents=True, exist_ok=True)
        pdf_filename = f"{employee_data['Employee ID']}.pdf"
        full_path = output_dir / pdf_filename

        c = canvas.Canvas(str(full_path), pagesize=letter)
        width, height = letter

        # Header Section with Background Color
        c.setFillColor(colors.blue)
        c.rect(0, height - 70, width, 100, fill=True)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 20)
        c.drawString(200, height - 50, "TINPROTA Hardware")

        # Invoice Information Section
        c.setFont("Helvetica", 12)
        c.setFillColor(colors.black)
        c.drawString(50, height - 120, f"Employee ID : {employee_data['Employee ID']}")
        c.drawString(400, height - 120, "Date: April 10, 2025")

        # Billed To Section with Grey Background
        c.setFont("Helvetica-Bold", 12)
        c.drawString(60, height - 160, "BILLED TO:")
        c.setFont("Helvetica", 12)
        c.drawString(60, height - 180, employee_data["Name"])
        c.drawString(60, height - 200, f"Email: {employee_data['Email']}")

        # TABLE
        data = [
        ("DESCRIPTION", "AMOUNT (USD)"),
        ("Basic Salary", f"${employee_data['Basic Salary']:,.2f}"),
        ("Allowances", f"${employee_data['Allowances']:,.2f}"),
        ("Deductions", f"${employee_data['Deductions']:,.2f}"),
        ("Net Salary", f"${employee_data['Net Salary']:,.2f}")
        ]

        table = Table(data, colWidths=[200, 300])
        table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1, 0), colors.blue),
        ('TEXTCOLOR', (0,0), (-1, 0), colors.white),
        ('ALIGN', (1,1), (-1, -1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('GRID', (0,0), (-1, -1), 0.5,colors.grey),
         
    ]))
        table.wrapOn(c, width, height)
        table.drawOn(c, 60, height-350)

        # Footer with Background Color
        c.setFillColor(colors.whitesmoke)
        c.rect(0, 0, width, 50, fill=True)
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        c.drawString(50, 20, "THANK YOU FOR YOUR HARDWORK!")

        c.save()
        logger.info(f"Payslip with design generated successfully for {employee_data['Name']}")
        return str(full_path)
    except Exception as e:
        logger.error(f"Error generating payslip for {employee_data['Name']}: {str(e)}", exc_info=True)
        raise

def send_payslip_via_email(email_config: Dict[str, str], employee_data: Dict[str, Any], pdf_path: str):
    """Send payslip via email."""
    try:
        yag = yagmail.SMTP(
            email_config["SENDER_EMAIL"],
            email_config["EMAIL_PASSWORD"],
            host=email_config["SMTP_SERVER"],
            port=int(email_config["SMTP_PORT"]),
            smtp_starttls=True,
            smtp_ssl=False,
        )

        subject = "Your Monthly Payslip"
        body = f"Dear {employee_data['Name']},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nTINPROTA Hardware."
        yag.send(to=employee_data["Email"], subject=subject, contents=body, attachments=pdf_path)
        logger.info(f"Payslip sent successfully to {employee_data['Email']}")
    except Exception as e:
        logger.error(f"Failed to send payslip to {employee_data['Email']}: {str(e)}", exc_info=True)
        raise

def process_payroll(input_file: str = "employees.xlsx", output_dir: str = "payslips"):
    """Main function to process payroll."""
    try:
        setup_environment()
        output_dir_path = Path(output_dir)
        df = read_employee_data(input_file)
        email_config = {
            "SMTP_SERVER": os.getenv("SMTP_SERVER"),
            "SMTP_PORT": os.getenv("SMTP_PORT"),
            "SENDER_EMAIL": os.getenv("SENDER_EMAIL"),
            "EMAIL_PASSWORD": os.getenv("EMAIL_PASSWORD"),
        }
        success_count = 0
        failure_count = 0

        for _, row in df.iterrows():
            employee_data = row.to_dict()
            try:
                pdf_path = generate_payslip_with_design(employee_data, output_dir_path)
                send_payslip_via_email(email_config, employee_data, pdf_path)
                success_count += 1
            except Exception as e:
                logger.error(f"Failed to process payslip for {employee_data.get('Name', 'Unknown')}: {str(e)}", exc_info=True)
                failure_count += 1

        logger.info(f"Payroll processing completed - Success: {success_count}, Failures: {failure_count}")
    except Exception as e:
        logger.error(f"Overall payroll processing failed: {str(e)}", exc_info=True)
        raise

if __name__ == "__main__":
    process_payroll()