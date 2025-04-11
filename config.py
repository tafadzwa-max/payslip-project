import os 
from pathlib import Path

class Config:
    SMTP_HOST = os.getenv('SMTP_HOST', 'smtp.gmail.com')
    SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
    FROM_EMAIL = os.getenv('FROM_EMAIL')
    EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
    PDF_OUTPUT_DIR = Path('payslips')
    SMTP_SERVER = 'smtp.gmail.com'
    