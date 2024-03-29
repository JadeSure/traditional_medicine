from docx import Document
from pathlib import Path
import pandas as pd

from dotenv import load_dotenv
import os

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

# Load environment variables from .env file
load_dotenv()

BASE_PATH = Path(__file__).resolve().parent


def replace_string_in_docx(docx_filename, old_string, new_string, output_path):
    doc = Document(docx_filename)
    file_name = str(docx_filename).split("/")[-1]
    for p in doc.paragraphs:
        for run in p.runs:
            if old_string in run.text:
                run.text = run.text.replace(old_string, new_string)
    doc.save(f'{output_path}_{file_name}')


def read_names(file_path):
    df = pd.read_excel(str(file_path))
    output = {}
    for row in df.itertuples():
        output[row.name] = row.email
    return output


def send_email(mail_user, mail_pass, mail_receiver):
    # QQ Mail SMTP server address and port
    smtp_server = 'smtp.qq.com'
    smtp_port = 465  # SSL

    # Your QQ Mail account and password
    mail_user = 'your-account@qq.com'
    # If you have enabled SMTP service on QQ Mail, use the authorization code.
    mail_pass = 'your-password'

    # Email content
    subject = 'Hello from Python'
    content = 'This is a test email sent from Python.'

    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = mail_user
    msg['To'] = 'receiver@example.com'
    msg['Subject'] = subject

    # Attach the email content
    msg.attach(MIMEText(content, 'plain'))

    # Specify the file path of your attachment
    file_path = '/path/to/your/file'

    # Create a MIMEBase object for the attachment
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(open(file_path, 'rb').read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment; filename="%s"' % os.path.basename(file_path))

    # Attach the file to the email
    msg.attach(part)

    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(mail_user, mail_pass)
        server.sendmail(mail_user, ['receiver@example.com'], msg.as_string())
        server.quit()
        print('Email sent successfully.')
    except smtplib.SMTPException as e:
        print('Error:', e)


if __name__ == "__main__":
    # Get environment variables
    mail_user = os.getenv('mail_user')
    mail_pass = os.getenv('mail_pass')
    mail_receiver = os.getenv('mail_receiver')

    source_cert = BASE_PATH / 'source' / '2024CPD证书.docx'
    source_people = BASE_PATH / 'source' / 'names.xlsx'

    print(BASE_PATH)
    people = read_names(source_people)
    for name in people.keys():
        print(type(name))
        replace_string_in_docx(source_cert, 'LEI    yao', name, f'{BASE_PATH}/res/{name}')
