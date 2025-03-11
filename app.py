import os
import re
import pdfplumber
import pandas as pd
import shutil
import smtplib
import time
import random
import logging
from PyPDF2 import PdfReader, PdfWriter
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import streamlit as st

# Configure logging
logging.basicConfig(filename="email_errors.log", level=logging.ERROR, format="%(asctime)s - %(message)s")

def auth_id(smtp_username, smtp_password, smtp_server, smtp_port):
    """Authenticate SMTP credentials before sending emails."""
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        try:
            server.login(smtp_username, smtp_password)
            print("✅ SMTP Authentication Successful!")
            return True
        except Exception as e:
            print(f"❌ SMTP Authentication Failed! Error: {e}")
            return False

def send_email(subject, emp_name, emp_salary, sender_email, receiver_email, attachment_filename, smtp_server, smtp_port, smtp_username, smtp_password, retries=3, delay=5):
    """Send an email with an attached payslip PDF, with retry logic."""
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    email_body = f"""
    Dear {emp_name},

    Please find attached your payslip for this month.
    
    If you have any questions, feel free to reach out.

    Best Regards,
    """
    msg.attach(MIMEText(email_body, 'plain'))

    if attachment_filename and os.path.exists(attachment_filename):
        with open(attachment_filename, 'rb') as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(attachment_filename)}")
        msg.attach(part)
    else:
        print(f"⚠️ Warning: Attachment file not found - {attachment_filename}")

    for attempt in range(1, retries + 1):
        try:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(smtp_username, smtp_password)
                server.sendmail(sender_email, receiver_email, msg.as_string())

            print(f"✅ Email sent to {receiver_email}")
            return True

        except Exception as e:
            print(f"❌ Attempt {attempt}/{retries}: Failed to send email to {receiver_email}. Error: {e}")
            logging.error(f"Failed to send email to {receiver_email} on attempt {attempt}. Error: {e}")
            time.sleep(delay)  # Wait before retrying

    print(f"❌ Final failure: Could not send email to {receiver_email} after {retries} attempts.")
    return False

def send_mail(email_id, email_pass, payslip_pdf, emp_sheet, subject, month, year):
    try:
        print("📂 Loading Employee Data...")
        data = pd.read_excel(emp_sheet)
        data.columns = data.columns.str.strip()
        print("✅ Employee Data Loaded Successfully!")

        pdf_file_path = payslip_pdf
        out_folder = f'{month}-{year}'
        shutil.rmtree(out_folder, ignore_errors=True)
        os.makedirs(out_folder, exist_ok=True)

        pdf_reader = PdfReader(pdf_file_path)
        total_pages = len(pdf_reader.pages)

        data_lst = []
        print("🔍 Extracting Employee Details from Payslips...")

        for page_num, page in enumerate(pdf_reader.pages):
            output_file_path = f"{out_folder}/{page_num}.pdf"
            pdf_writer = PdfWriter()
            pdf_writer.add_page(page)
            with open(output_file_path, 'wb') as output_file:
                pdf_writer.write(output_file)

        for file_name in os.listdir(out_folder):
            file_path = os.path.join(out_folder, file_name)
            with pdfplumber.open(file_path) as pdf:
                text = "\n".join([p.extract_text() or "" for p in pdf.pages])

            emp_code_match = re.search(r'Employee\s*Code[:\s]*([\d]+)', text, re.IGNORECASE)
            emp_code = emp_code_match.group(1) if emp_code_match else ''
            emp_salary_match = re.search(r'NET\s*AMOUNT\s*PAYABLE[:\s]*([\d,.]+)', text, re.IGNORECASE)
            emp_salary = emp_salary_match.group(1) if emp_salary_match else ''

            if not emp_code or not emp_salary:
                continue

            emp_data = data[data['Emp Code.'].astype(str) == emp_code]
            if emp_data.empty:
                continue

            emp_name = emp_data['Employee Name'].values[0]
            emp_email = emp_data['Email Address'].values[0]
            if pd.isna(emp_email) or emp_email.strip() == "":
                data_lst.append([emp_code, emp_name, emp_email, '', 'No Email Found'])
                continue

            new_file_name = f"{out_folder}/{emp_code}-{emp_name} Payslip.pdf"
            os.rename(file_path, new_file_name)
            data_lst.append([emp_code, emp_name, emp_email, new_file_name, 'Ready to Send'])

        smtp_server = 'smtp.gmail.com'
        smtp_port = 465
        smtp_username = email_id
        smtp_password = email_pass

        auth = auth_id(smtp_username, smtp_password, smtp_server, smtp_port)
        if not auth:
            return 0, total_pages, data_lst

        emails_sent = 0
        batch_count = 0  # Track number of emails sent in a batch

        for i, (emp_code, emp_name, emp_email, attachment_filename, status) in enumerate(data_lst):
            if status == 'Ready to Send':
                if send_email(subject, emp_name, emp_salary, email_id, emp_email, attachment_filename, smtp_server, smtp_port, smtp_username, smtp_password):
                    emails_sent += 1
                    data_lst[i][-1] = 'Sent'
                    batch_count += 1
                else:
                    data_lst[i][-1] = 'Not Sent'

                time.sleep(random.randint(2, 5))

                # Add a 2-minute delay after every 10 emails sent
                if batch_count == 50:
                    print("⏳ Waiting 2 minutes before sending more emails...")
                    time.sleep(120)
                    batch_count = 0  # Reset the batch count

        shutil.rmtree(out_folder, ignore_errors=True)
        return emails_sent, total_pages, data_lst

    except Exception as e:
        print(f"❌ Error Occurred: {e}")
        return 0, 0, []


def main():
    st.title("PaySlip Email Sender")
    email_id = st.text_input("Enter your email address:")
    email_pass = st.text_input("Enter your email password:", type="password")
    payslip_pdf = st.file_uploader("Upload the payslip PDF file:", type=["pdf"])
    emp_sheet = st.file_uploader("Upload the employee sheet (Excel file):", type=["xlsx", "xls"])
    subject = st.text_input("Enter email subject:")
    month = st.text_input("Enter month:")
    year = st.text_input("Enter year:")

    if st.button("Send Email"):
        count, total_pages, data_lst = send_mail(email_id, email_pass, payslip_pdf, emp_sheet, subject, month, year)
        st.success(f"Out of {total_pages} payslips, {count} emails were sent successfully!")
        st.dataframe(pd.DataFrame(data_lst, columns=['Employee ID', 'Employee Name', 'Email Address', 'Attachment', 'Status']))

if __name__ == "__main__":
    main()
