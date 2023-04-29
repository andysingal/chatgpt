import os
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Set the path to the Excel workbook and the name of the sheet
workbook_path = 'Financial_Data.xlsx'
sheet_name = 'Email_List'

# Set the path to the folder containing the attachments
attachment_path = os.path.join(os.path.dirname(workbook_path), 'Attachments')

# Connect to Gmail SMTP server
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = 'andysingal@gmail.com'
smtp_password = 'gfqhcfjlotfavsbo'
smtp_conn = smtplib.SMTP(smtp_server, smtp_port)
smtp_conn.ehlo()
smtp_conn.starttls()
smtp_conn.login(smtp_username, smtp_password)

# Open the Excel workbook and select the sheet
workbook = openpyxl.load_workbook(workbook_path)
sheet = workbook[sheet_name]

# Iterate over the rows in the sheet and compose the emails
for row in sheet.iter_rows(min_row=2, values_only=True):
    attachment_file = row[0]
    recipient_name = row[1]
    recipient_email = row[2]
    cc_email = row[3]

    # Construct the file path to the attachment
    attachment_path_full = os.path.join(attachment_path, attachment_file)

    # Compose the email
    msg = MIMEMultipart()
    msg['To'] = recipient_email
    msg['Cc'] = cc_email
    msg['Subject'] = 'Financial Data'

    body = f'Dear {recipient_name},\n\nPlease find attached the financial data you requested.\n\nBest regards,\nJohn'
    msg.attach(MIMEText(body))

    with open(attachment_path_full, 'rb') as f:
        part = MIMEApplication(f.read(), Name=attachment_file)
        part['Content-Disposition'] = f'attachment; filename="{attachment_file}"'
        msg.attach(part)

    # Send the email
    smtp_conn.send_message(msg)

# Close the SMTP connection
smtp_conn.quit()

# Close the workbook
workbook.close()