import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# Email details
email_subject = "HW3 Grading Sheet"
email_body = """
Hello,

Please find the attached grading sheet for HW3.
Please let me know if you have any questions, via email or you can join any of the TAs office hours.

Thank You,

Regards,
Soumya

"""
sender_email = "Your_email_address_here"
sender_password = "Your_email_password_here"

# Directory containing the Excel files
directory = 'Path_to_folder_here'

# SMTP server details (example for Gmail)
smtp_server = "smtp.gmail.com"
smtp_port = 587

def send_email(recipient_email, file_path):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = email_subject

    # Attach the body with the msg instance
    msg.attach(MIMEText(email_body, 'plain'))

    # Open the file to be sent
    attachment = open(file_path, "rb")

    # Instance of MIMEBase and named as p
    part = MIMEBase('application', 'octet-stream')

    # To change the payload into encoded form
    part.set_payload((attachment).read())

    # Encode into base64
    encoders.encode_base64(part)

    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file_path)}")

    # Attach the instance 'part' to instance 'msg'
    msg.attach(part)

    # Create SMTP session for sending the mail
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    text = msg.as_string()
    server.sendmail(sender_email, recipient_email, text)
    server.quit()

# Loop through all Excel files in the directory
for file_name in os.listdir(directory):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(directory, file_name)

        # Read the Excel file to get the recipient email
        df = pd.read_excel(file_path)
        # print(df)
        email_cell_value = df.iloc[0, 4] 
        # print(email_cell_value)

        recipient_email = email_cell_value.split(':')[1].strip()
        print("Email :" + recipient_email)

        # Send the email with the attachment
        send_email(recipient_email, file_path)

print("Emails sent successfully!")
