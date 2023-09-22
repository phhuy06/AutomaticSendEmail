import pandas as pd
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
import os
load_dotenv()

excel_file = os.getenv("excel_path")
df = pd.read_excel(excel_file)

recipient_email = [] #list email 
recipient_name = [] #list name
recipient_SBD = [] #list SBD

for index, row in df.iterrows():
    recipient_email_inexcel = row[os.getenv("field1")]
    recipient_name_inexcel = row[os.getenv("field2")]
    recipient_SBD_inexcel = row[os.getenv("field3")]
    recipient_email.append(recipient_email_inexcel)
    recipient_name.append(recipient_name_inexcel)
    recipient_SBD.append(recipient_SBD_inexcel)

subject = os.getenv("header") #header email
body = os.getenv("body") #body email
sender = os.getenv("sender") #sender email
password = os.getenv("password") #application password

def send_email(subject, body, sender, recipient_email, password):
    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipient_email)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
       smtp_server.login(sender, password)
       smtp_server.sendmail(sender, recipient_email, msg.as_string())
    print("Message sent!")


send_email(subject, body, sender, recipient_email, password)