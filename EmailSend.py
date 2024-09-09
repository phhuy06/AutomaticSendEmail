import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os

load_dotenv()

excel_file = os.getenv("excel_path")
df = pd.read_excel(excel_file)

recipient_email = [] #list email 
recipient_name = [] #list name
recipient_SBD = [] #list SBD

for index, row in df.iterrows():
    recipient_email_inexcel = row["Email"]
    recipient_name_inexcel = row["Name"]
    recipient_SBD_inexcel = row["SBD"]
    recipient_email.append(recipient_email_inexcel)
    recipient_name.append(recipient_name_inexcel)
    recipient_SBD.append(recipient_SBD_inexcel)


html = """\
<html>
  <head></head>
  <body>
    <p style="color: red;">
        Body
    </p>
  </body>
</html>
"""

subject = "header" #header email
body = "body" #body email
sender = os.getenv("sender") #sender email
password = os.getenv("password") #application password

cnt = 0


def send_email(subject, body, sender, recipient_email, password, cnt):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = recipient_email
    part2 = MIMEText(html, 'html')
    msg.attach(part2)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
       smtp_server.login(sender, password)
       smtp_server.sendmail(sender, recipient_email, msg.as_string())
    print("Message sent " + str(cnt))

for recipient in recipient_email:
    cnt = cnt + 1
    send_email(subject, body, sender, recipient, password, cnt)
print("Done!!")

