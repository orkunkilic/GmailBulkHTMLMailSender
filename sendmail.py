from webbrowser import Error
from Google import Create_Service
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

CLIENT_SECRET_FILE = 'client_secret.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)


def send(msg, to, subject):

    emailMsg = msg
    mimeMessage = MIMEMultipart()
    mimeMessage['to'] = to
    mimeMessage['subject'] = subject
    mimeMessage.attach(MIMEText(emailMsg, 'html'))
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()

    try:
        message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
        print(to + " ", message)
    except Exception as e:
        print("Error, ", e)




mail_list = pd.read_excel("list.xlsx", header=None)


for x in range(0, len(mail_list[0])):

    msg = f'<div><p>HTML mail to dear {mail_list[1][x]}</div>'

    subject = "Mail Subject"

    to = mail_list[0][x]

    send(msg, to, subject)
