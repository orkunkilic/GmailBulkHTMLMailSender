# Gmail Bulk HTML Mail Sender

This python program sends HTML mails to your receiver excel list. You can add new columns (i.e. names) to this excel list in order to create customized HTML mails.

## Installation & Usage

1.  Install the requirements of the project by

    `pip install -r requirements.txt`

2.  Create your **client_secret.json** file by creating new project on Google API Console. Please follow the official tutorial: [https://developers.google.com/identity/protocols/oauth2/web-server#creatingcred](https://developers.google.com/identity/protocols/oauth2/web-server#creatingcred)
3.  Create your **list.xlsx** file. First column should be the mail addresses. You can add other columns and use them in your template.
4.  When your **list.xlsx** and **client_secret.json** file are ready, run:

    ` python sendmail.py`

5.  You are done. App should close automatically at the end.
