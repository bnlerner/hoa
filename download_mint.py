#! /bin/bash/python3

import mintapi
import imaplib
import email
import time
import datetime
import pandas as pd

import smtplib
import mimetypes
from email import encoders
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase

from passwords import *

SMTP_SERVER = "imap.gmail.com"
SMTP_PORT = 993

time_format = '%m-%Y'

def style_html(df, title=''):
    '''
    Write an entire dataframe to an HTML file with nice formatting.
    '''

    result = '''
<html>
<head>
<style>

    h2 {
        text-align: center;
        font-family: Helvetica, Arial, sans-serif;
    }
    table { 
        margin-left: auto;
        margin-right: auto;
    }
    table, th, td {
        border: 1px solid black;
        border-collapse: collapse;
    }
    th, td {
        padding: 5px;
        text-align: center;
        font-family: Helvetica, Arial, sans-serif;
        font-size: 90%;
    }
    table tbody tr:hover {
        background-color: #dddddd;
    }
    .wide {
        width: 90%; 
    }

</style>
</head>
<body>
    '''
    result += '<h2> %s </h2>\n' % title
    result += df.to_html(classes='wide', escape=False, index=False)
    result += '''
</body>
</html>
'''
    return result

def add_attachments(msg, files):

    for file_name in files:
        ctype, encoding = mimetypes.guess_type(file_name)
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'

        maintype, subtype = ctype.split("/", 1)

        if maintype == 'text':
            fp = open(file_name, 'r+', encoding='utf-8')
            attachment = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
        elif maintype == 'image':
            fp = open(file_name, 'rb')
            attachment = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
        elif maintype == 'audio':
            fp = open(file_name, 'rb')
            attachment = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
        else:
            fp = open(file_name, 'rb')
            attachment = MIMEBase(maintype, subtype)
            attachment.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(attachment)

        attachment.add_header('Content-Disposition', 'attachment', filename=file_name)

        msg.attach(attachment)

    return msg


def send_email(files, send_group, subject, body=None):

    COMMASPACE = ', '

    msg = MIMEMultipart()

    email_from = from_email

    msg['Subject'] = subject
    msg['From'] = email_from
    msg['To'] = COMMASPACE.join(send_group)
    msg.preamble = 'hi this is the preamble'

    if body:
        msg.attach(MIMEText(body, 'html'))

    msg = add_attachments(msg, files)

    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()

    user = email_user
    password = email_pass
    server.login(user, password)
    ret = server.sendmail(email_from, send_group, msg.as_string())

    server.quit()
    return ret


def read_email_from_gmail(error_message):
    time.sleep(5)

    mail = imaplib.IMAP4_SSL(SMTP_SERVER, SMTP_PORT)
    mail.login(email_user, email_pass)
    mail.select()

    typ, message_numbers = mail.search(None, 'ALL')

    num = message_numbers[0].split()[-1]
    typ, msg = mail.fetch(num, '(RFC822)')
    msg = email.message_from_bytes(msg[0][1])

    mail.close()
    mail.logout()

    for part in msg.walk():

        def find_num(to_find_msg):
            for word in to_find_msg:
                try:
                    num = int(word)
                    if len(word) == 6:
                        return num
                except:
                    pass
            return None


        msg_type = part.get_content_type()
        if "multipart" in msg_type:
            continue
        elif "text/html" in msg_type:
            final_msg = part.get_payload(decode=1).decode().replace('\n', '').replace(' ', '').split('\r')
            final_num = find_num(final_msg)
            if final_num:
                return final_num
        else:
            final_msg = part.get_payload(decode=1).decode().split()
            final_num = find_num(final_msg)
            if final_num:
                return final_num

    raise RuntimeError('did not find the mfa')

def get_mint_info():

    # A callback accepting a single argument (the prompt)
    # which returns the user-inputted 2FA code. By default
    # the default Python `input` function is used.
    mint = mintapi.Mint(email_user, mint_pwd, mfa_method='email',
                        headless=True, mfa_input_callback=read_email_from_gmail)

    # Initiate an account refresh
    mint.initiate_account_refresh()

    # Get basic account information
    #accounts = mint.get_accounts()

    nw = mint.get_net_worth()

    hoa_account = 'BUSINESS CHECKING'

    # Get transactions
    df = mint.get_transactions(include_investment=False) # as pandas dataframe

    return df[df['account_name'] == hoa_account], nw


if __name__ == "__main__":
    df, net_worth = get_mint_info()

    df.sort_values(by='date', inplace=True, ascending=False)
    df = df[['date', 'description', 'original_description', 'amount', 'transaction_type', 'category']]

    df.sort_values(by=['date'], inplace=True, ascending=False)
    csv_name = 'hoa_account.csv'
    df.to_csv(csv_name, index=False)

    today = datetime.datetime.now()
    one_month = today.replace(month=today.month - 1)
    df = df[df.date >= one_month]
    df.sort_values(by=['transaction_type', 'amount', 'category'], ascending=[True, False, False]).reset_index(
        inplace=True)

    html = list()

    df_cash_flow = df.groupby(['transaction_type']).sum()
    tot_cash_flow = df_cash_flow.loc['credit', :] - df_cash_flow.loc['debit', :]
    df_cat = df.groupby(['category', 'transaction_type']).sum()

    df_i = pd.Series(name=('Total', 'Total'))
    df_i['amount'] = tot_cash_flow.values[0]
    df_cat = df_cat.append(df_i)
    df_cat.reset_index(inplace=True)

    """
    df_html_output = df.style.set_table_styles(
        [{'selector': 'thead th',
          'props': [('background-color', 'red')]},
         {'selector': 'thead th:first-child',
          'props': [('display', 'none')]},
         {'selector': 'tbody th:first-child',
          'props': [('display', 'none')]}]
    ).render()
    """

    html.append(style_html(df_cat, title='Cash flow by category'))
    html.append(style_html(df, title='All transactions for month'))

    body = '\r\n\n<br>'.join('%s' % item for item in html)

    files = (csv_name,)
    month = datetime.datetime.strftime(datetime.datetime.now(), time_format)
    subject = "544 waller HOA finance data, sent for month {}, balance: ${}, cash flow: ${}".format(month, net_worth, tot_cash_flow.values[0])

    ret = send_email(files, hoa_emails, subject, body)
    print(ret, 'complete!', datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d %H:%M%:%S"))
