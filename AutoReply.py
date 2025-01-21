import email
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta
import pandas as pd
from time import time,sleep
import os
from datetime import datetime
import imaplib


#  IMAP CONNECTIVITY
def connect_to_imap(username, password):
    imap_server = "imap.gmail.com"
    imap_port = 993
    mail = imaplib.IMAP4_SSL(imap_server, imap_port)
    mail.login(username, password)
    return mail

#  GET UNREAD EMAILS
def fetch_unread_emails(mail, sender_email, days=0):
    mail.select("inbox")
    date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
    _, message_numbers = mail.search(None, f'(UNSEEN FROM "{sender_email}" SINCE "{date}")')
    
    # print(message_numbers)
    # exit()
    
    return message_numbers[0].split()

#  GET RECIEVING TIME FROM HEADER 
def extract_receiving_time(original_message):
    date_header = original_message['Date']   
    if date_header:
        try:
            receiving_time = datetime.strptime(date_header, '%a, %d %b %Y %H:%M:%S %z').time()
        except ValueError:
            receiving_time = None
        return receiving_time
    else:
        return None

#  READ EMAIL
def parse_email(mail, email_id):
    status, msg_data = mail.fetch(email_id, "(RFC822)")
    if status != 'OK':
        raise Exception(f"Failed to fetch email with ID {email_id}")
    
    raw_email = msg_data[0][1]
    email_message = email.message_from_bytes(raw_email)    
    return email_message

# REPLY 
def send_email_reply(sender_email, sender_password, to_email, subject, reply_message, client_name,original_message,protection_verification):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = to_email
    message["Subject"] = f"Re: {subject}"
    
    if protection_verification == 'YES':
        addedNote = "Please Buy Protection of 2-10 Rupees of 500 Points far away CE and PE,\nIn Case there is rejection of Margin Shortfall in the above trade(s)"
    else:
        addedNote = ''
    
    reply_text = f"\n{reply_message}\n{addedNote}\n\nRegards,\n{client_name}\n---------------------------------------------------------\n\n"
    reply_text = reply_text + f"{original_message['Date']}\n{original_message['From']}\n\n"
    
    if original_message.is_multipart():
        for part in original_message.walk():
            if part.get_content_type() == "text/plain":
                reply_text += part.get_payload(decode=True).decode('utf-8', errors='replace')
                break

    message.attach(MIMEText(reply_text, "plain"))
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(message)
    

# Processing all email ids
def process_client_emails(username, password, reply_message,protection_verification,sender_email):
    
    reply_sent = False
    mail = connect_to_imap(username, password)
    unread_emails = fetch_unread_emails(mail, sender_email)
    client_df = pd.read_csv('./Account Info/account.csv')[['Email Id', 'Name']]
    client_df = client_df.set_index('Email Id')
    
    if unread_emails:
        print(username, end= ' ')
        
        for email_id in unread_emails:
            email_message = parse_email(mail, email_id)
            
            if 'Pre-Trade Confirmation' in email_message['Subject']\
            or 'Post-Trade Confirmation' in email_message['Subject']\
            or 'trade confirmation' in email_message['Subject'].lower():
                
                send_email_reply(username, password, sender_email, email_message['Subject'], reply_message, client_df.loc[username]['Name'],email_message,protection_verification)
                receiving_time = extract_receiving_time(email_message)
                log_df.loc[len(log_df)] = [username, password, sender_email, receiving_time, datetime.now().strftime("%H:%M:%S"),reply_message,True]
                print(": Success", end= ' ')
                reply_sent = True

        print()
    mail.logout()  
    return reply_sent

# EMAIL CREDENTIALS
sender_email = "discipline_ops@bp.sharekhan.com"
AccountDf = pd.read_csv('./Account Info/account.csv')[['Email Id', 'Passkey', 'Message', 'Name','Protection Message']].dropna(how='any').reset_index()
AccountClientList = AccountDf['Email Id'].tolist()

# LOG FILE
log_main_folder = './Log File'
if not os.path.exists(log_main_folder):
    os.makedirs(log_main_folder)

log_folder_name = f'{datetime.now().strftime("%d-%m-%Y")}'
log_file_name = f'{datetime.now().strftime("%H-%M")}'
log_df = pd.DataFrame(columns=['Email','Passkey','Sender','ReceivingTime','Sent Time','Message','Status'])

if not os.path.exists(os.path.join(log_main_folder, log_folder_name)):
    os.makedirs(os.path.join(log_main_folder, log_folder_name))

MasterClientDf = pd.read_csv('./master.csv')
MasterClientDf['Reply Check'] = 3

MasterClientList = MasterClientDf[MasterClientDf['Reply Check'] != 0]["Email id"].tolist()
client_email = MasterClientList.pop(0)

# REPLY LOOP
while client_email:
    
    CheckStatus = False
    print()
    print(client_email)
    
    if client_email  in AccountClientList:
        
        try:
            CheckStatus = process_client_emails(client_email, AccountDf[AccountDf['Email Id'] == client_email]['Passkey'].iloc[0], AccountDf[AccountDf['Email Id'] == client_email]['Message'].iloc[0],AccountDf[AccountDf['Email Id'] == client_email]['Protection Message'].iloc[0], sender_email)    
        except Exception as e:
            print(f'EXCEPTION In {client_email}: {str(e)}')

        if CheckStatus == False:
            MasterClientDf.loc[MasterClientDf['Email id'] == client_email, 'Reply Check'] -= 1
            print(f'Retrying in {client_email}')
        else:
            MasterClientDf.loc[MasterClientDf['Email id'] == client_email, 'Reply Check'] = 0
    
    else:
        MasterClientDf.loc[MasterClientDf['Email id'] == client_email, 'Reply Check'] = 0
        print(f'{client_email} NOT IN AccountFile')
    
    if len(MasterClientList) == 0: 
        MasterClientList = MasterClientDf[MasterClientDf['Reply Check'] != 0]['Email id'].tolist()
    
    client_email = MasterClientList.pop(0) if len(MasterClientList) != 0 else 0
   

# SAVING LOG FILE 
if len(log_df) > 0:
    log_df[log_df['Status'] == True].to_csv(os.path.join(os.path.join(log_main_folder, log_folder_name), \
                                        f'{log_file_name}_success.csv'), index=False)