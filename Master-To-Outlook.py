import pandas as pd
import win32com.client
from time import time,sleep
from datetime import datetime, timedelta

def convertDataIntoSpreadSheetFormat():
    df = pd.read_csv('./master.csv')
    returnDf = pd.DataFrame(columns=['Client Code', 'Client Name', 'Recipient', 'Scrip Name'])
    
    if len(df) > 0:
        listOfColumns = list(df.columns[4::4])
        allColumns = list(df.columns)

        for clientId in df[df.columns[0]].unique():
            try:
                clientName = df.set_index(df.columns[0]).loc[clientId][df.columns[1]]
                recipient = df.set_index(df.columns[0]).loc[clientId][df.columns[2]]
                
                returnVal = ''
                priceRange = None
                statusType = None
                qty = None
                exchangeType = ''
               
                for column in listOfColumns:
                    if not pd.isna(df.set_index(df.columns[0]).loc[clientId][column]):
                        if len(returnVal) > 0:
                            returnVal += '<br>'
                        index = allColumns.index(column)
                        qty = df.set_index(df.columns[0]).loc[clientId][column]
                        exchangeType = df.set_index(df.columns[0]).loc[clientId][allColumns[index-1]]
                        statusType = df.set_index(df.columns[0]).loc[clientId][allColumns[index+1]]
                        priceRange = df.set_index(df.columns[0]).loc[clientId][allColumns[index+2]]
                        returnVal += f'Ticker - {column}; Qty - {qty}; Exchange - {exchangeType}; Type - {statusType}; Price Range - {priceRange}.'
                
                if len(returnVal) > 0:
                    returnDf = pd.concat([
                        returnDf, pd.DataFrame({
                            'Client Code': [clientId],
                            'Client Name': [clientName],
                            'Recipient': [recipient],
                            'Scrip Name': [returnVal]
                        })
                    ])
                
            except Exception as e:
                print(f"Error processing client ID {clientId} {e}")
                print("Recheck master file and run it again")
                return None, allColumns
    
    return returnDf,allColumns


def initialize_outlook(timeout=60):
    start_time = time()
    while (time() - start_time) < timeout:
        try:
            outlook = win32com.client.Dispatch('Outlook.Application')
            outlook.Session.GetDefaultFolder(6)  # Check if the MAPI session is accessible
            return outlook
        except Exception as e:
            print("Outlook not ready yet. Waiting...")
            sleep(5)  # Wait before retrying
    
    print("Timeout reached. Outlook is not ready.")
    return None

def sendMailViaOutlook():
    
    df = pd.DataFrame()
    try:
        df,allColumns = convertDataIntoSpreadSheetFormat()
        # if(df.isna()):
        #     print("Error encountered while processing master file. Check column sequence below")
        
    except:
        print("Error in column sequence in master file")
    
    listOfClientCodes = list(df['Client Code'].values)
    df = df.set_index('Client Code')
    
    outlook = initialize_outlook()
    if outlook is not None:
        print("Outlook is successfully connected.")
    else:
        print("Failed to connect to Outlook.")

    try:
        accounts = outlook.Session.Accounts
        target_email = "discipline_ops@bp.sharekhan.com" 
        sending_account = None
        
        # mail = outlook.CreateItem(0)
        for account in accounts:
            if account.SmtpAddress == target_email:
                sending_account = account
                # mail.SendUsingAccount = sending_account
                break
        
        if not sending_account:
            print(f"No account found with email address {target_email}")
            return   
        
    except Exception as e:
        print(f"Failed to retrieve accounts: {e}")
        return    


    for clientCode in listOfClientCodes:
        try:
            clientName = df.loc[clientCode, df.columns[0]]
            emailID = df.loc[clientCode, df.columns[1]]
            scripName = df.loc[clientCode, df.columns[2]]
            
            if not isinstance(emailID, str) or '@' not in emailID:
                print(f"Invalid email ID for client {emailID}")
                continue
            
            mail = outlook.CreateItem(0)
            if(sending_account):
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, sending_account))

            mail.To = emailID
            current_time = datetime.now().time()
            end_time = datetime.strptime("16:30", "%H:%M").time()  

            if current_time <= end_time:
                mail.Subject = 'Pre-Trade Confirmation'
            else:
                mail.Subject = 'Post-Trade Confirmation'

            # mail.Subject = 'Trade Confirmation'
            mail.HTMLBody = f"""
                <p>Dear {clientName},</p>
                <p>Client Code - {clientCode}</p>
                <p>Kindly confirm the following Trades:</p>
                <p>{scripName}</p> <br>
                <p>Regards,<br>Team Gautam</p>
                """
            
            # mail.SendUsingAccount = sending_account

            mail.Send()
            print(f"Email from {sending_account} sent to {emailID}")
            # time.sleep(1)
        except Exception as e:
            print(f"Failed to send email from {sending_account} to {emailID}")

sendMailViaOutlook()
