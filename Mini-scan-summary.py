import pandas as pd
import os
import glob


import smtplib
# MIMEMultipart send emails with both text content and attachments.
from email.mime.multipart import MIMEMultipart
# MIMEText for creating body of the email message.
from email.mime.text import MIMEText
# MIMEApplication attaching application-specific data (like CSV files) to email messages.
#from email.mime.application import MIMEApplication
#import Sqlite_export_to_csv as dataframe

from datetime import date

def get_date():
     today = date.today().strftime('%Y-%m-%d')
     #print(today)
     return(today)


def prepare_stats():
       
    today = get_date()
    os.chdir('C:\\Users\\ryanford\\OneDrive - SLAC National Accelerator Laboratory\\2025\\Python\\Attachments\\')
    df = pd.DataFrame()
    for name in glob.glob('C:\\Users\\ryanford\\OneDrive - SLAC National Accelerator Laboratory\\2025\\Python\\Attachments\\'+ today + '*'):
        data = pd.read_excel(name)
        if not data.empty:
            df = df._append(data, ignore_index = True)
    table = pd.pivot_table(df, index=['HOST','TYPE', 'DATE'], values='DOSI_ID', aggfunc='count', margins=True, margins_name='Total').reset_index()
    table1 = table.rename(columns={'HOST':'Scanner', 'TYPE': 'Type', 'DATE':'Date', 'DOSI_ID':'Count'})
    table_html = table1.to_html(index=False, col_space='150px', justify='center', bold_rows=True, border=1, columns=['Scanner', 'Type', 'Date','Count'])
    #table2 = pd.concat([y._append(y.sum().rename((x, 'SubTotal'))) for x, y in table.groupby(level=0)])._append(table.sum().rename(('Grand', 'Total')))
    #print(table_html)
    #print(table1)
    return(table_html)


def send_email():
        
    line_break = '<p>&#x000D;</p>'            
    email_header_0 = ("Scanner statistics for the past 10 days are shown below:\n")

    email_footer = ("\nIf you have any questions regarding the dosimetery service, "
                    "please contact ESH-DREP@SLAC.STANFORD.EDU." + line_break +
                    "Sincerely," + line_break + "Radiation Protection Dosimetry Group" + line_break +
                    "ODTSSCAN01 = Dosimetry Lab" + line_break +
                    "ODTSSCAN02 = Unassigned" + line_break +
                    "ODTSSCAN03 = Unassigned" + line_break +
                    "ODTSSCAN04 = Badging Office" + line_break +
                    "ODTSSCAN05 = Dosimetry Administrator's Office" + line_break)

    smtp_host = 'SMTPOUT.slac.stanford.edu'
    smtp_port = 25
    subject = 'ODTS Mini-Scan Statistics'
    table = prepare_stats()
    df_html = table #table.to_html(index=False, col_space='150px', justify='center', bold_rows=True, border=1)
    sender_email = 'ryanford@slac.stanford.edu'
    receiver_email = 'esh-DREP@slac.stanford.edu'

    #print(table)
    # MIMEMultipart() creates a container for an email message that can hold
    # different parts, like text and attachments and in next line we are
    # attaching different parts to email container like subject and others.

    message = MIMEMultipart()
    message['Subject'] = subject
    message['From'] = sender_email
    message['To'] = receiver_email
    message.attach(MIMEText(email_header_0 + line_break + df_html + line_break + email_footer, 'html'))
    
    try:
            with smtplib.SMTP(smtp_host, smtp_port, timeout = 5) as server:
                server.sendmail(sender_email, receiver_email, message.as_string())
                server.quit()

    except Exception as e:
            print(e)
            print(type(e))


get_date()
prepare_stats()
send_email()
