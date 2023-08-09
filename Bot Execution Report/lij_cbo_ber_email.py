import pandas as pd
import win32com.client as win32
import numpy as np
from datetime import datetime as dt
from datetime import timedelta

def run():
    FILE_PATH = "M:/CPP-Data/CBO Westbury Managers/LEADERSHIP/Bot Folder/Dashboards/data/Bot Execution Reports"
    FILE_NAME = "Bot Execution Report - RAW Data.xlsx"
    COMB_FILE_PATH = f"{FILE_PATH}/{FILE_NAME}"
    NCOA_OUTBOUND = "M:/FPPShare/FPP-Production/NCOA BOT"
    IS_OUTBOUND = "M:/FPPShare/FPP-Production/Itemized Statement HCX BOT"

    df = pd.read_excel(COMB_FILE_PATH, sheet_name="Type B Bots", engine="openpyxl")

    df_ncoa = df[df['BotName'].str.contains('Patient.*Address.*Update*', regex=True)]
    df_statement = df[df['BotName'].str.contains('Printing.*Itemized*', regex=True)]

    df_list = [df_ncoa, df_statement]

    df_final = pd.concat(df_list)

    today = dt.today()

    # if today is Monday
    if today.weekday() == 0:
        # set FILE_DATE to today - 3
        file_date = today - timedelta(days=3)
    else:
        # set FILE_DATE to today - 1
        file_date = today - timedelta(days=1)

    # Set the time to 00:00:00.000
    file_date = file_date.replace(hour=0, minute=0, second=0, microsecond=0)
    #Generate a string of the date in MMDDYYYY format
    fd_MM_DD_YYYY = dt.strftime(file_date, format="%m/%d/%Y")

    # Set the values in Date column 
    df_final = df_final[df_final['Date'] == file_date]

    # Rename the BotName column and drop Batch ID column
    df_final.rename(columns={'BotName': 'Bot Name'}, inplace=True)
    df_final.drop('Batch ID', axis=1, inplace=True)

    # Update the names of the Bot Use Cases to include space for better readability
    df_final['Bot Name'] = df_final['Bot Name'].replace({
        'PatientAddressUpdateviaNCOA': 'Patient Address Update via NCOA',
        'PrintingItemizedStatement': 'Printing Itemized Statement'
    })

    # Create an HTML Table of the dataframe to include in the email
    html_table = df_final.to_html(index=False, classes="dataframe", border=2, justify="center")

    # Generate the email body
    html_body = f"""
    <p>Good Morning,</p>
    <p>See below for the status of the latest files for the Printing of Itemized Statements and Patient Address Update via NCOA use cases.</p>
    {html_table}
    <p><strong>Note(s):</strong> 
    <ul>
    <li>A file is still 'In Progress' if there is a number greater than 0 in the "Pending" column.</li>
    <li>Printing of Itemized Statements Output Files can be found <a href="file:///{IS_OUTBOUND}">here</a></li>
    <li>Patient Address Update via NCOA Output Files can be found <a href="file:///{NCOA_OUTBOUND}">here</a></li>
    </ul></p>
    """

    # Compose the email
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = f'Bot Execution Report - {fd_MM_DD_YYYY}'
    mail.HTMLBody = html_body
    mail.To = 'denglish2@northwell.edu'
    # mail.CC = 'rmuncipinto@northwell.edu; nmitrako@northwell.edu; ttaylor6@northwell.edu; dbenjamin3@northwell.edu; tclouden@northwell.edu; djoseph8@northwell.edu'

    # Attach the bar graph file
    # attachment = mail.Attachments.Add(Source=bar_graph_filename)
    mail.Send()

if __name__ == '__main__':
    run()