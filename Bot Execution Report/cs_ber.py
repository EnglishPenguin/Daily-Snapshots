import pandas as pd
import win32com.client as win32
from bs4 import BeautifulSoup
import numpy as np
from datetime import datetime as dt


def run():
    # save path for the Bot Execution Report Archive
    FILEPATH = "M:/CPP-Data/CBO Westbury Managers/LEADERSHIP/Bot Folder/Dashboards/data/Bot Execution Reports"


    def concat_and_sort(df, combined_df, column, brick):
        if brick == 'even':
            check = 0
        elif brick == 'odd':
            check = 1
        df = pd.concat([df.reset_index(drop=True) for i, df in enumerate(combined_df) if i % 2 == check])
        # print(df_copy[column])
        df[column] = pd.to_datetime(df[column], errors='coerce')
        df[column] = df[column].dt.strftime('%m/%d/%Y')
        df = df.sort_values(by=column)
        return df


    # Connect to Microsoft Outlook
    outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')

    # Access the inbox folder that contains the Bot Execution Report emails from PLATFORMOPS @ Sutherland
    inbox = outlook.GetDefaultFolder(6).Folders['BER']  # Change the index if needed

    # get the message in that folder
    messages = inbox.Items

    # initialize the list that dataframes will be added to
    all_dfs = []

    # Column order and column datatype for the Claim Status dataframe
    table1_columns = [
        'Bot Name', 
        'Input Date', 
        'Batch ID', 
        'Total Downloaded', 
        'Processed', 
        'Pending', 
        'Response'
        ]
    table1_data_types = {
        'Bot Name': str, 
        'Input Date': 'datetime64[ns]', 
        'Batch ID': str, 
        'Total Downloaded': int, 
        'Processed': int, 
        'Pending': int, 
        'Response': str
        }

    for mail in messages:
        # Debug for pulling date of each file
        received_time = mail.ReceivedTime
        received_time = received_time.strftime("%m/%d/%Y")
        # print(received_time)

        table_dictionary = [
        np.nan, 
        f'{received_time}',
        np.nan,
        0,
        0, 
        0, 
        np.nan
        ]

        # Pull email body and uses BeautifulSoup to isolate the tables
        body = mail.HTMLBody
        soup = BeautifulSoup(body, 'lxml')
        tables = soup.find_all('table')
        table1_html = str(tables[0])
        
        # Create a BeautifulSoup object for each table HTML
        soup1 = BeautifulSoup(table1_html, 'lxml')

        # Convert the HTML tables to pandas DataFrames
        table1_df = pd.read_html(str(soup1), header=0)[0]

        # Rename columns to conform to the desired column names
        table1_df = table1_df.rename(columns=dict(zip(table1_df.columns, table1_columns)))

        # Replace "-" with NaN in table1_df
        for column, dictionary in zip(table1_df.columns, table_dictionary):
            table1_df[column] = table1_df[column].replace("-", dictionary)

        # Set the datatype of the Batch ID column to the desired type
        table1_df['Batch ID'] = 'ID ' + table1_df['Batch ID'].astype(str)

        # Add cs dataframes to list
        all_dfs.append(table1_df)

    # set the writer for the file. appends to an already created file
    writer = pd.ExcelWriter(f'{FILEPATH}/Bot Execution Report - RAW Data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')

    cs_df = pd.DataFrame(columns=table1_columns)

    # Concatenate all even data frames into one
    cs_df = concat_and_sort(df=cs_df, combined_df=all_dfs, column='Input Date', brick='even')

    cs_df['Input Date'] = pd.to_datetime(cs_df['Input Date'], errors='coerce', dayfirst=False)

    # Write the even data frame to the "Claim Status Bots" sheet
    cs_df.to_excel(writer, sheet_name='Claim Status Bots', index=False)

    # saves and exits the file
    writer.close()

if __name__ == '__main__':
    run()
