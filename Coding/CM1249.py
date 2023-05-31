import pandas as pd
import glob
from datetime import date as dt
from datetime import timedelta
import numpy as np
import matplotlib.pyplot as plt
import win32com.client as win32


def run():
    file_path = 'M:/CPP-Data/Sutherland RPA/Coding/TES1249'

    today = dt.today()
    # if today is Monday
    if today.weekday() == 0:
        # set FILE_DATE to the previous Friday
        file_date = today - timedelta(days=3)
    else:
        # set FILE_DATE to Yesterday
        file_date = today - timedelta(days=1)

    fd_MMDDYYYY = file_date.strftime('%m%d%Y')
    fd_MM_DD_YYYY = file_date.strftime('%m/%d/%Y')

    file = glob.glob(f'{file_path}/*Outbound_{fd_MMDDYYYY}*')
    for f in file:
        df = pd.read_excel(f, engine="openpyxl")
        df = pd.DataFrame(df, columns=['INVNUM', 'CRN#', 'RetrievalStatus', 'RetrievalDescription'])

    df['Business Status'] = np.where(
    (df['RetrievalDescription'] == 'C00-E/M Changed to 99024') |
    (df['RetrievalDescription'] == 'C00-Modifier Added to E/M') ,
    'Success',
    'Exception'
    )

    # Concatenate StatusCode and RetrievalDescription
    df['StatusCode_Description'] = df['RetrievalDescription']

    df2 = pd.DataFrame(df,columns=['INVNUM', 'CRN#', 'StatusCode_Description', 'Business Status'])
    df2 = df2[df2['Business Status'] == 'Exception']
    df2 = df2.sort_values(by='StatusCode_Description', ascending=False)
    df2.rename(columns={'INVNUM': 'Encounter Number', 'CRN#': 'Transaction Number', 'StatusCode_Description': 'Exception Description'}, inplace=True)
    html_table = df2.to_html(index=False, classes="dataframe", border=2, justify="center")

    # Count total rows
    total_rows = len(df)

    # Calculate percentage rate for each Business Status
    business_status_counts = df['Business Status'].value_counts()
    business_status_percentage = business_status_counts / total_rows * 100
    business_status_percentage = business_status_percentage.round(2)

    # Create a bar graph for the different Status Codes
    status_code_counts = df['RetrievalDescription'].value_counts()
    # Set the figure size and margins    
    fig, ax = plt.subplots(figsize=(12, 9))
    plt.subplots_adjust(bottom=0.2)
    status_code_counts.plot(kind='bar', ax=ax)
    plt.xlabel('RetrievalDescription')
    plt.ylabel('Count')
    plt.title('Counts by RetrievalDescription')
    plt.xticks(rotation=20, ha='right', fontsize=6)
    # Add data labels to the bars
    for i, v in enumerate(status_code_counts):
        ax.text(i, v, str(v), ha='center', va='bottom')
    # Specify the file path and name
    bar_graph_filename = f'U:/zORCCA TEAM/Daily Analysis/Coding/{fd_MMDDYYYY}_CM1249.png'
    # Save the bar graph with the specified file path
    plt.savefig(bar_graph_filename)

    # Get a count for each StatusCode
    status_code_counts = df['StatusCode_Description'].value_counts()
    status_code_counts_text = "Count for each Retrieval Description:\n"
    business_status_rate_text = "Rate for each Business Status:\n"
    for index, count in status_code_counts.items():
        status_code_counts_text += f"{index} - {count}\n"
    for index, count in business_status_percentage.items():
        business_status_rate_text += f"{index} Rate - {count}%\n"

    html_body = f"""
    <p><strong>Total Cases Processed:</strong> {total_rows}</p>
    <p><strong>Count for each Retrieval Description:</strong></p>
    """

    # Generate HTML for each item in status_code_counts
    for index, count in status_code_counts.items():
        html_body += f"<p>{index}: {count}</p>"

    html_body += f"""
    <strong>Rate for each Business Status:</strong><br>
    Success Rate - {business_status_percentage['Success']}%<br>
    Exception Rate - {business_status_percentage['Exception']}%<br><br>
    <strong>List of Exceptions:</strong>
    {html_table}
    """


    # Compose the email
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = f'ClaimsManager 1249 Daily Snapshot - {fd_MM_DD_YYYY}'
    mail.HTMLBody = html_body
    mail.To = 'denglish2@northwell.edu'
    mail.CC = 'klawrence3@northwell.edu; ovicuna@northwell.edu; jcarney1@northwell.edu; gbunce@northwell.edu; mpuma@northwell.edu'

    # Attach the bar graph file
    attachment = mail.Attachments.Add(Source=bar_graph_filename)
    mail.Send()

if __name__ == '__main__':
    run()