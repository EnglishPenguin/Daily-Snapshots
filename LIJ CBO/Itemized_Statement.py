import pandas as pd
import glob
from datetime import date as dt
from datetime import timedelta
import numpy as np
import matplotlib.pyplot as plt
import win32com.client as win32


def run():
    file_path = 'M:/FPPShare/FPP-Production/Itemized Statement HCX BOT'

    today = dt.today()
    # if today is Monday
    if today.weekday() == 0:
        # set FILE_DATE to the previous Friday
        file_date = today - timedelta(days=3)
    else:
        # set FILE_DATE to Yesterday
        file_date = today - timedelta(days=1)

    fd_YYYYMMDD = file_date.strftime('%Y%m%d')
    fd_MM_DD_YYYY = file_date.strftime('%m/%d/%Y')

    file = glob.glob(f'{file_path}/{fd_YYYYMMDD}*.xlsx')
    for f in file:
        df = pd.read_excel(f, engine="openpyxl")
        df = pd.DataFrame(df, columns=['POLICYID', 'RetrievalStatus', 'RetrievalDescription', 'StatusCode'])

    df['Business Status'] = np.where(
        (df['StatusCode'] == 'ISP'),
        'Success',
        'Exception'
    )

    # Concatenate StatusCode and RetrievalDescription
    df['StatusCode_Description'] = df['StatusCode'] + ': ' + df['RetrievalDescription']

    df2 = pd.DataFrame(df,columns=['POLICYID', 'StatusCode_Description', 'Business Status'])
    df2 = df2[df2['Business Status'] == 'Exception']
    df2 = df2.sort_values(by='StatusCode_Description', ascending=False)
    df2.rename(columns={'POLICYID': 'Account Number', 'StatusCode_Description': 'Exception Description'}, inplace=True)
    html_table = df2.to_html(index=False, classes="dataframe", border=2, justify="center")

    # Count total rows
    total_rows = len(df)

    # Calculate percentage rate for each Business Status
    business_status_counts = df['Business Status'].value_counts()
    business_status_percentage = business_status_counts / total_rows * 100
    business_status_percentage = business_status_percentage.round(2)

    # Create a bar graph for the different Status Codes
    status_code_counts = df['StatusCode'].value_counts()
    fig, ax = plt.subplots(figsize=(10, 8))
    plt.subplots_adjust(bottom=0.2)
    status_code_counts.plot(kind='bar', ax=ax)
    plt.xlabel('Status Code')
    plt.ylabel('Count')
    plt.title('Counts by Status Code')
    for i, v in enumerate(status_code_counts):
        ax.text(i, v, str(v), ha='center', va='bottom')
    # Specify the file path and name
    bar_graph_filename = f'U:/zORCCA TEAM/Daily Analysis/LIJ CBO/{fd_YYYYMMDD}_Itemized_Statement.png'
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
    <p><strong>Rate for each Business Status:</strong></p>
    <p>Success Rate - {business_status_percentage['Success']}%</p>
    <p>Exception Rate - {business_status_percentage['Exception']}%</p><br>
    <strong>List of Exceptions:</strong>
    {html_table}
    """


    # Compose the email
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = f'Itemized Statements Daily Snapshot - {fd_MM_DD_YYYY}'
    mail.HTMLBody = html_body
    mail.To = 'denglish2@northwell.edu'
    mail.CC = 'rmuncipinto@northwell.edu'

    # Attach the bar graph file
    attachment = mail.Attachments.Add(Source=bar_graph_filename)
    mail.Send()

if __name__ == '__main__':
    run()