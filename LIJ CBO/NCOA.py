import pandas as pd
import glob
from datetime import date as dt
from datetime import timedelta
import numpy as np
import win32com.client as win32
import os


def run():

    def truncate_file_name(file_name):
        """
        Takes the full file path and Returns a truncated file name
        :param file_name:
        """
        short_name = file_name.replace(f"{file_path}", "")
        return short_name


    file_path = 'M:/FPPShare/FPP-Production/NCOA BOT'

    today = dt.today()
    # if today is Monday or Tuesday
    if today.weekday() == 0 or today.weekday() == 1:
        # set file_date to the previous Thursday or Friday
        file_date = today - timedelta(days=4)
    else:
        # set file_date to two days ago
        file_date = today - timedelta(days=2)

    fd_YYYYMMDD = file_date.strftime('%Y%m%d')
    fd_MM_DD_YYYY = file_date.strftime('%m/%d/%Y')
    file = glob.glob(f'{file_path}/{fd_YYYYMMDD}*.xlsx')

    # Error check if the NCOA File exists
    try:
        if os.path.exists(file[0]):
            file = glob.glob(f'{file_path}/{fd_YYYYMMDD}*.xlsx')
            num_files = len(file)
            df_list = []

            # Add the dataframes to a list
            for f in file:
                df = pd.read_excel(f, engine="openpyxl")
                df = pd.DataFrame(df, columns=['PTFULLNAME', 'INVNUM', 'RetrievalStatus', 'RetrievalDescription', 'File'])
                df['File'] = truncate_file_name(f)
                df_list.append(df)

            # Concatenate the dataframe list to one dataframe
            df_all = pd.concat(df_list)

            # Set business status based on Retrieval Status
            df_all['Business Status'] = np.where(
                (df_all['RetrievalStatus'] == 'C00'),
                'Success',
                'Exception'
            )

            # Create second dataframe that will be the table w ithin the email
            df2 = pd.DataFrame(df_all,columns=['PTFULLNAME', 'RetrievalDescription', 'Business Status', 'File'])
            df2 = df2[df2['Business Status'] == 'Exception']
            df2.drop('Business Status', axis=1, inplace=True)
            df2 = df2.sort_values(by='RetrievalDescription', ascending=False)
            df2.rename(columns={'PTFULLNAME': 'Patient Name', 'RetrievalDescription': 'Exception Description'}, inplace=True)
            html_table = df2.to_html(index=False, classes="dataframe", border=2, justify="center")

            # Count total rows
            total_rows = len(df_all)

            # Calculate percentage rate for each Business Status
            business_status_counts = df_all['Business Status'].value_counts()
            business_status_percentage = business_status_counts / total_rows * 100
            business_status_percentage = business_status_percentage.round(2)

            # Retrieve the Success/Exception % or return 0 if no value can be found
            success_percentage = business_status_percentage.get('Success', 0.0)
            exception_percentage = business_status_percentage.get('Exception', 0.0)

            # Get a count for each StatusCode
            status_code_counts = df_all['RetrievalDescription'].value_counts()
            status_code_counts_text = "Count for each Retrieval Description:\n"
            business_status_rate_text = "Rate for each Business Status:\n"
            for index, count in status_code_counts.items():
                status_code_counts_text += f"{index} - {count}\n"
            for index, count in business_status_percentage.items():
                business_status_rate_text += f"{index} Rate - {count}%\n"

            # Start of body text for email; shows total num of files processed and total cases processed 
            
            html_body = f"""
            <p><strong>Total Files Processed:</strong> {num_files}<br>
            Links to File(s) can be found below:</p>
            """
            
            if len(file) > 1:
                for num in range(1,len(file)+1):
                    html_body += f"""<p><a href="file:///{file_path}/{fd_YYYYMMDD}.NSH.NCOA{num}_OUTBOUND.xlsx">Link to File Number {num}</a></p>"""
            else:
                html_body += f"""<p><a href="file:///{file_path}/{fd_YYYYMMDD}.NSH.NCOA_OUTBOUND.xlsx">Link to File</a></p>"""
            
            html_body += f"""
            <p><strong>Total Cases Processed:</strong> {total_rows}</p>
            <p><strong>Count for each Retrieval Description:</strong></p>
            """

            # Generate HTML for each item in status_code_counts
            for index, count in status_code_counts.items():
                html_body += f"{index}: {count}<br>"

            # Show total success rate and the list of exceptions using df2 from above
            html_body += f"""
            <p><strong>Rate for each Business Status:</strong><br>
            Success - {success_percentage}%<br>
            Exception - {exception_percentage}%</p>
            <strong>List of Exceptions:</strong>
            {html_table}
            """


            # Generate the email in outlook and add the To, CC, Subject and Body
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.Subject = f'NCOA Daily Snapshot - {fd_MM_DD_YYYY}'
            mail.HTMLBody = html_body
            mail.To = 'denglish2@northwell.edu'
            mail.CC = 'rmuncipinto@northwell.edu; nmitrako@northwell.edu'

            mail.Send()
    
    # If there are not any NCOA files, it will error out and send a different email instead
    except IndexError:
        html_body = f"""
        <p><strong>There were not any NCOA files to be processed on:</strong> {fd_MM_DD_YYYY} </p>
        <p><strong>Thank you, <br>
        ORCCA Team</strong></p>
        """

        # Generate the email in outlook and add the To, CC, Subject and Body
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.Subject = f'NCOA Daily Snapshot - {fd_MM_DD_YYYY}'
        mail.HTMLBody = html_body
        mail.To = 'denglish2@northwell.edu'
        mail.CC = 'rmuncipinto@northwell.edu; nmitrako@northwell.edu'

        mail.Send()


if __name__ == '__main__':
    run()