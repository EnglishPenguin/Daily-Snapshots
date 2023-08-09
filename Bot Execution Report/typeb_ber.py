import pandas as pd
import win32com.client as win32


def run():
    # save path for the Bot Execution Report Archive
    FILEPATH = "M:/CPP-Data/CBO Westbury Managers/LEADERSHIP/Bot Folder/Dashboards/data/Bot Execution Reports"

    # Text to be excluded later
    PHISH = "External Email."
    TYPE_B = "Type B"
    NOTES = "Note"
    SIG_L5 = "Thank you for your cooperation."

    # Connect to Microsoft Outlook
    outlook = win32.Dispatch('Outlook.Application').GetNamespace('MAPI')

    # Access the inbox folder that contains the Bot Execution Report emails from PLATFORMOPS @ Sutherland
    inbox = outlook.GetDefaultFolder(6).Folders['BER']  # Change the index if needed

    # get the message in that folder
    messages = inbox.Items

    table2_columns = [
        'BotName', 
        'Date', 
        'Batch ID', 
        'Total Downloaded', 
        'Completed', 
        'Pending', 
        'Response File'
        ]

    table2_data_types = {
        'BotName': str, 
        'Date': 'datetime64[ns]', 
        'Batch ID': str, 
        'Total Downloaded': int, 
        'Completed': int, 
        'Pending': int, 
        'Response File': str
        }

    final_rows = []
    for mail in messages:
        body = mail.Body
        # print(body)

        # Find the start and end positions of the text to remove
        start_pos = body.find(PHISH)
        end_pos = body.find(TYPE_B)

        # Remove the unwanted text from the email body
        if start_pos != -1 and end_pos != -1:
            end_pos += len(TYPE_B)
            modified_body_1 = body[:start_pos] + body[end_pos:]
        else:
            modified_body_1 = body
        
        # Repeat the process but for the bottom of the email
        if "Below Bot" in modified_body_1:
            start_pos_2 = modified_body_1.find("Below Bot")
        elif NOTES not in modified_body_1:
            start_pos_2 = modified_body_1.find("<https:") 
        else:
            start_pos_2 = modified_body_1.find(NOTES)
        end_pos_2 = modified_body_1.find(SIG_L5)
        if start_pos_2 != -1 and end_pos_2 != -1:
            end_pos_2 += len(SIG_L5)
            modified_body_2 = modified_body_1[:start_pos_2] + modified_body_1[end_pos_2:]
        else:
            modified_body_2 = modified_body_1
        # print(modified_body_2)
        stripped_body = modified_body_2.replace(" ","")
        # print(stripped_body)

        # Split the string into lines and remove leading/trailing spaces from each line
        lines = [line.strip() for line in stripped_body.splitlines() if line.strip()]
        stripped_body = "\n".join(lines)
        # print(stripped_body)

        # Split the email body into individual lines
        lines = stripped_body.splitlines()

        # Group lines into chunks of 16 lines each
        chunks = [lines[i:i+7] for i in range(0, len(lines), 7)]

        # Join each chunk with "|" delimiter to create rows
        rows = ["|".join(chunk) for chunk in chunks]

        # Join rows with newlines to reconstruct the modified email body
        final_body = "\n".join(rows)
        # print(final_body)  
        final_body = final_body.replace("||","|")

        # print(final_body)  
        split_rows = final_body.splitlines()
        for row in range(1,len(split_rows)):
            final_rows.append(split_rows[row])
        # print(final_body)

    # Find the maximum number of columns in the 'final_rows' data
    max_columns = max(len(row.split("|")) for row in final_rows)

    # Adjust the number of columns in 'table2_columns' to match the maximum number of columns found
    table2_columns = table2_columns + ['Column{}'.format(i+1) for i in range(max_columns - len(table2_columns))]

    # Create a DataFrame from the 'final_rows' data
    df = pd.DataFrame([row.split("|") for row in final_rows], columns=table2_columns)

    # Replace None with blank strings ('')
    df = df.replace(to_replace=pd.NA, value='')
    df = df.iloc[:, :7]

    # Set the datatype of the Batch ID column to the desired type
    df['Batch ID'] = 'ID ' + df['Batch ID'].astype(str)

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce', dayfirst=False)

    # set the writer for the file. appends to an already created file
    writer = pd.ExcelWriter(f'{FILEPATH}/Bot Execution Report - RAW Data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='overlay')

    # Write the even data frame to the "Claim Status Bots" sheet
    df.to_excel(writer, sheet_name='Type B Bots', index=False)

    # saves and exits the file
    writer.close()

if __name__ == '__main__':
    run()