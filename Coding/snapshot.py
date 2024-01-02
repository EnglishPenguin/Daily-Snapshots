import pandas as pd
import glob
from datetime import date as dt
from datetime import timedelta
import numpy as np
import win32com.client as win32
import coding_mappings
from coding_logger import logger

class Snapshot():
    def __init__(self, use_case):
        self.use_case = use_case
        self.file_path = coding_mappings.coding_dict[f"{self.use_case}"]["file_path"]
        self.status_crosswalk = coding_mappings.coding_dict[f"{self.use_case}"]["status_crosswalk"]
        self.scenario_crosswalk = coding_mappings.coding_dict[f"{self.use_case}"]["scenario_crosswalk"]
        self.columns = coding_mappings.coding_dict[f"{self.use_case}"]["columns"]
        self.columns_crosswalk = coding_mappings.coding_dict[f"{self.use_case}"]["column_crosswalk"]
        self.cc_emails = coding_mappings.coding_dict[f"{self.use_case}"]["carbon_copy"]
        self.name_format_str = coding_mappings.coding_dict[f"{self.use_case}"]["name_format"]
        self.today = dt.today()
        self.today_dow = self.today.weekday()

        if self.today_dow == 0:
            self.file_date = self.today - timedelta(days=3)
        else:
            self.file_date = self.today - timedelta(days=1)

        logger.info(f'Starting Process for {self.use_case}')
        self.fd_MMDDYYYY = self.file_date.strftime('%m%d%Y')
        self.fd_MM_DD_YYYY = self.file_date.strftime('%m/%d/%Y')
        logger.info(f'File Date: {self.fd_MM_DD_YYYY}')
        self.month_str = dt.strftime(self.file_date, "%m")
        self.day_str = dt.strftime(self.file_date, "%d")
        self.year_str = dt.strftime(self.file_date, "%Y")

        self.name_format = self.name_format_str.format(
            file_path=self.file_path,
            month_str=self.month_str,
            day_str=self.day_str,
            year_str=self.year_str
        )
        logger.info('retrieving file')
        self.get_file()
        logger.info('applying business rules')
        self.apply_business_rules()
        logger.info('getting counts')
        self.get_results()
        logger.info('writing email body')
        self.write_email_body()
        logger.info('creating and sending email')
        self.compose_and_send_email()


    def get_file(self):
        self.df_snapshot = pd.read_excel(self.name_format, engine='openpyxl')
        self.df_snapshot = pd.DataFrame(self.df_snapshot, columns=self.columns)
        self.df_snapshot.rename(columns=self.columns_crosswalk, inplace=True)
        self.df_snapshot['RD + Reason'] = self.df_snapshot['Retrieval Description']+" - "+self.df_snapshot['Reason']

    def apply_business_rules(self):
        self.df_snapshot['Business Status'] = self.df_snapshot.apply(lambda row: self.get_business_status(row), axis=1)
        self.df_snapshot['Business Scenario'] = self.df_snapshot.apply(lambda row: self.get_business_scenario(row), axis=1)

    def get_business_status(self, row):
        # based on mappings status crosswalk, retrievel relevant business status
        # If Key not found in crosswalk, will return Unknown value
        return self.status_crosswalk.get(row['RD + Reason'], 'Unknown')
    
    def get_business_scenario(self, row):
        # based on mappings scenarios crosswalk, retrievel relevant business scenarios
        # If Key not found in crosswalk, will return Unknown value
        return self.scenario_crosswalk.get(row['RD + Reason'], 'Unknown')
    
    def get_results(self):
        self.total_rows = len(self.df_snapshot)

        self.business_status_counts = self.df_snapshot['Business Status'].value_counts()
        self.business_status_percentage = self.business_status_counts / self.total_rows * 100
        self.business_status_percentage = self.business_status_percentage.round(2)

        self.rd_reason_counts = self.df_snapshot['RD + Reason'].value_counts()

        self.get_exceptions()

    def get_exceptions(self):
        self.df_exceptions = pd.DataFrame(self.df_snapshot)
        self.df_exceptions = self.df_exceptions[self.df_exceptions['Business Status'] == 'Exception']
        self.df_exceptions.drop(columns=['Patient Name', 'Retrieval Status', 'Retrieval Description', 'Reason', 'RD + Reason', 'Business Status'], axis=1, inplace=True)
        self.df_exceptions = self.df_exceptions.sort_values(by='Business Scenario', ascending=False)
        self.html_table = self.df_exceptions.to_html(index=False, classes="dataframe", border=2, justify="center")
    
    def write_email_body(self):
        self.email_body = f"""
        <p><strong>Total Cases Processed:</strong> {self.total_rows}</p>
        <p>Link to File can be found: <a href="file:///{self.name_format}">here</a></p>
        <p><strong>Count for each Description - Reason:</strong></p>
        """
        for index, count in self.rd_reason_counts.items():
                self.email_body += f"<p>{index}: {count}</p>"
        
        self.email_body += f"""
        <strong>Rate for each Business Status:</strong><br>
        Success Rate - {self.business_status_percentage['Success']}%<br>
        Exception Rate - {self.business_status_percentage['Exception']}%<br><br>
        <strong>List of Exceptions:</strong>
        {self.html_table}
        """

    def compose_and_send_email(self):
        self.outlook = win32.Dispatch('Outlook.Application')
        self.mail = self.outlook.CreateItem(0)
        self.mail.Subject = f'{self.use_case} Daily Snapshot - {self.fd_MM_DD_YYYY}'
        self.mail.HTMLBody = self.email_body
        self.mail.To = 'denglish2@northwell.edu'
        self.mail.CC = self.cc_emails
        self.mail.Send()
