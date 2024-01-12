import pandas as pd
from datetime import date as dt
from datetime import timedelta
import win32com.client as win32
import mappings
from snapshot_logger import logger
from glob import glob

class Snapshot():
    def __init__(self, use_case):
        self.use_case = use_case
        self.file_path = mappings.mappings_dict[f"{self.use_case}"]["file_path"]
        self.status_crosswalk = mappings.mappings_dict[f"{self.use_case}"]["status_crosswalk"]
        self.scenario_crosswalk = mappings.mappings_dict[f"{self.use_case}"]["scenario_crosswalk"]
        self.columns = mappings.mappings_dict[f"{self.use_case}"]["columns"]
        self.columns_crosswalk = mappings.mappings_dict[f"{self.use_case}"]["column_crosswalk"]
        self.cc_emails = mappings.mappings_dict[f"{self.use_case}"]["carbon_copy"]
        self.name_format_str = mappings.mappings_dict[f"{self.use_case}"]["name_format"]
        self.drop_columns = mappings.mappings_dict[f"{self.use_case}"]["drop_columns"]
        self.today = dt.today()
        self.today_dow = self.today.weekday()

        if self.today_dow == 0:
            self.file_date = self.today - timedelta(days=3)
        else:
            self.file_date = self.today - timedelta(days=1)

        if self.use_case == "NCOA" and self.file_date.weekday() == 0:
            self.file_date -= timedelta(days=3)
        elif self.use_case == "NCOA":
            self.file_date -= timedelta(days=1)


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
        try:
            self.get_file()
        except AttributeError:
            logger.critical(f"No NCOA File for {self.fd_MM_DD_YYYY}")
            logger.info("Sending no file found email")
            self.write_no_ncoa_file()
            self.compose_and_send_email()
        else:
            logger.info('applying business rules')
            self.apply_business_rules()
            logger.info('getting counts')
            self.get_results()
            logger.info('writing email body')
            self.write_email_body()
            logger.info('creating and sending email')
            self.compose_and_send_email()


    def get_file(self):
        if self.use_case == "NCOA":
            self.files_list = glob(self.name_format)
            if len(self.files_list) == 0:
                pass
            else:
                self.combine_ncoa_files()
        else:
            self.df_snapshot = pd.read_excel(self.name_format, engine='openpyxl', na_values=" ", keep_default_na=False)
        self.df_snapshot = pd.DataFrame(self.df_snapshot, columns=self.columns)
        self.df_snapshot.rename(columns=self.columns_crosswalk, inplace=True)
        self.df_snapshot['RD + Reason'] = self.df_snapshot['Retrieval Description']+" - "+self.df_snapshot['Reason']

    def combine_ncoa_files(self):
        # take glob list and read excel. Combine dataframes into 1
        logger.debug(f'Attempting to combine {self.use_case} files')
        self.df_list = []
        for f in self.files_list:
            df_temp = pd.read_excel(f, engine='openpyxl', na_values=" ", keep_default_na=False)
            self.df_list.append(df_temp)
        self.df_snapshot = pd.concat(self.df_list)

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
        self.df_exceptions.drop(columns=self.drop_columns, axis=1, inplace=True)
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
    
    def write_no_ncoa_file(self):
        self.email_body = f"""
        <p><strong>There were not any NCOA files to be processed on:</strong> {self.fd_MM_DD_YYYY} </p>
        <p><strong>Thank you, <br>
        ORCCA Team</strong></p>
        """

    def compose_and_send_email(self):
        self.outlook = win32.Dispatch('Outlook.Application')
        self.mail = self.outlook.CreateItem(0)
        self.mail.Subject = f'{self.use_case} Daily Snapshot - {self.fd_MM_DD_YYYY}'
        self.mail.HTMLBody = self.email_body
        self.mail.To = 'denglish2@northwell.edu'
        self.mail.CC = self.cc_emails
        self.mail.Send()
