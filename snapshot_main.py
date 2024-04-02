from mappings import mappings_dict
from snapshot_logger import logger
from snapshot_updated import *
from datetime import date as dt

if __name__ == '__main__':
    today = dt.today()
    dt_func = DateFunctions()
    date = dt_func.ask_if_correct_date(today)
    try:
        outbound_df = MainSpreadsheet(init_date=date)
    except FileNotFoundError:
        logger.critical(f"Main outbound spreadsheet not found in M:\CPP-Data\Sutherland RPA\Combined Outputs")
    else:
        for use_case in mappings_dict:
            if use_case == "NCOA":
                continue

        # Use the below line to run Ad Hoc daily snapshot
            # "CSE1235"
            # "CSE1236"
            # "TES1249"
            # "TES6146"
            # "MCD MCO Available"
            # "MCR Advantage"
            # "MCR PartB Inactive"
            # "Medicare Not Primary"
            # "LIJ IS Printing"
            # "NCOA"
            # "BD IS Printing"

            else:
                snapshot = Snapshot(use_case, outbound_df.main_df, date=date)
                try:
                    snapshot.parse_spreadsheet()
                    if len(snapshot.use_case_df) == 0:
                        logger.critical(f"No rows in spreadsheet for {use_case}")
                        continue
                    snapshot.apply_business_rules()
                    snapshot.get_exceptions()
                except AttributeError:
                    logger.critical(f"{use_case} encountered an AttributeError error")
                    continue
                else:
                    snapshot.calc_results()
                    snapshot.write_email_body()
                    snapshot.compose_and_send()
            
        