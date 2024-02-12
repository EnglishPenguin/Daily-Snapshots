from snapshot import Snapshot
from mappings import mappings_dict
from snapshot_logger import logger

if __name__ == '__main__':
    for use_case in mappings_dict:
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
    # for use_case in [""]:
        try:
            Snapshot(use_case)
        except FileNotFoundError:
            logger.critical(f"File for {use_case} not found")
        