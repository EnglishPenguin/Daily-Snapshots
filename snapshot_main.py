from snapshot import Snapshot
from mappings import mappings_dict
from snapshot_logger import logger
import time

if __name__ == '__main__':
    for use_case in mappings_dict:
    # for use_case in ["MCR PartB Inactive", "Medicare Not Primary", "LIJ IS Printing", "NCOA", "BD IS Printing"]:
        time.sleep(1)
        try:
            Snapshot(use_case)
        except FileNotFoundError:
            logger.critical(f"File for {use_case} not found")
        