from snapshot import Snapshot
from mappings import mappings_dict
from snapshot_logger import logger

if __name__ == '__main__':
    for use_case in mappings_dict:
    # for use_case in ["MCD MCO Available"]:
        try:
            Snapshot(use_case)
        except FileNotFoundError:
            logger.critical(f"File for {use_case} not found")
        