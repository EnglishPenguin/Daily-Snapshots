from snapshot import Snapshot
from coding_mappings import coding_dict

if __name__ == '__main__':
    for use_case in coding_dict:
        Snapshot(use_case)