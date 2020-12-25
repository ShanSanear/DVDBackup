import time
from collections import namedtuple
from pathlib import Path
from configparser import ConfigParser

import pywintypes
import win32api

import win32file

drive_info = {}

TIMEOUT_SLEEP = 5

VolumeInformation = namedtuple("VolumeInformation", "Label info1 info2 info3 type")


def get_volume_label(*args, **kwargs):
    volume_information = VolumeInformation(*win32api.GetVolumeInformation(*args, **kwargs))
    return volume_information.Label


def test_drive(current_letter):
    try:
        win32file.GetDiskFreeSpaceEx(current_letter)
    except win32api.error as err:
        print(err)
        is_drive_available = False
    else:
        is_drive_available = True
    return is_drive_available


def list_files_in_drive(current_letter):
    return list(Path(current_letter).rglob("*.*"))


# noinspection PyUnresolvedReferences
def main():
    drive = "D:"
    while True:
        print(test_drive(drive))
        try:
            label = get_volume_label(f"{drive}\\")
            files = list_files_in_drive(drive)
            drive_info[label] = files
            autorun_file = Path(f"{drive}Autorun.inf")
            if autorun_file.is_file():
                    parser = ConfigParser()
                    parser.read_string(autorun_file.read_text(encoding='utf-8'))
                    autorun_label = parser['autorun']['label']
                    print(f"autorun label: {autorun_label}")
                    print(f"Standard label: {label}")
                    print(f"List of files: {files}")

        except pywintypes.error as err:
            print(err)

        time.sleep(TIMEOUT_SLEEP)


if __name__ == '__main__':
    main()
