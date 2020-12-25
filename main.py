import time
from collections import namedtuple
from pathlib import Path

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
            drive_info[label] = list_files_in_drive(drive)
        except pywintypes.error as err:
            print(err)

        time.sleep(TIMEOUT_SLEEP)


if __name__ == '__main__':
    main()
