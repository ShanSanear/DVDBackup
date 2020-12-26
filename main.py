import subprocess
import time
from collections import namedtuple
from pathlib import Path
from configparser import ConfigParser
import hashlib

import pywintypes
import win32api
import win32file

TIMEOUT_SLEEP = 5

VolumeInformation = namedtuple("VolumeInformation", "Label info1 info2 info3 type")


def get_volume_label(*args, **kwargs):
    volume_information = VolumeInformation(*win32api.GetVolumeInformation(*args, **kwargs))
    return volume_information.Label


def test_drive(current_letter: str):
    try:
        win32file.GetDiskFreeSpaceEx(current_letter)
    except win32api.error as err:
        print(err)
        is_drive_available = False
    else:
        is_drive_available = True
    return is_drive_available


def save_to_iso(drive: str, output_file: Path):
    params = [r"C:\Programs\ImgBurn\ImgBurn.exe", "/MODE", "READ", "/SRC", drive, "/DEST",
              str(output_file), "/EJECT", "/START", "/CLOSE", "/WAITFORMEDIA", "/LOG",
              r"C:\temp\test.log", "/LOGHEADER"]
    proc = subprocess.Popen(params, shell=True)
    proc.wait()
    status = proc.poll()
    print(f"Status: {status}")


def list_files_in_drive(drive_letter):
    return list(Path(drive_letter).rglob("*.*"))


# noinspection PyUnresolvedReferences
def process_drive(drive: str, output_folder: Path):
    print(test_drive(drive))
    try:
        label = get_volume_label(f"{drive}\\")
        files = list_files_in_drive(drive)
        autorun_file = Path(f"{drive}Autorun.inf")
        if autorun_file.is_file():
            parser = ConfigParser()
            parser.read_string(autorun_file.read_text(encoding='utf-8'))
            autorun_label = parser['autorun']['label']
        else:
            autorun_label = "NO_AUTORUN_LABEL"

        print(f"autorun label: {autorun_label}")
        print(f"Standard label: {label}")
        print(f"List of files: {files}")
        text_for_files = "\r\n".join(str(file) for file in files)
        files_hash = hashlib.sha512(text_for_files)
        iso_folder = output_folder / f"{label}_{files_hash}"
        iso_folder.mkdir(parents=True)
        output_file = iso_folder / Path(f"{label}_{autorun_label}").with_suffix(".iso")
        save_to_iso(drive, output_file=output_file)
        (iso_folder / "list_of_files.txt").write_text(data=text_for_files)

    except pywintypes.error as err:
        print(err)


def main():
    drive = "D:"
    while True:
        process_drive(drive, output_folder=Path(r"C:\Users\Shan\Documents"))
        time.sleep(TIMEOUT_SLEEP)


if __name__ == '__main__':
    main()
