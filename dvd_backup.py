import logging
import subprocess
import time
from collections import namedtuple
from pathlib import Path
from configparser import ConfigParser
import hashlib

import pywintypes
import win32api
import win32file

import fire

TIMEOUT_SLEEP = 5

VolumeInformation = namedtuple("VolumeInformation", "Label info1 info2 info3 type")


def get_volume_label(*args, **kwargs):
    volume_information = VolumeInformation(*win32api.GetVolumeInformation(*args, **kwargs))
    return volume_information.Label


def test_drive(current_letter: str):
    try:
        win32file.GetDiskFreeSpaceEx(current_letter)
    except win32api.error as err:
        if 'The device is not ready.' not in err.args:
            logging.error(err)
            raise err
        is_drive_available = False
    else:
        is_drive_available = True
    return is_drive_available


def save_to_iso(img_burn_exe: str, drive: str, output_file: str):
    params = [img_burn_exe, "/MODE", "READ", "/SRC", drive, "/DEST",
              output_file, "/EJECT", "/START", "/CLOSE", "/WAITFORMEDIA", "/LOG",
              r"C:\temp\test.log", "/LOGHEADER"]
    proc = subprocess.Popen(params, shell=True)
    proc.wait()
    status = proc.poll()
    logging.debug(f"Status: {status}")
    if status:
        raise ValueError(f"Status error: {status}")
    else:
        logging.debug("All fine")


def list_files_in_drive(drive_letter):
    return list(Path(drive_letter).rglob("*.*"))


# noinspection PyUnresolvedReferences
def process_drive(img_burn_exe: str, drive: str, output_folder: str):
    if not test_drive(drive):
        print(f"Drive is not ready yet...")
        return
    try:
        label = get_volume_label(f"{drive}\\")
        files = list_files_in_drive(drive)
        autorun_file = Path(f"{drive}Autorun.inf")
        if autorun_file.is_file():
            parser = ConfigParser()
            parser.read_string(autorun_file.read_text(encoding='utf-8').lower())
            if 'label' in parser['autorun']:
                autorun_label = parser['autorun']['label'].upper()
            else:
                autorun_label = "NO_AUTORUN_LABEL"
        else:
            autorun_label = "NO_AUTORUN_LABEL"

        logging.info(f"autorun label: {autorun_label}")
        logging.info(f"Standard label: {label}")
        text_for_files = "\r\n".join(str(file) for file in files)
        logging.info(f"List of files: {text_for_files}")
        files_hash = hashlib.md5(text_for_files.encode('utf-8')).hexdigest()
        iso_folder = Path(output_folder) / f"{label}_{files_hash}"
        logging.info(f"Iso folder: {iso_folder}")
        iso_folder.mkdir(parents=True)
        output_file = str(iso_folder / Path(f"{label}_{autorun_label}").with_suffix(".iso"))
        save_to_iso(img_burn_exe, drive, output_file=output_file)
        (iso_folder / "list_of_files.txt").write_text(data=text_for_files)

    except pywintypes.error as err:
        print(err)
    except ValueError as err:
        print(err)


def poll_drive_for_backup(img_burn_exe: str, drive: str, output_folder: str):
    while True:
        process_drive(img_burn_exe, drive, output_folder=output_folder)
        time.sleep(TIMEOUT_SLEEP)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format="[%(asctime)s]:[%(name)-12s] [%(levelname).1s]: %(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S")
    fire.Fire(poll_drive_for_backup)
