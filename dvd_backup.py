import logging
import subprocess
import time
from collections import namedtuple
from pathlib import Path
from configparser import ConfigParser, DuplicateOptionError
import hashlib
from typing import List

import win32api
import win32file

from charset_normalizer import CharsetNormalizerMatches as CnM

import fire

TIMEOUT_SLEEP = 5

VolumeInformation = namedtuple("VolumeInformation", "Label info1 info2 info3 type")


def get_volume_label(*args, **kwargs):
    """
    Gets volume information using Windows API.
    :param args: Arguments for GetVolumeInformation from Windows API
    :param kwargs: Keyword arguments for GetVolumeInformation from Windows API
    :return: label from VolumeInformation tuple instance
    """
    volume_information = VolumeInformation(*win32api.GetVolumeInformation(*args, **kwargs))
    return volume_information.Label


def test_drive(drive: str):
    """
    Tests if drive is ready using Windows API
    :param drive: Drive from which backup is meant to be performed
    :return: True if device is ready, False otherwise
    """
    try:
        win32file.GetDiskFreeSpaceEx(drive)
    except win32api.error as err:
        if 'The device is not ready.' not in err.args:
            logging.error(err)
            raise err
        return False
    return True


def save_to_iso(img_burn_exe: str, drive: str, output_file: str) -> None:
    """
    Uses ImgBurn exe CLI to perform backup. Automatically runs backup, starts when device is ready and closes when finished.
    :param img_burn_exe: path to Exe file
    :param drive: Drive from which backup will be performed
    :param output_file: Where ISO file will be saved
    """
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


def list_files_in_drive(drive_letter) -> List[Path]:
    """
    Lists files from drive
    :param drive_letter: Drive letter
    :return: List of file paths
    """
    return list(Path(drive_letter).rglob("*.*"))


# noinspection PyUnresolvedReferences
def backup_disk(autorun_label, drive, img_burn_exe, output_folder):
    """
    Performs backup
    :param autorun_label: Detected label from autorun.inf file
    :param img_burn_exe: path to Exe file
    :param drive: Drive from which backup will be performed
    :param output_folder: Folder to which output will be saved
    """
    label = get_volume_label(f"{drive}\\")
    files = list_files_in_drive(drive)
    logging.info(f"autorun label: {autorun_label}")
    logging.info(f"Standard label: {label}")
    text_for_files = "\n".join(str(file) for file in files)
    logging.debug(f"List of files: {text_for_files}")
    files_hash = hashlib.md5(text_for_files.encode('utf-8')).hexdigest()
    iso_folder = Path(output_folder) / f"{label}_{files_hash}"
    logging.info(f"Iso folder: {iso_folder}")
    iso_folder.mkdir(parents=True)
    if autorun_label:
        file_name = Path(f"{label}_{autorun_label}").with_suffix(".iso")
    else:
        file_name = Path(f"{label}").with_suffix(".iso")
    output_file = str(iso_folder / file_name)
    try:
        save_to_iso(img_burn_exe, drive, output_file=output_file)
    except ValueError:
        return
    (iso_folder / "list_of_files.txt").write_text(data=text_for_files, encoding='utf-8')


def process_drive(img_burn_exe: str, drive: str, output_folder: str):
    """
    Processes drive, tests if it is ready and if it is - tries to backup
    :param img_burn_exe: path to Exe file
    :param drive: Drive from which backup will be performed
    :param output_folder: Folder to which output will be saved
    :return:
    """
    if not test_drive(drive):
        logging.info("Waiting for drive: %s to be ready", drive)
        return
    autorun_file = Path(f"{drive}Autorun.inf")
    autorun_label = ""
    if autorun_file.is_file():
        parser = ConfigParser()
        encoding = CnM.from_path(autorun_file).best().first().encoding
        logging.debug("Detected autorun.inf encoding: %s", encoding)
        try:
            parser.read_string(autorun_file.read_text(encoding=encoding).lower())
        except DuplicateOptionError as err:
            pass
        else:
            if 'label' in parser['autorun']:
                autorun_label = parser['autorun']['label'].upper()

    backup_disk(autorun_label, drive, img_burn_exe, output_folder)


def poll_drive_for_backup(img_burn_exe: str, drive: str, output_folder: str) -> None:
    """
    Polls drive for processing
    :param img_burn_exe: path to Exe file
    :param drive: Drive from which backup will be performed
    :param output_folder: Folder to which output will be saved
    """
    while True:
        process_drive(img_burn_exe, drive, output_folder=output_folder)
        time.sleep(TIMEOUT_SLEEP)


if __name__ == '__main__':
    logging.basicConfig(level=logging.DEBUG, format="[%(asctime)s][%(levelname).1s]: %(message)s",
                        datefmt="%Y-%m-%d %H:%M:%S")
    fire.Fire(poll_drive_for_backup)
