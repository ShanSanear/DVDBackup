# DVDBackup

This simple script uses ImgBurn application to create backup copies from CD/DVD files to specified directory as ISO file.

Works only on Windows due to used libraries.

## Usage

`pip install -r requirements.txt`

`python dvd_backup.py path_to_img_burn_exe drive_letter: path_to_output_folder`

## How it works

Script polls drive and checks if it is ready. 
Then when it is, gets list of files.

From this list creates hash from it to generate unique folder name, containing also label of the disk. 
Those are used for ImgBurn CLI.

## Libraries used

Fire - for easy CLI processing.

Charset normalizer - to automatically detect `autorun.inf` file encoding

Pywin32 - for detecting if there is disk in drive