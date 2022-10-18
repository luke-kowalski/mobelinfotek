from modules import settings, excel_data
from datetime import datetime
import os

time_now = datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p")


def validate_excel_file():
    files = get_excel_files_list()
    if len(files) > 0:
        for file in files:
            excel_data.check_excel_data(file)
    else:
        print("ERROR: There is no Excel files in the directory!")


def get_excel_files_list():
    files_list = []
    for file in os.listdir(settings.EXCEL_SOURCE_PATH):
        if file.endswith(".xlsx"):
            files_list.append(file)

    return files_list


def move_excel_file_to_archive_dir(file):
    try:
        os.replace(os.path.join(settings.EXCEL_SOURCE_PATH, file),
                   os.path.join(settings.EXCEL_ARCHIVE_PATH, time_now + "_" + file))
        print(f"OK: File: {file} has been moved to archive directory!")
    except Exception as e:
        print(f"ERROR: Can't move file to archive: {e}")
