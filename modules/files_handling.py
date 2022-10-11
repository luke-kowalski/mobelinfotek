from modules import settings, excel_data
from datetime import datetime
import os

time_now = datetime.now().strftime("%Y_%m_%d-%I_%M_%S_%p")


def get_excel_files():
    for file in os.listdir(settings.EXCEL_SOURCE_PATH):
        if file.endswith(".xlsx"):
            excel_data.delete_existed_entries_in_ms_sql(file)
            excel_data.insert_excel_data_to_ms_sql(file)
            move_excel_file_to_archive(file)


def move_excel_file_to_archive(file):
    try:
        os.replace(os.path.join(settings.EXCEL_SOURCE_PATH, file),
                   os.path.join(settings.EXCEL_ARCHIVE_PATH, time_now + "_" + file))
    except Exception as e:
        print(f"Can't move file to archive: {e}")
