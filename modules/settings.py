from configparser import ConfigParser

parser = ConfigParser()
parser.read("config.ini")

SQL_driver = parser.get("SERVER_CONN", "SQL_DRIVER")
SQL_server = parser.get("SERVER_CONN", "SQL_SERVER")
SQL_database = parser.get("SERVER_CONN", "SQL_DATABASE")
SQL_username = parser.get("SERVER_CONN", "SQL_USERNAME")
SQL_password = parser.get("SERVER_CONN", "SQL_PASSWORD")

EXCEL_SOURCE_PATH = parser.get("EXCEL_FILE", "EXCEL_SOURCE_PATH")
EXCEL_ARCHIVE_PATH = parser.get("EXCEL_FILE", "EXCEL_ARCHIVE_PATH")
EXCEL_START_ROW_NUMBER = parser.get("EXCEL_FILE", "EXCEL_START_ROW_NUMBER")
EXCEL_SHEET_NAME = parser.get("EXCEL_FILE", "EXCEL_SHEET_NAME")
