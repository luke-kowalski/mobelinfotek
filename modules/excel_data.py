from openpyxl import load_workbook
from modules import ms_sql_database
from modules import settings
import os


def get_excel_sheet_data(file_name):
    wb = load_workbook(filename=os.path.join(settings.EXCEL_SOURCE_PATH, file_name))
    sheet = wb[settings.EXCEL_SHEET_NAME]
    return sheet


def delete_existed_entries_in_ms_sql(file_name):
    sheet = get_excel_sheet_data(file_name)
    orders_name_list = set()
    for row in sheet.iter_rows(min_row=int(settings.EXCEL_START_ROW_NUMBER), values_only=True):
        imos_order_name = row[3]
        orders_name_list.add(imos_order_name)

    for order_name in orders_name_list:
        ms_sql_database.ms_sql_delete_orders(order_name)
        # print(f"Data entry: {order_name} has been deleted!")


def insert_excel_data_to_ms_sql(file_name):
    sheet = get_excel_sheet_data(file_name)
    for row in sheet.iter_rows(min_row=int(settings.EXCEL_START_ROW_NUMBER), values_only=True):
        building = row[0]
        building_floor = row[1]
        building_unit = row[2]
        imos_order_name = row[3]
        furniture_type = row[4]
        ms_sql_database.ms_sql_insert_data_to_database(building, building_floor, building_unit, imos_order_name, furniture_type)
        print(f"Data entry: {building}_{building_floor}_{building_unit}_{imos_order_name}_{furniture_type} has been added!\n")
