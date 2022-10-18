import pyodbc
from modules import settings


def ms_sql_conn():
    try:
        conn = pyodbc.connect(
            "DRIVER="
            + settings.SQL_driver
            + ";SERVER="
            + settings.SQL_server
            + ";DATABASE="
            + settings.SQL_database
            + ";UID="
            + settings.SQL_username
            + ";PWD="
            + settings.SQL_password
        )
        return conn
    except pyodbc.Error as e:
        sqlstate = e.args[1]
        print(f"ERROR: {sqlstate}")
        raise


def ms_sql_delete_orders(name):
    try:
        conn = ms_sql_conn()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM MOBELINFOTEK_FURNITURES WHERE IMOS_ORDER_NAME = ?;", name)
        conn.commit()

    except pyodbc.Error as e:
        sqlstate = e.args[1]
        print(f"ERROR: {sqlstate}")
        raise

    finally:
        conn.close()


def ms_sql_insert_data_to_database(BUILDING, BUILDING_FLOOR, BUILDING_UNIT, IMOS_ORDER_NAME, FURNITURE_TYP):
    try:
        conn = ms_sql_conn()
        cursor = conn.cursor()
        cursor.execute("INSERT INTO MOBELINFOTEK_FURNITURES (BUILDING, BUILDING_FLOOR, BUILDING_UNIT, IMOS_ORDER_NAME, FURNITURE_TYP, FLAG, CREATION_DATE) VALUES (?,?,?,?,?,1,GETUTCDATE());", BUILDING, BUILDING_FLOOR, BUILDING_UNIT, IMOS_ORDER_NAME, FURNITURE_TYP)
        conn.commit()

    except pyodbc.Error as e:
        sqlstate = e.args[1]
        print(f"ERROR: {sqlstate}")
        raise

    finally:
        conn.close()
