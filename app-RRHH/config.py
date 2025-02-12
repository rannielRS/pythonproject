import pymssql
import psycopg2

COLOR_BARRA_SUPERIOR = "#d52323"
COLOR_MENU_LATERAL = "#d52323"
COLOR_MENU_LATERAL_UP = "#f1faff"
COLOR_CUERPO_PRINCIPAL = "#f1faff"
COLOR_MENU_CURSOR_ENCIMA = "#d52323"
CONN_ZUN = pymssql.connect(
            server = '10.105.213.6',
            user='userutil',
            password = '1234',
            database='ZUNpr',
            as_dict = True
            )
CURSOR_ZUN = CONN_ZUN.cursor()
CONN_LOC= psycopg2.connect(
            host="localhost",
            database="postgres",
            user="postgres",
            password="proyecto")
CURSOR_LOC = CONN_LOC.cursor()