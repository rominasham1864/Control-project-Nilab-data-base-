import pandas as pd
import pymysql
import openpyxl

workbook = openpyxl.load_workbook('REM-40204-777-294(06).xlsm', keep_vba=True)

worksheet = workbook['REM-40204-777-294']

cell_value = worksheet['W4'].value

conn = pymysql.connect(host="localhost", user="root", password="1122", database="request_control")


cur = conn.cursor()

cur.execute("CREATE TABLE IF NOT EXISTS mytable (value VARCHAR(255))")

cur.execute("INSERT INTO mytable VALUES (%s)", (cell_value,))

conn.commit()

# Close the cursor and the connection to the MySQL database
cur.close()