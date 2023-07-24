import pandas as pd
import pymysql
import openpyxl
import os

workbook = openpyxl.load_workbook(
    "REM-40204-777-294(06).xlsm", keep_vba=True, data_only=True
)
conn = pymysql.connect(
    host="localhost", user="root", password="1122", database="request_control"
)
cur = conn.cursor()


worksheet = workbook["REM-40204-777-294"]

req_n = os.path.basename(
    "C:/Users/alire/Desktop/rominas workspace/REM-40204-777-294(06).xlsm"
)

o_date = worksheet.cell(row=6, column=14).value
ref_date = worksheet.cell(row=17, column=14).value

for row in range(8, 14):
    prod = worksheet.cell(row=row, column=3).value
    qty = worksheet.cell(row=row, column=9).value
    available = worksheet.cell(row=row, column=11).value
    print(qty)
    if prod != None:
        sql = "INSERT INTO alteymour (product, req_n, Qty, o_date, available_in_stock, ref_date) VALUES (%s, %s, %s,%s,%s, %s)"
        val = (prod, req_n, qty, o_date, available, ref_date)
        cur.execute(sql, val)
        
           

conn.commit()
cur.close()