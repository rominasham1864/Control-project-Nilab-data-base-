import pandas as pd
import pymysql
import openpyxl
import os
from datetime import datetime
import tkinter as tk

conn = pymysql.connect(
    host="localhost", user="root", password="1122", database="request_control"
)

window = tk.Tk()
window.geometry("500x500")
project_name_label = tk.Label(window, text="Project Name:")
project_name_label.grid(row=0, column=0)

# Create the request number text box
project_name_entry = tk.Entry(window)
project_name_entry.grid(row=0, column=1)

# Create a search button for the request number
def save_project_name():
    global fileName
    fileName = project_name_entry.get()
    
project_name_button = tk.Button(window, text="Search", command=save_project_name)
project_name_button.grid(row=0, column=2)
window.mainloop()

workbook = openpyxl.load_workbook(
    fileName, keep_vba=True, data_only=True
)
cur = conn.cursor()


worksheet = workbook[fileName]

req_n = os.path.basename(
    "C:/Users/alire/Desktop/rominas workspace/"+fileName+".xlsm"
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