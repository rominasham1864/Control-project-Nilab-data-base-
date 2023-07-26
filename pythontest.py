import pandas as pd
import pymysql
import openpyxl
import os
import tkinter as tk
from tkinter import *
from tkinter.ttk import *


def connect_to_database():
    """Establishes a connection to the MySQL server and returns a cursor object."""
    conn = pymysql.connect(
        host="localhost", user="root", password="1122", database="request_control"
    )
    return conn.cursor(), conn


def errorWindow(error_massage):
    newWindow = Toplevel(window)
    newWindow.title("Error")
    newWindow.geometry("400x100")
    title_label = tk.Label(
        newWindow, text=error_massage, fg="red", font="Verdana 10 bold"
    )
    title_label.place(x=50, y=10)
    close_button = tk.Button(
        newWindow, text="ok", height=1, width=8, command=newWindow.destroy
    )
    close_button.place(x=160, y=60)


def run_program():
    window.destroy()
    os.system("python test.py")


def choice(code):
    return {
        "770": "khin",
        "667": "jask",
        "666": "roudan",
        "110": "markazi",
        "210": "rasht",
        "777": "quem",
        "880": "alteymour",
    }[code]


def save_data_to_database():
    try:
        global file_name
        file_name = request_number_entry.get()
        worksheet = worksheet[file_name]
        req_n = file_name
        o_date = worksheet.cell(row=6, column=14).value
        ref_date = worksheet.cell(row=17, column=14).value
        code = worksheet.cell(row=6, column=8).value
        table = choice(str(code))
        for row in range(8, 14):
            prod = worksheet.cell(row=row, column=3).value
            qty = worksheet.cell(row=row, column=9).value
            available = worksheet.cell(row=row, column=11).value
            print(qty)
            if prod != None:
                sql = f"INSERT INTO {table} (product, req_n, Qty, o_date, available_in_stock, ref_date) VALUES (%s, %s, %s,%s,%s, %s)"
                val = (prod, req_n, qty, o_date, available, ref_date)
                cursor.execute(sql, val)
        done_label = tk.Label(
            window, text="شد ثبت موفقیت با اطلاعات", fg="green", font="Verdana 10 bold"
        )
        done_label.place(x=200, y=200)
        ok_button = tk.Button(window, text="ok", height=1, width=8, command=run_program)
        ok_button.place(x=250, y=250)
        conn.commit()
        cursor.close()
        conn.close()
    except FileNotFoundError as e:
        errorWindow(
            "ندارد وجود پوشه در نظر مورد فایل\nباشد REM-####-###-### صورت به باید نام فرمت"
        )
    except KeyError as e:
        errorWindow(
            "است قبول قابل غیر فایل نام\n آورید در REM-####-###-### صورت به را آن لطفا"
        )


# Create the main window
window = tk.Tk()
window.geometry("500x500")
window.title("سیستم کنرل درخواست کالا و کار")
title_label = tk.Label(window, text="شرکت نیلآب صنعت البرز")
title_label.place(x=190, y=10)

# Create the project name label and entry widgets
request_number_label = tk.Label(window, text="Riquest Number:")
request_number_label.place(x=10, y=80)

request_number_entry = tk.Entry(window)
request_number_entry.place(x=105, y=80)

# Create a search button for the project name
request_number_button = tk.Button(
    window, text="Submite", command=lambda: save_data_to_database()
)
request_number_button.place(x=250, y=75)

# Display the company logo image
image_label = tk.Label(window)
image_file = tk.PhotoImage(file="C:/Users/alire/Desktop/rominas workspace/logo5.png")
resized_image_file = image_file.subsample(3, 3)

# Set the image for the label
image_label.config(image=resized_image_file)
image_label.place(x=400, y=10)
# Run the GUI
window.mainloop()
