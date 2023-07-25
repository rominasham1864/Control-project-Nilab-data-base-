import pandas as pd
import pymysql
import openpyxl
import os
import tkinter as tk

def connect_to_database():
    """Establishes a connection to the MySQL server and returns a cursor object."""
    conn = pymysql.connect(
        host="localhost", user="root", password="1122", database="request_control"
    )
    return conn.cursor(), conn

def save_project_name():
    """Saves the project name from the project_name_entry widget to the global fileName variable."""
    global fileName
    fileName = project_name_entry.get()


def save_data_to_database(file_name):
    """Saves data from the Excel file to the MySQL database."""
    workbook = openpyxl.load_workbook(
        file_name+".xlsm", keep_vba=True, data_only=True
    )
    cursor, conn = connect_to_database()

    worksheet = workbook[file_name]

    req_n =file_name

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
            cursor.execute(sql, val)

    conn.commit()
    cursor.close()
    conn.close()

# Create the main window
window = tk.Tk()
window.geometry("500x500")
window.title("سیستم کنرل درخواست کالا و کار")

# Create the project name label and entry widgets
request_number_label = tk.Label(window, text="Riquest Number:")
request_number_label.place(x=10, y=80)

request_number_entry = tk.Entry(window)
request_number_entry.place(x=105, y=80)

# Create a search button for the project name
request_number_button = tk.Button(window, text="Submit", command=lambda: save_project_name())
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

# Save the data to the database
save_data_to_database(fileName)
print(fileName)