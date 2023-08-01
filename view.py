import os
import pymysql
import tkinter as tk
from tkinter import *
from tkinter.ttk import *
import tkinter as tk
from tkinter import ttk
import os
import sqlite3

conn = pymysql.connect(
    host="localhost", user="root", password="1122", database="request_control"
)
cursor = conn.cursor()


def notFound():
    newWindow = Toplevel(window)
    newWindow.title("FileNotFound")
    newWindow.geometry("300x100")
    title_label = tk.Label(
        newWindow,
        text="ندارد وجود درخواست ",
        fg="red",
        font="Verdana 10 bold",
    )
    title_label.place(x=100, y=10)

    def ok():
        newWindow.destroy()

    update_button = tk.Button(newWindow, text="ok", height=1, width=8, command=ok)
    update_button.place(x=120, y=60)


def checkForFile(req_n, table_name):
    if len(req_n) == 0:
        print("here")
        querynull = f"SELECT * FROM {table_name}"
        cursor.execute(querynull)
    else:
        print("there")
        query = f"SELECT * FROM {table_name} WHERE req_n = %s"
        value = (req_n,)
        cursor.execute(query, value)
    data = cursor.fetchall()
    if data:
        showTable(data)
    else:
        notFound()


def showTable(data): 
    h = Scrollbar(window, orient = 'horizontal')
    h.pack(side = BOTTOM, fill = X)
    table = ttk.Treeview(window, columns=('1', '2', '3', '4', '5', '6', '7', '8', '9', '10'), show='headings', height=8, xscrollcommand=h)
    table.pack()
    h.config(command=table.xview)
    table.column("1",anchor=CENTER, stretch=YES, width=10)
    table.heading('1', text='Id')
    table.column("2",anchor=CENTER, stretch=YES, width=120)
    table.heading('2', text='شماره درخواست')
    table.column("3",anchor=CENTER, stretch=YES, width=110)
    table.heading('3', text='محصول')
    table.column("4",anchor=CENTER, stretch=YES, width=100)
    table.heading('4', text='تعداد آیتم')
    table.column("5",anchor=CENTER, stretch=YES, width=100)
    table.heading('5', text='موجود در نبار')
    table.column("6",anchor=CENTER, stretch=YES, width=100)
    table.heading('6', text='درخواست کننده')
    table.column("7",anchor=CENTER, stretch=YES, width=100)
    table.heading('7', text='تاریخ درخواست')
    table.column("8",anchor=CENTER, stretch=YES, width=100)
    table.heading('8', text='محل مصرف')
    table.column("9",anchor=CENTER, stretch=YES, width=100)
    table.heading('9', text='واحد')
    table.column("10",anchor=CENTER, stretch=YES, width=100)
    table.heading('10', text='واحد')
    table.place(x=50, y=170)
    i=0
    for row in data:
        table.insert('', 'end',text=i,values=(i+1,data[i][2], data[i][0], data[i][1], data[i][6], data[i][9], data[i][3], data[i][4], data[i][7], data[i][8]))
        i+=1

    
def upload():
    window.destroy()
    os.system("python pythontest.py")


def create_upload_button(window):
    upload_button = tk.Button(window, text="Upload New File", command=upload)
    upload_button.place(x=50, y=100)


def chooseTable(name):
    return {
        "فاضلاب قم 5 ساله": "quem",
        "فاضلاب خین عرب": "khin",
        "فاضلاب التیمور": "alteymour",
        "آبرسانی جاسک": "jask",
        "رودان 2": "roudan",
        "رشت": "rasht",
        "مرکزی": "markazi",
    }[name]


# Create the main window
window = tk.Tk()
window.geometry("1200x500")
window.title("سیستم کنرل درخواست کالا و کار")
#####################################################
request_number_label = tk.Label(window, text="Request Number:")
request_number_label.place(x=50, y=20)
# Create the request number text box
request_number_entry = tk.Entry(window)
request_number_entry.place(x=160, y=20)

#####################################################
# Create a label for the project name choice box
project_name_label = tk.Label(window, text="Project Name:")
project_name_label.place(x=51, y=60)

# Create the project name choice box
project_name_var = tk.StringVar()
options = ttk.Combobox(window, width=17, textvariable=project_name_var)
options.place(x=160, y=60)
options["values"] = (
    "مرکزی",
    "رشت",
    "رودان 2",
    "آبرسانی جاسک",
    "فاضلاب التیمور",
    "فاضلاب خین عرب",
    "فاضلاب قم 5 ساله",
)


def search_project():
    selected_value = project_name_var.get()
    checkForFile(request_number_entry.get(), chooseTable(selected_value))
project_name_button = tk.Button(window, text="Search", command=search_project)
project_name_button.place(x=350, y=45)

###################################################
create_upload_button(window)
image_label = tk.Label(window)
image_file = tk.PhotoImage(file="C:/Users/alire/Desktop/rominas workspace/logo5.png")
resized_image_file = image_file.subsample(2, 2)
# Set the image for the label
image_label.config(image=resized_image_file)
image_label.place(x=1050, y=10)

# Run the GUI
window.mainloop()
