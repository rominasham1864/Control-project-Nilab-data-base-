import openpyxl
import pymysql
from tkinter import *
from tkinter.ttk import *
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from openpyxl.styles import PatternFill
import pyperclip
import os
from tkinter.filedialog import askopenfilename
import customtkinter as ctk
from PIL import Image, ImageTk

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

conn = pymysql.connect(
    host="localhost", user="root", password="1122", database="request_control"
)
cursor = conn.cursor()


def upload():
    def errorWindow(error_massage):
        newWindow = ctk.CTk()
        newWindow.title("Error")
        newWindow.geometry("400x100")
        title_label = ctk.CTkLabel(newWindow, text=error_massage)
        title_label.place(x=80, y=10)
        close_button = ctk.CTkButton(newWindow, text="ok", command=newWindow.destroy)
        close_button.place(x=110, y=60)
        newWindow.mainloop()

    def run_program():
        window.destroy()
        View()

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

    def chooseCode(name):
        return {
            "فاضلاب قم 5 ساله": 777,
            "فاضلاب خین عرب": 770,
            "فاضلاب التیمور": 880,
            "آبرسانی جاسک": 667,
            "رودان 2": 666,
            "رشت": 210,
            "مرکزی": 110,
        }[name]

    def checkForFile(fileLocation):
        file_name = os.path.basename(fileLocation)
        file_name = file_name.replace(".xlsm", "")
        query = "SELECT * FROM main WHERE req_n = %s"
        value = (file_name,)
        cursor.execute(query, value)
        if cursor.fetchone():
            updateWindow(file_name, fileLocation)
        else:
            save_data_to_database(file_name, False, fileLocation)

    def updateWindow(file_name, file_path):
        newWindow = ctk.CTk()
        newWindow.title("Error")
        newWindow.geometry("400x100")
        title_label = ctk.CTkLabel(
            newWindow,
            text="کنید؟ آپدیت را آن میخواهید آیا. دارد وجود دیتابیس در فایل این",
        )
        title_label.place(x=60, y=10)

        def update():
            newWindow.destroy()
            save_data_to_database(file_name, True, file_path)

        update_button = ctk.CTkButton(newWindow, text="update", command=update)
        update_button.place(x=220, y=60)

        def close():
            newWindow.destroy()

        close_button = ctk.CTkButton(newWindow, text="cancle", command=close)
        close_button.place(x=50, y=60)
        newWindow.mainloop()

    def save_data_to_database(file_name, delete_needed, file_path):
        try:
            workbook = openpyxl.load_workbook(file_path, keep_vba=True, data_only=True)

            worksheet = workbook[file_name]
            # file name
            req_n = file_name
            # request type
            if "REM" in req_n:
                req_t = "کالا"
            else:
                req_t = "کار"
            # order date
            o_date = worksheet["N6"].value
            # referral date
            if req_t == "کالا":
                ref_date = worksheet["N17"].value
            else:
                ref_date = worksheet["N16"].value
            # request status
            if worksheet["X8"].value == True:
                status = "توقف"
            else:
                if worksheet["X13"].value == TRUE:
                    if req_t == "کالا":
                        status = "تکمیل و ارسال به سایت"
                    else:
                        status = "تکمیل کار"
                elif worksheet["X12"].value == TRUE:
                    if req_t == "کالا":
                        status = "پروسه ی خرید"
                    else:
                        status = "پروسه ی ارجاع کار به پیمانکار"

                elif worksheet["X11"].value == TRUE:
                    status = "تأمین بودجه"
                elif worksheet["X10"].value == TRUE:
                    if req_t == "کالا":
                        status = "جستجوی کالا "
                    else:
                        status = "جستجوی پیمانکار"
                else:
                    status = "درخواست اطلاعات از سایت"
            # project name
            project_name = worksheet["G6"].value

            # get the project name and code
            table = chooseTable(project_name)
            code = chooseCode(project_name)

            # check if it needs updates
            if delete_needed:
                query = f"DELETE FROM {table} WHERE req_n=%s"
                value = (file_name,)
                query_main = f"DELETE FROM main WHERE req_n=%s"
                value_main = (file_name,)
                cursor.execute(query, value)
                cursor.execute(query_main, value_main)
            # applicant
            Applicant = worksheet["D6"].value
            # get the materials value
            for row in range(8, 14):
                if req_t == "کالا":
                    prod = worksheet.cell(row=row, column=3).value
                    qty = worksheet.cell(row=row, column=9).value
                    available = worksheet.cell(row=row, column=11).value
                    place_of_usage = worksheet.cell(row=row, column=14).value
                    unit = worksheet.cell(row=row, column=13).value
                else:
                    prod = worksheet.cell(row=row, column=3).value
                    place_of_usage = worksheet["K7"].value
                    qty = None
                    available = None
                    unit = None
                if prod != None and prod != "شرح خدمات درخواستی :	\n":
                    sql = f"INSERT INTO {table} (product, req_n, Qty, o_date, available_in_stock, ref_date, place_of_usage, unit ,Applicant, st_request, req_t) VALUES (%s, %s, %s,%s,%s, %s, %s, %s, %s, %s, %s)"
                    val = (
                        prod,
                        req_n,
                        qty,
                        o_date,
                        available,
                        ref_date,
                        place_of_usage,
                        unit,
                        Applicant,
                        status,
                        req_t,
                    )
                    cursor.execute(sql, val)
            sql_main = f"INSERT INTO main (project, req_n, o_date, f_date, p_code, st_request, req_t) VALUES (%s ,%s, %s, %s, %s, %s, %s)"
            val_main = (project_name, req_n, o_date, ref_date, code, status, req_t)
            cursor.execute(sql_main, val_main)
            conn.commit()
            done_wondow()
            

        except FileNotFoundError as e:
            errorWindow(
                "ندارد وجود پوشه در نظر مورد فایل\nباشد RE#-#####-###-### صورت به باید نام فرمت"
            )
        except KeyError as e:
            errorWindow(
                "است قبول قابل غیر فایل نام\n است RE#-#####-###-### قبول قابل فرمت"
            )

    def done_wondow():
        new_window = ctk.CTk()
        new_window.title("saved")
        new_window.geometry("300x100")
        done_label = ctk.CTkLabel(
            new_window,
            text="شد ثبت موفقیت با اطلاعات",
        )
        done_label.place(x=100, y=15)
        ok_button = ctk.CTkButton(new_window, text="ok", command=new_window.destroy)
        ok_button.place(x=80, y=60)
        new_window.mainloop()

    def codeTableDisplay():
        table = ttk.Treeview(window, columns=("1", "2"), show="headings", height=6)
        table.pack()
        table.column("1", anchor=CENTER, stretch=YES, width=100)
        table.heading("1", text="نام پروژه")
        table.column("2", anchor=CENTER, stretch=YES, width=50)
        table.heading("2", text="کد پروژه")
        table.insert("", "end", values=("فاضلاب قم 5 ساله", "777"))
        table.insert("", "end", values=("فاضلاب خین عرب", "770"))
        table.insert("", "end", values=("فاضلاب التیمور", "880"))
        table.insert("", "end", values=("آبرسانی جاسک", "667"))
        table.insert("", "end", values=("رودان 2", "666"))
        table.insert("", "end", values=("رشت", "210"))
        table.insert("", "end", values=("مرکزی", "110"))
        table.place(x=300, y=100)

    def askForFile():
        checkForFile(askopenfilename())

    # Create the main window
    window = ctk.CTk()
    window.geometry("500x500")
    window.title("سیستم کنرل درخواست کالا و کار")
    title_label = ctk.CTkLabel(window, text="البرز صنعت نیلآب شرکت")
    title_label.place(x=200, y=10)

    # Create the project name label and entry widgets
    request_number_label = ctk.CTkLabel(window, text="Select File : ")
    request_number_label.place(x=10, y=80)
    find_button = ctk.CTkButton(window, text="  bowers   ", command=askForFile)
    find_button.place(x=90, y=78)

    back_button = ctk.CTkButton(
        window, text="   back   ", command=lambda: run_program()
    )
    back_button.place(x=10, y=125)
    # Display the company logo image
    image = Image.open("C:/Users/alire/Desktop/rominas workspace/logo5.png")
    resized_image = image.resize((100, 80))
    photo = ImageTk.PhotoImage(resized_image)
    canvas = Canvas(
        window, bg="#282424", width=resized_image.width, height=resized_image.height
    )
    canvas.pack(side="top", anchor="ne")
    canvas.create_image(0, 0, anchor=NW, image=photo)

    codeTableDisplay()
    # Run the GUI

    window.mainloop()
    cursor.close()
    conn.close()


def View():
    def reverse_name(name):
        return {
            "ساله 5 قم فاضلاب": "فاضلاب قم 5 ساله",
            "عرب خین فاضلاب": "فاضلاب خین عرب",
            "التیمور فاضلاب": "فاضلاب التیمور",
            "جاسک آبرسانی": "آبرسانی جاسک",
            "2 رودان": "رودان 2",
            "رشت": "رشت",
            "مرکزی": "مرکزی",
        }[name]

    def errorwindow(string, n):
        new_window = ctk.CTk()
        new_window.title("Error")
        new_window.geometry("300x100")
        title_label = ctk.CTkLabel(
            new_window,
            text=string,
        )
        title_label.pack()

        if n == 1:
            title_label.place(x=100, y=10)
        elif n == 2:
            title_label.place(x=70, y=10)
        else:
            title_label.place(x=70, y=10)

        update_button = ctk.CTkButton(new_window, text="ok", command=new_window.destroy)
        update_button.place(x=80, y=60)
        new_window.mainloop()

    def checkForFile(search, table_name):
        if len(search) == 0:
            querynull = f"SELECT * FROM {table_name}"
            cursor.execute(querynull)
        else:
            query = f"SELECT * FROM {table_name} WHERE product LIKE '%{search}%' OR req_n LIKE '%{search}%' "
            cursor.execute(query)
        data = cursor.fetchall()
        if data:

            def print():
                printData(list(data), table_name)

            back_button = ctk.CTkButton(window, text="   print   ", command=print)
            back_button.place(x=1005, y=175)
            if table_name == "main":
                mainTable(data)
                mainFiltering(data)

            else:
                showTable(data, table_name)
        else:
            errorwindow("ندارد وجود درخواست", 1)

    def mainFiltering(data):
        project_name = ctk.StringVar()
        options = ctk.CTkOptionMenu(
            window,
            variable=project_name,
            values=(
                "مرکزی",
                "رشت",
                "2 رودان",
                "جاسک آبرسانی",
                "التیمور فاضلاب",
                "عرب خین فاضلاب",
                "ساله 5 قم فاضلاب",
            ),
        )
        project_name.set("مرکزی")
        options.place(x=520, y=60)

        def filtering():
            table_name = reverse_name(project_name.get())
            query = f"SELECT * FROM main WHERE project = %s"
            value = (table_name,)
            cursor.execute(query, value)
            data = cursor.fetchall()
            mainTable(data)

        button = ctk.CTkButton(window, text="Search in main", command=filtering)
        button.place(x=690, y=60)
        return data

    def printData(data, table_name):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        workbook = openpyxl.Workbook()
        worksheet = workbook.active
        if table_name == "main":
            worksheet.cell(row=1, column=1).value = "نام پروژه"
            worksheet.cell(row=1, column=2).value = "کد پروژه"
            worksheet.cell(row=1, column=3).value = "شماره درخواست"
            worksheet.cell(row=1, column=4).value = "تاریخ درخواست"
            worksheet.cell(row=1, column=5).value = "تاریخ ارجا"
            worksheet.cell(row=1, column=6).value = "نوع درخواست"
            worksheet.cell(row=1, column=7).value = "وضعیت درخواست"
            worksheet.cell(row=1, column=8).value = "پرداخت اول"
            worksheet.cell(row=1, column=9).value = "پرداخت دوم"
        else:
            worksheet.cell(row=1, column=1).value = "درخواست"
            worksheet.cell(row=1, column=2).value = "شماره درخواست"
            worksheet.cell(row=1, column=3).value = "تاریخ درخواست"
            worksheet.cell(row=1, column=4).value = "تاریخ ارجا"
            worksheet.cell(row=1, column=5).value = "تعداد"
            worksheet.cell(row=1, column=6).value = "موجود"
            worksheet.cell(row=1, column=7).value = "واحد"
            worksheet.cell(row=1, column=8).value = "محل مصرف"
            worksheet.cell(row=1, column=11).value = "نوع درخواست"
            worksheet.cell(row=1, column=9).value = "درخواست کننده"
            worksheet.cell(row=1, column=10).value = "وضعیت درخواست"

        for row_idx, row_data in enumerate(data, start=2):
            for col_idx, cell_value in enumerate(row_data, start=1):
                worksheet.cell(row=row_idx, column=col_idx).value = cell_value
                if cell_value == "توقف":
                    # Apply red fill to the entire row
                    fill = PatternFill(
                        start_color="FF0000", end_color="FF0000", fill_type="solid"
                    )
                    for cell in worksheet[row_idx]:
                        cell.fill = fill
                elif cell_value == "تکمیل و ارسال به سایت":
                    # Apply red fill to the entire row
                    fill = PatternFill(
                        start_color="00FF50", end_color="00FF50", fill_type="solid"
                    )
                    for cell in worksheet[row_idx]:
                        cell.fill = fill
        for column in range(ord("A"), ord("L")):
            column_letter = chr(column)
            worksheet.column_dimensions[column_letter].width = 18

        workbook.save(file_path)

        newWindow = ctk.CTk()

        newWindow.geometry("300x100")
        title_label = ctk.CTkLabel(
            newWindow,
            text="شد ثبت موفقیت با اطلاعات",
        )
        title_label.place(x=100, y=10)
        close_button = ctk.CTkButton(newWindow, text="ok", command=newWindow.destroy)
        close_button.place(x=80, y=60)
        newWindow.mainloop()

    def copy_to_clipboard(event, table):
        selected_row = table.focus()
        data = table.item(selected_row)["values"]
        pyperclip.copy(data[2])

    def handle_double_click(event, table):
        selected_row = table.focus()
        data = table.item(selected_row)["values"]
        req_n = data[3]
        
        newWindow1 = ctk.CTk()
        newWindow1.geometry("400x100")
        title_label = ctk.CTkLabel(
            newWindow1,
            text=req_n,
        )
        title_label.place(x=150, y=10)

        def comment():
            newWindow1.destroy()
            newWindow = ctk.CTk()
            newWindow.geometry("400x130")
            title_label = ctk.CTkLabel(
                newWindow,
                text=" سفارش  توضیحات ",
            )
            title_label.place(x=165, y=5)
            comment_entry = ctk.CTkEntry(
                newWindow, placeholder_text="comment", width=300
            )
            comment_entry.place(x=50, y=40)

            def insertButton():
                sql = f"Update main set comment = %s where req_n = %s"
                val = (comment_entry.get(), req_n)
                cursor.execute(sql, val)
                conn.commit()
                newWindow.destroy()
            insert_button = ctk.CTkButton(newWindow, text="ثبت", command=insertButton)
            insert_button.place(x=140, y=85)
            newWindow.mainloop()

        comment_button = ctk.CTkButton(
            newWindow1, text="توضیحات کردن اضافه", command=comment
        )
        comment_button.place(x=40, y=40)

        def pay_window():
            def insert():
                newWindow.destroy()
                try:
                    save = True
                    workbook = openpyxl.load_workbook(
                        "C:/Users/alire/Desktop/rominas workspace/payment order.xlsx"
                    )
                    sheet = workbook["Sheet1"]
                    amount = amount_entry.get()
                    sheet["F6"] = req_n
                    sheet["E10"] = receiver_entry.get()
                    sheet["L3"] = number_entry.get()
                    sheet["K9"] = amount
                    if data[8] == "None":
                        sql_main = f"Update main set payment1 = %s where req_n = %s"
                        val_main = (amount, req_n)
                        cursor.execute(sql_main, val_main)
                        conn.commit()
                    elif data[9] == "None":
                        sql_main = f"Update main set payment2 = %s where req_n = %s"
                        val_main = (amount, req_n)
                        cursor.execute(sql_main, val_main)
                        conn.commit()
                    else:
                        errorwindow("است شده پر درخواست 2 تعداد", 3)
                        save = False
                    if save == True:
                        workbook.save(
                            filedialog.asksaveasfilename(
                                defaultextension=".xlsx", initialfile=req_n + "-PO"
                            )
                        )
                    # newWindow1.destroy()
                except PermissionError as e:
                    errorwindow("نمیشود داده باز فایل برای دسترسی اجازه", 2)

            newWindow1.destroy()
            newWindow = ctk.CTk()
            newWindow.geometry("400x200")
            title_label = ctk.CTkLabel(
                newWindow,
                text=req_n + " پرداخت دستور صدور ",
            )
            title_label.place(x=80, y=10)
            receiver_label = ctk.CTkLabel(newWindow, text=": کننده دریافت")
            receiver_label.place(x=290, y=40)
            receiver_entry = ctk.CTkEntry(newWindow, placeholder_text="اداری")
            receiver_entry.place(x=110, y=45)

            amount_label = ctk.CTkLabel(newWindow, text=": مبلغ")
            amount_label.place(x=320, y=70)
            amount_entry = ctk.CTkEntry(newWindow, placeholder_text="2000000")
            amount_entry.place(x=110, y=75)

            number_entry = ctk.CTkEntry(newWindow, placeholder_text="103")
            number_entry.place(x=110, y=105)
            number_label = ctk.CTkLabel(newWindow, text=": شماره")
            number_label.place(x=320, y=100)

            insert_button = ctk.CTkButton(newWindow, text="insert", command=insert)
            insert_button.place(x=210, y=140)

            close_button = ctk.CTkButton(
                newWindow, text="cancle", command=newWindow.destroy
            )
            close_button.place(x=40, y=140)
            newWindow.mainloop()

        payment_button = ctk.CTkButton(
            newWindow1, text="پرداخت دستور صدور", command=pay_window
        )
        payment_button.place(x=210, y=40)
        newWindow1.mainloop()

    def mainTable(data):
        table = ttk.Treeview(
            window,
            columns=("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"),
            show="headings",
            height=14,
        )
        table.pack()
        table.column("1", anchor=CENTER, stretch=YES, width=40)
        table.heading("1", text="Id")
        table.column("2", anchor=CENTER, stretch=YES, width=105)
        table.heading("2", text="نام پروژه")
        table.column("3", anchor=CENTER, stretch=YES, width=50)
        table.heading("3", text="کد پروژه")
        table.column("4", anchor=CENTER, stretch=YES, width=110)
        table.heading("4", text="شماره درخواست")
        table.column("5", anchor=CENTER, stretch=YES, width=90)
        table.heading("5", text="تاریخ درخواست")
        table.column("6", anchor=CENTER, stretch=YES, width=75)
        table.heading("6", text="تاریخ ارجا")
        table.column("7", anchor=CENTER, stretch=YES, width=75)
        table.heading("7", text="نوع درخواست")
        table.column("8", anchor=CENTER, stretch=YES, width=140)
        table.heading("8", text="وضعیت درخواست")
        table.column("9", anchor=CENTER, stretch=YES, width=90)
        table.heading("9", text="پرداخت 1")
        table.column("10", anchor=CENTER, stretch=YES, width=90)
        table.heading("10", text="پرداخت 2")
        table.column("11", anchor=CENTER, stretch=YES, width=100)
        table.heading("11", text="توضیحات")
        table.place(x=20, y=170)
        i = 0
        for row in data:
            table.insert(
                "",
                "end",
                text=i,
                values=(
                    i + 1,
                    data[i][0],
                    data[i][1],
                    data[i][2],
                    data[i][3],
                    data[i][4],
                    data[i][5],
                    data[i][6],
                    data[i][7],
                    data[i][8],
                    data[i][9]
                ),
            )
            i += 1
        table.bind("<Double-1>", lambda event: handle_double_click(event, table))
    def showTable(data, table_name):
        table = ttk.Treeview(
            window,
            columns=("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13"),
            show="headings",
            height=14,
        )
        table.pack()
        table.column("1", anchor=CENTER, stretch=YES, width=10)
        table.heading("1", text="Id")
        table.column("2", anchor=CENTER, stretch=YES, width=120)
        table.heading("2", text="درخواست")
        table.column("3", anchor=CENTER, stretch=YES, width=115)
        table.heading("3", text="شماره درخواست")
        table.column("4", anchor=CENTER, stretch=YES, width=85)
        table.heading("4", text="تاریخ درخواست")
        table.column("5", anchor=CENTER, stretch=YES, width=65)
        table.heading("5", text="تاریخ ارجا")
        table.column("10", anchor=CENTER, stretch=YES, width=80)
        table.heading("10", text="درخواست کننده")
        table.column("6", anchor=CENTER, stretch=YES, width=35)
        table.heading("6", text="تعداد")
        table.column("7", anchor=CENTER, stretch=YES, width=35)
        table.heading("7", text="موجود")
        table.column("9", anchor=CENTER, stretch=YES, width=100)
        table.heading("9", text="محل مصرف")
        table.column("8", anchor=CENTER, stretch=YES, width=50)
        table.heading("8", text="واحد")
        table.column("11", anchor=CENTER, stretch=YES, width=120)
        table.heading("11", text="وضعیت درخواست")
        table.column("12", anchor=CENTER, stretch=YES, width=70)
        table.heading("12", text="نوع درخواست")
        table.column("13", anchor=CENTER, stretch=YES, width=80)
        table.heading("13", text="توضیحات")
        table.place(x=20, y=170)
        i = 0
        for row in data:
            table.insert(
                "",
                "end",
                text=i,
                values=(
                    i + 1,
                    data[i][0],
                    data[i][1],
                    data[i][2],
                    data[i][3],
                    data[i][4],
                    data[i][5],
                    data[i][6],
                    data[i][7],
                    data[i][8],
                    data[i][9],
                    data[i][10],
                    data[i][11]
                ),
                tag=("odd"),
            )
            i += 1
        table.bind("<c>", lambda event: copy_to_clipboard(event, table))
        def comment():
            selected_row = table.focus()
            row = table.item(selected_row)["values"]
            product = row[1]
            newWindow = ctk.CTk()
            newWindow.geometry("400x130")
            title_label = ctk.CTkLabel(
                newWindow,
                text=" سفارش  توضیحات ",
            )
            title_label.place(x=165, y=5)
            comment_entry = ctk.CTkEntry(
                newWindow, placeholder_text="comment", width=300
            )
            comment_entry.place(x=50, y=40)

            def insertButton():
                print(product)
                sql = f"Update {table_name} set comment = %s where product = %s"
                val = (comment_entry.get(), product)
                cursor.execute(sql, val)
                conn.commit()
                newWindow.destroy()
            insert_button = ctk.CTkButton(newWindow, text="ثبت", command=insertButton)
            insert_button.place(x=140, y=85)
            newWindow.mainloop()
        table.bind("<Double-1>", lambda event: comment())

    def chooseTable(name):
        return {
            "ساله 5 قم فاضلاب": "quem",
            "عرب خین فاضلاب": "khin",
            "التیمور فاضلاب": "alteymour",
            "جاسک آبرسانی": "jask",
            "2 رودان": "roudan",
            "رشت": "rasht",
            "مرکزی": "markazi",
            "main": "main",
        }[name]

    def discriptions():
        table = ttk.Treeview(window, columns=("1", "2"), show="headings", height=6)
        table.pack()
        table.column("1", anchor=CENTER, stretch=YES, width=100)
        table.heading("1", text="نام پروژه")
        table.column("2", anchor=CENTER, stretch=YES, width=50)
        table.heading("2", text="کد پروژه")
        table.insert("", "end", values=("فاضلاب قم 5 ساله", "777"))
        table.insert("", "end", values=("فاضلاب خین عرب", "770"))
        table.insert("", "end", values=("فاضلاب التیمور", "880"))
        table.insert("", "end", values=("آبرسانی جاسک", "667"))
        table.insert("", "end", values=("رودان 2", "666"))
        table.insert("", "end", values=("رشت", "210"))
        table.insert("", "end", values=("مرکزی", "110"))
        table.place(x=865, y=0)

    # Create the main window
    window = ctk.CTk()
    window.geometry("1200x500")
    image = Image.open("C:/Users/alire/Desktop/rominas workspace/logo5.png")
    resized_image = image.resize((140, 120))
    photo = ImageTk.PhotoImage(resized_image)
    canvas = Canvas(
        window, bg="#282424", width=resized_image.width, height=resized_image.height
    )
    canvas.pack(side="top", anchor="ne")
    canvas.create_image(0, 0, anchor=NW, image=photo)

    #####################################################
    request_number_label = ctk.CTkLabel(window, text="Search:")
    request_number_label.place(x=50, y=20)
    # Create the request number text box
    request_number_entry = ctk.CTkEntry(window, placeholder_text="REW-40108-666-101")
    request_number_entry.place(x=160, y=20)
    ######################################################
    name_label = ctk.CTkLabel(window, text="البرز صنعت نیلآب شرکت")
    name_label.place(x=525, y=5)
    #####################################################
    # Create a label for the project name choice box
    project_name_label = ctk.CTkLabel(window, text="Project Name:")
    project_name_label.place(x=51, y=60)

    # Create the project name choice box
    project_name_var = ctk.StringVar()
    options = ctk.CTkOptionMenu(
        window,
        variable=project_name_var,
        values=(
            "مرکزی",
            "رشت",
            "2 رودان",
            "جاسک آبرسانی",
            "التیمور فاضلاب",
            "عرب خین فاضلاب",
            "ساله 5 قم فاضلاب",
            "main",
        ),
    )
    project_name_var.set("main")
    options.place(x=160, y=60)

    def search_project():
        selected_value = project_name_var.get()
        checkForFile(request_number_entry.get(), chooseTable(selected_value))

    project_name_button = ctk.CTkButton(window, text="Search", command=search_project)
    project_name_button.place(x=320, y=60)

    ###################################################
    def uploadB():
        window.destroy()
        upload()

    upload_button = ctk.CTkButton(window, text="Upload New File", command=uploadB)
    upload_button.place(x=50, y=120)

    discriptions()
    # image_file = PhotoImage(file="C:/Users/alire/Desktop/rominas workspace/table1.png")
    # resized_image_file = image_file.subsample(3,3)
    # width , hight = resized_image_file.width(), resized_image_file.height()
    # canvas = Canvas(window, bg="#282424", width=width, height=hight)
    # canvas.pack(side='right')
    # canvas.create_image(0, 0, image=resized_image_file, anchor=NW)

    # Run the GUI
    window.mainloop()


####### main
View()
