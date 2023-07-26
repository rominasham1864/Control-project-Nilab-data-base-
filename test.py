import tkinter as tk
from tkinter import ttk


# Create the main window
window = tk.Tk()
window.geometry("500x500")

# Create a label for the request number text box
request_number_label = tk.Label(window, text="Request Number:")
request_number_label.place(x=50, y=20)

# Create the request number text box
request_number_entry = tk.Entry(window)
request_number_entry.place(x=160, y=20)

# Create a search button for the request number
def save_request_number():
    req_n = request_number_entry.get()
request_number_button = tk.Button(window, text="Search", command=save_request_number)
request_number_button.place(x=300, y=15)

# Create a label for the project name choice box
project_name_label = tk.Label(window, text="Project Name:")
project_name_label.place(x=51, y=60)

# Create the project name choice box
project_name_var = tk.StringVar()
options =ttk.Combobox(window, width = 17, textvariable = project_name_var)
options.place(x=160 , y=60)
options['values'] = ('خین',
                     'التیمور',
                     'جاسک',
                     'رودان')
                     

# Create a search button for the project name
def save_project_name():
    global project_name
    project_name = project_name_var.get()
project_name_button = tk.Button(window, text="Search",  command=save_project_name)
project_name_button.place(x=300, y=57)

def save_project_name():
    global project_name
    project_name = project_name_var.get()
upload_button = tk.Button(window, text="Upload New File",  command=save_project_name)
upload_button.place(x=50, y=100)

image_label = tk.Label(window)
image_file = tk.PhotoImage(file="C:/Users/alire/Desktop/rominas workspace/logo5.png")
resized_image_file = image_file.subsample(3, 3)

    # Set the image for the label
image_label.config(image=resized_image_file)
image_label.place(x=400, y=10)
# Run the GUI
window.mainloop()