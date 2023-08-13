import tkinter as tk
import os
import openpyxl
import pymysql
import tkinter as tk
from tkinter import *
from tkinter.ttk import *
import tkinter as tk
from tkinter import ttk
import os
import sqlite3
from tkinter import filedialog
from openpyxl.styles import PatternFill

def handle_double_click(event):
    # Get the selected row
    selected_row = tree.focus()

    # Perform the desired action
    # For example, print the selected row's data
    data = tree.item(selected_row)['values']
    print("Selected row:", data)

root = tk.Tk()

# Create a Treeview widget
tree = ttk.Treeview(root)
tree['columns'] = ('Column 1', 'Column 2')

# Insert some sample data
tree.insert('', 'end', text='Row 1', values=('Value 1', 'Value 2'))
tree.insert('', 'end', text='Row 2', values=('Value 3', 'Value 4'))
tree.insert('', 'end', text='Row 3', values=('Value 5', 'Value 6'))

# Bind the double-click event to the Treeview widget
tree.bind("<Double-1>", handle_double_click)

# Display the Treeview widget
tree.pack()

root.mainloop()