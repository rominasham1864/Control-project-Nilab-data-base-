import tkinter as tk
from tkinter import ttk

def change_row_text_color(row_id, color):
    tree.item(row_id, tags=(color,))

root = tk.Tk()

# Create a Treeview widget
tree = ttk.Treeview(root)
tree.pack()

# Define the table columns
tree["columns"] = ("Name", "Age", "Country")

# Format the table columns
tree.column("#0", width=0, stretch=tk.NO)  # Hide the default first column
tree.column("Name", width=100)
tree.column("Age", width=50)
tree.column("Country", width=100)

# Create table headers
tree.heading("Name", text="Name")
tree.heading("Age", text="Age")
tree.heading("Country", text="Country")

# Insert table data with internal identifiers
row1 = tree.insert("", tk.END, text="Row 1", values=("John Doe", "25", "USA"))
row2 = tree.insert("", tk.END, text="Row 2", values=("Jane Smith", "30", "Canada"))
row3 = tree.insert("", tk.END, text="Row 3", values=("Bob Johnson", "40", "Australia"))

# Change the text color of a specific row
change_row_text_color(row2, "red")

root.mainloop()