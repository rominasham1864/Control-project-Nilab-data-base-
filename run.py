from tkinter import Tk
from tkinter.filedialog import askopenfilename
print("Please select your file")
# time.sleep(1)
filelocation = askopenfilename()
file_name=filelocation.split("/")[-1]
print(filelocation)
print("C:\Users\alire\Desktop\rominas workspace\REM-40109-777-272.xlsm")