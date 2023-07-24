#   TODO: Run through every billing statement, taking only the regular
#   monthly charges, equipment and services, service fees, and taxes sections
#   
#   Then output them to a spreadsheet.

import os
from tkinter import filedialog as fd
from tkinter import messagebox
from tkinter import *
from helpers import getRMCPage, getCharges, outputToSpreadsheet

def directorySelectCallBack():
    folder_selected = fd.askdirectory()
    entry_text.set(folder_selected)
    return

def scanCallBack():
    if os.path.isdir(entry_text.get()):
        outputToSpreadsheet(entry_text.get())
        messagebox.showinfo("Info", "Charges written to output.xls")
    else:
        messagebox.showerror("Error", "Directory does not exist!")
    return
    
root = Tk()
root.title('Xfinity billing statement reader')
root.resizable(False, False)
root.geometry('300x150')
L1 = Label(root, text="Statements directory")
L1.pack( side = LEFT)
entry_text = StringVar()
E1 = Entry(root, bd = 5, state=DISABLED, textvariable=entry_text)
E1.pack( side = LEFT)
B1 = Button(root, text= "...", command = directorySelectCallBack)
B1.pack(side = LEFT)

B2 = Button(root, text="Scan into spreadsheet", command = scanCallBack)
B2.place(relx=0.5, rely=0.8, anchor=CENTER)
root.mainloop()
#root.withdraw()
#folder_selected = filedialog.askdirectory()






    
    
