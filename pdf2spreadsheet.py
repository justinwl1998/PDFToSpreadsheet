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
root.geometry('400x150')

L1 = Label(root, text="Statements directory", justify="left", anchor="w")
#L1.place(relx=0, rely=0.3, anchor = W)
L1.grid(column=0, row=3, sticky=W, pady=(15,0), padx=(30,0))

entry_text = StringVar()
E1 = Entry(root, width=45, bd = 5, state=DISABLED, textvariable=entry_text)
E1.grid(column=0, row=4, sticky=W, padx=(30,0))
#E1.place(relx=0.25, rely=0.4)
B1 = Button(root, text= "...", command = directorySelectCallBack, width=8)
B1.grid(column=1, row=4, sticky=W)
#B1.place(relx=1.0, rely=0.48, anchor=E)

B2 = Button(root, text="Convert to spreadsheet", command = scanCallBack)
B2.place(relx=0.5, rely=0.8, anchor=CENTER, height=32, width=150)
root.mainloop()







    
    
