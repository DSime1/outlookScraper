import tkinter as tk
import searcher as sear
import os
from tkinter import *
from tkinter import filedialog, messagebox
from datetime import datetime
from PIL import ImageTk, Image



downloadPathText = os.path.expanduser("~\Desktop")

def changePath():
    global downloadPathText
    selectedFolder = filedialog.askdirectory()
    downloadPathText = selectedFolder
    defPath.config(text="Current operation folder: \n\n" + downloadPathText)


def ShowInfo():

    messagebox.showinfo("Info","Developer: Simeone, Davide\nContact:\nVersion: 1\nDescription:\nSoftware works only if outlook application is installed and open before launch.\n"
                               "If you started Outlook app after please close this program and re-open it.\n"
                        "To move files please ensure the proper folder has been selected in the path!"
                               )

root = Tk()

root.wm_iconbitmap('Icon.ico')
root.wm_title('Outlook Scraper')

root.geometry("600x400")

imageframe = Frame(root)
imageframe.grid(row=0, column=0)

resized = Image.open('images.png').resize((130, 100), Image.ANTIALIAS)
ssb = ImageTk.PhotoImage(resized)
panel = tk.Label(imageframe, image=ssb)
panel.grid(row=0, column=0, padx=5)
HelpBtn = tk.Button(imageframe, text='Help!', command=lambda: ShowInfo())
HelpBtn.grid(row=1, column=0, pady=5)

loginFrame = Frame(root)
loginFrame.grid(row=0, column=1, pady=10)

lblmail = tk.Label(loginFrame, text="E-mail account")
lblmail.grid(row=0, column=0, pady=5, padx=10)

e1 = tk.Entry(loginFrame, width=30)
e1.grid(row=1, column=0, pady=5, padx=10)

lblkey1 = tk.Label(loginFrame, text="Search keyword")
lblkey1.grid(row=2, column=0, pady=5, padx=10)

e2 = tk.Entry(loginFrame, width=30)
e2.grid(row=3, column=0, pady=5, padx=10)

lbldate = tk.Label(loginFrame, text="Insert start date (dd-mm-yyyy)")
lbldate.grid(row=2, column=1, pady=5, padx=10)

v = StringVar(loginFrame, value=datetime.today().strftime('%d-%m-%Y'))

e3 = tk.Entry(loginFrame, textvariable=v, width=30)
e3.grid(row=3, column=1, pady=5, padx=10)

g = StringVar(loginFrame, value="Inbox")
lblSubfolder = tk.Label(loginFrame, text="Outlook folder name")
lblSubfolder.grid(row=0, column=1, pady=5, padx=10)

e4 = tk.Entry(loginFrame, textvariable=g, width=30)
e4.grid(row=1, column=1, pady=5, padx=10)



pathsButtonsframe = Frame(root)
pathsButtonsframe.grid(row=2, column=0, columnspan=2)

downloadPathBtn = tk.Button(pathsButtonsframe, text='Select operation folder', command=lambda: changePath(), bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
downloadPathBtn.grid(row=0, column=0, pady=30, padx=5)

defPath = tk.Label(root, text="Current operation folder: \n\n" + downloadPathText)
defPath.grid(row=2, column=0, columnspan=2, pady=30, padx=30)

moveFilesBtn = tk.Button(pathsButtonsframe, text='Move last files', command=lambda: sear.MoveFiles(downloadPathText), bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
moveFilesBtn.grid(row=0, column=1, pady=30, padx=5)

browse_OutlookBtn = tk.Button(pathsButtonsframe, text='Start!', command=lambda: sear.getMail(str(e1.get()), str(e2.get()), str(e3.get()), str(e4.get()), str(downloadPathText), datetime.today().strftime('%Y-%m-%d')), bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
browse_OutlookBtn.grid(row=4, column=0, columnspan=2, pady=50, padx=70)

root.mainloop()





