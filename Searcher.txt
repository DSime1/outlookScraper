import os
import win32com.client
from datetime import datetime
from tkinter import messagebox
import shutil

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts;

def ensure_dir(file_path):
    directory = file_path
    if not os.path.exists(directory):
        os.makedirs(directory)

def valiDate(date_text):
    print(date_text)
    try:
        datetime.strptime(date_text, '%d-%m-%Y')
    except :
        messagebox.showerror("Error!", "Incorrect data format, should be dd-mm-yyyy", icon='error')
        return "wrong"

def getMail(acc, kk, dat, fold, dpath, dfold):

 #VALIDATION OF INPUTS:

    if not acc:
        messagebox.showerror("Error!", "Account name cannot be empty value!\nPlease insert valid account name", icon='error')
        return

    if not fold:
        answer = messagebox.askquestion("Warning!", "Subfolder field is empty!\nProgram will use Inbox as default folder for current search.\nDo you wish to proceed?", icon='warning')
        if not answer == 'yes':
            #messagebox.showinfo('Return', 'You will now return to the application screen')
            return

    if not dat:
        messagebox.showerror("Error!", "Date field cannot be empty!\nPlease insert a date in fomat: dd-mm-yyyy", icon='error')
        return None

    if valiDate(dat) =="wrong":
        return

    if not kk:

        answer = messagebox.askquestion("Warning!",
                                        "Keyword field is empty!\nDo you wish to proceed without a keyword as parameter?",
                                        icon='warning')
        if not answer == 'yes':
            # messagebox.showinfo('Return', 'You will now return to the application screen')
            return None

        else:
            answer = messagebox.askquestion('Double confirmation',
                                            'Program will save all mails and attachmentes from given folder after provided date!\nDo you really want to proceed?',
                                            icon='warning')

            if not answer == 'yes':
                # messagebox.showinfo('Return', 'You will now return to the application screen')
                return None

#//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

# start of function:

    downloadPathText2 = dpath + "\\" + dfold

    for x in accounts:

        if str(x) == acc:
            inb = outlook.GetDefaultFolder(6)
            if fold == "":

                inb = outlook.GetDefaultFolder(6)

            elif fold!="Inbox":

                try:
                    inb = outlook.GetDefaultFolder(6).folders(fold)

                except:
                    messagebox.showerror("Error!", "outlook folder name not found! \nPlease verify it and try again", icon='error')
                    return

            messages = inb.Items
        else:
            messagebox.showerror("Error!", "Account not found in your Outlook\nPlease verify it and try again", icon='error')

        ncounter = 0

        for mail in messages:


            if kk.lower() in str(mail.subject).lower():
                print(dat)

                dat2 = datetime.strptime(str(dat),'%d-%m-%Y')
                dateMail = datetime.strptime(mail.receivedTime.strftime('%d-%m-%Y'), '%d-%m-%Y')

                if dat2.date() < dateMail.date():
                    print(type(dat))
                    ncounter = ncounter + 1

                    #print(dat < mail.receivedTime.strftime('%d-%m-%Y'))

                    print(mail.receivedTime)
                    print(mail.receivedTime.strftime('%d-%m-%Y'))
                    print(dat)

                    downloadPathText2 = dpath + "\\" + dfold
                    print(downloadPathText2)

                    ensure_dir(str(downloadPathText2))

                    mail.SaveAs(os.path.join(downloadPathText2, "(" + str(ncounter) + ")" + str(mail) + ".msg"))

                    try:
                        for attachment in mail.Attachments:
                            attachment.SaveAsFile(
                                os.path.join(downloadPathText2, "(" + str(ncounter) + ")" + str(attachment)))

                    except:

                        pass

        if ncounter == 0:
            messagebox.showinfo("Info",
                                "No result found for given keyword and date\nTry with different parameters")
        else:
            messagebox.showinfo("Download completed!", "Mails found: " + str(ncounter)  +"\nFiles saved at: " + downloadPathText2)

def MoveFiles( downlowadP ):

    print(downlowadP)
    answer = messagebox.askquestion("Warning!",
                                    "You will copy all files in current operational folder to a new destination!\n"
                                    "Depending on how many files this operation may take some time.\nDo you wish to proceed?",
                                    icon='warning')
    if not answer == 'yes':
        # messagebox.showinfo('Return', 'You will now return to the application screen')
        return None


    try:

        src = downlowadP

        print("Src" + src)
        src_files = os.listdir(src)

        dest = os.path.expanduser("~/Desktop/Pending")
        ensure_dir(dest)
        k = 0
        for file_name in src_files:
            full_file_name = os.path.join(src, file_name)
            if os.path.isfile(full_file_name):
                shutil.copy(full_file_name, dest)
                k = k + 1
        messagebox.showinfo("Completed!", str(k) + " files from:\n" + src + "\nhave been moved to:\n" + dest)

    except:
        messagebox.showerror("Error!", "Something went wrong!\nPlease verify the path",
                             icon='error')

