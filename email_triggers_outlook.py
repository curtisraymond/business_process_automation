# ----------------------- CREATOR INFO -----------------------
#
# NAME:
# Curtis Raymond
# https://www.linkedin.com/in/curtis-raymond/
#
# ------------------------------------------------------------

import win32com.client
import psutil
import pythoncom
import re
import os
import urllib.request
import random
import string
from tkinter import *
from tkinter import messagebox

# Where you will save email attachments
attachment_dir = r"ADD PATH"

# Should email attachment(s) have unique file name(s)? ("True" for yes, "False" for no)
unique_file_name = True


# Loop that continuously scans your Microsoft Outlook inbox
class Handler_Class(object):
    def OnNewMailEx(self, receivedItemsIDs):
        # "recrivedItemIDs" is a collection of email IDs separated by a ","
        # Sometimes more than 1 email is received at the same moment
        for ID in receivedItemsIDs.split(","):
            mail = outlook.Session.GetItemFromID(ID)
            subject = mail.Subject
            print(subject)
            try:
                # If using Microsoft Outlook 2013 or older
                sender = mail.Sender.GetExchangeUser().PrimarySmtpAddress
            except Exception:
                # If using Microsoft Outlook 2016 or newer
                sender = mail.SenderEmailAddress
                pass
            if subject == "ADD EMAIL SUBJECT" and sender == "ADD EMAIL":
                mail.UnRead = False
                print(sender)
                for attachment in mail.Attachments:
                    if unique_file_name == False:
                        attachment.SaveAsFile(os.path.join(attachment_dir, attachment.FileName))
                    elif unique_file_name == True:
                        file_type = attachment.FileName[attachment.FileName.rfind('.'):]
                        file_name = ''.join(random.choices(string.ascii_uppercase + string.digits, k=30))
                        attachment.SaveAsFile(os.path.join(attachment_dir, file_name + file_type))


# Function to check if Microsoft Outlook is open on your computer
def check_outlook_open():
    list_process = []
    for pid in psutil.pids():
        p = psutil.Process(pid)
        # Append to the list of processes
        list_process.append(p.name())
    # If Microsoft Outlook is open then return True
    if 'OUTLOOK.EXE' in list_process:
        return True
    else:
        return False


while True:
    try:
        outlook_open = check_outlook_open()
    except:
        outlook_open = False
        # Error message if Microsoft Outlook is not open on your computer
        window = Tk()
        window.eval('tk::PlaceWindow %s center' % window.winfo_toplevel())
        window.withdraw()
        message = "Error!\n\nUnable to detect Microsoft Outlook on your desktop.\n" \
                  "Please open Microsoft Outlook first before running this automated process."
        messagebox.showerror('Microsoft Outlook Automated Process', message)
        window.deiconify()
        window.destroy()
        window.quit()
        os._exit(1)
    if outlook_open == True:
        outlook = win32com.client.DispatchWithEvents("Outlook.Application", Handler_Class)
        pythoncom.PumpMessages()