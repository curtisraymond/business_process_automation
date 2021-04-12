# ----------------------- CREATOR INFO -----------------------
#
# NAME:
# Curtis Raymond
# https://www.linkedin.com/in/curtis-raymond/
#
# ------------------------------------------------------------

import email
import imaplib
import os
import os.path
import random
import string
import re
import win32com.client
import psutil
import pythoncom
import webbrowser
import pyautogui
from imagesearch import *
from time import sleep
from tkinter import *
from tkinter import messagebox
import fnmatch

# Connect to the Gmail IMAP server
mail = imaplib.IMAP4_SSL("imap.gmail.com")

# Your Gmail credentials
user = "test558.simplyautomated@gmail.com"
pwd = "simplyautomated"
mail.login(user, pwd)

# Where you will save the downloadable url(s)
file_dir = r"C:\Users\curtis\Desktop\Projects\email_triggers_gmail\test_2"

# Should the downloadable url(s) be given unique file name(s)? ("True" for yes, "False" for no)
unique_file_name = True

# All url(s) that are found within the body of an email are organized as a list. Please indicate only the
# relevant url(s). However, the url(s) indicated must be in sequence order.
# ----------------------------------------------------------------------------------------------------------
# [For example] url's: (1 to 3), or (2 to 7), or even (1 to 1) if only the 1st url is relevant.
#               url's: (1 and 3) is NOT valid as they are not in sequence order.
# ----------------------------------------------------------------------------------------------------------
url_start_location = 1
url_end_location = 4


# Error messages if the automated process fails at any given point
def error(message):
    window = Tk()
    window.eval('tk::PlaceWindow %s center' % window.winfo_toplevel())
    window.withdraw()
    messagebox.showerror('Gmail Automated Process', message)
    window.deiconify()
    window.destroy()
    os._exit(1)


# Loop that continuously scans your Gmail inbox
def loop():
    mail.select("INBOX")
    n = 0
    (retcode, messages) = mail.search(None, '(UNSEEN FROM "info.simplyautomated@gmail.com" SUBJECT "hello")')
    if retcode == 'OK':
        for num in messages[0].split():
            n = n + 1
            print(n)
            typ, data = mail.fetch(num, '(RFC822)')
            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    print(msg['From'])
                    print(msg['Subject'])
                    for part in msg.walk():
                        if part.get_content_type() == 'text/plain':
                            body_content = part.get_payload()
                            urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', body_content)
                            for y in range(url_start_location - 1, url_end_location):
                                print(urls[y])
                                pyautogui.moveTo(0, 500)
                                webbrowser.open(urls[y])
                                timeout = 30
                                timeout_start = time.time()
                                while time.time() < timeout_start + timeout:
                                    pos_chrome = imagesearch(r"C:\Users\curtis\Desktop\Projects\email_triggers_gmail\images\save_pdf_2.png")
                                    pos_iexplore = imagesearch(r"C:\Users\curtis\Desktop\Projects\email_triggers_gmail\images\save_pdf_iexplore_v1.png")
                                    if pos_chrome[0] == -1 and pos_iexplore[0] == -1:
                                        pos_iexplore = imagesearch(r"C:\Users\curtis\Desktop\Projects\email_triggers_gmail\images\save_pdf_iexplore_v2.png")
                                    if pos_chrome[0] != -1 or pos_iexplore[0] != -1:
                                        timeout = 10
                                        timeout_start = time.time()
                                        while time.time() < timeout_start + timeout:
                                            pyautogui.hotkey('ctrl', 'shift', 's')
                                            pyautogui.hotkey('ctrl', 's')
                                            sleep(2)
                                            pos_chrome = imagesearch(r"C:\Users\curtis\Desktop\Projects\email_triggers_gmail\images\save_as.png")
                                            pos_iexplore = imagesearch(r"C:\Users\curtis\Desktop\Projects\email_triggers_gmail\images\save_as_iexplore.png")
                                            if pos_chrome[0] != -1 or pos_iexplore[0] != -1:
                                                if unique_file_name == True:
                                                    file_name = ''.join(random.choices(string.ascii_uppercase + string.digits, k=30))
                                                    pyautogui.typewrite(str(file_name))
                                                elif unique_file_name == False:
                                                    pyautogui.hotkey('ctrl', 'c')
                                                    root = Tk()
                                                    root.withdraw()
                                                    file_name = root.clipboard_get()
                                                pyautogui.press('tab', presses=6)
                                                pyautogui.hotkey('enter')
                                                pyautogui.typewrite(str(file_dir))
                                                pyautogui.hotkey('alt', 's')
                                                file = []
                                                while not file:
                                                    file = list(filter(lambda f: fnmatch.fnmatch(f, file_name + "*"), os.listdir(file_dir)))
                                                    sleep(0.025)
                                                pos_chrome = imagesearch(r"C:\Users\curtis\Desktop\Projects\email_triggers_gmail\images\chrome_close.png")
                                                if pos_chrome[0] != -1:
                                                    pyautogui.click(pyautogui.moveTo(pos_chrome[0], pos_chrome[1]))
                                                elif pos_chrome[0] == -1:
                                                    os.system("taskkill /im iexplore.exe /F")
                                                break
                                        if pos_chrome[0] == -1 and pos_iexplore[0] == -1:
                                            message = "Timeout Error!\nUnable to save the downloadable file within the given url."
                                            error(message)
                                        break
                                if pos_chrome[0] == -1 and pos_iexplore[0] == -1:
                                    message = "Timeout Error!\nUnable to detect a downloadable file within the given url."
                                    error(message)
                    if data == 'eject':
                        ctypes.windll.WINMM.mciSendStringW(u"set cdaudio door open", None, 0, None)
                    typ, data = mail.store(num, '+FLAGS', '\\Flagged')  # '\\Flagged' means to star the message


if __name__ == '__main__':
    try:
        while True:
            loop()
    finally:
        os._exit(1)