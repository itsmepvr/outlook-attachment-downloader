# Outlook Attachmeent Downloader - Windows
# Downloads attachments from unread mails and continously monitors for new mails
# Author: Venkata Ramana P
# <pvrreddy155@gmail.com>
# <github.com/itsmepvr>
# -*- mode: python ; coding: utf-8 -*-
# >= Python 3.6

import datetime
import os
import win32com.client
import zipfile
from tkinter import *
import psutil
import sys

# Extend Recurrsion limit warning
sys.setrecursionlimit(10**6) 

# Current date
today = datetime.date.today()
print("Date: "+today)
td_path = str(today)

# Path to save attachments
user_path = os.path.expanduser("~/Desktop/Attachments")

# Path to the file --save
path = os.path.join(user_path, td_path)
if not os.path.isdir(path):
    os.mkdir(path)    

#Outlook API call
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) 
# All messages in inbox
messages = inbox.Items

def saveattachments(subject):
    for message in messages:
        if subject == ''  and message.Unread and message.Senton.date() == today:
            attachments = message.Attachments
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                print("File saved "+str(attachment))
                try:
                    zipExtract(attachment, path)
                except:
                    pass    
                message.Unread = False
        elif (message.Subject).replace('Fwd: ', '') == subject and message.Unread and message.Senton.date() == today:
            attachments = message.Attachments
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                print("File saved "+str(attachment))
                try:
                    zipExtract(attachment, path)
                except:
                    pass    
                message.Unread = False
    # Re run function for next messges            
    saveattachments(subject)

# Extract Zip File and delete zip file
def zipExtract(fname, path):
    zippath = os.path.join(path, str(fname))
    file = zipfile.ZipFile(zippath)
    for filename in file.namelist():
        file.extract(filename,path)
    file.close()
    try:    
        os.remove(zippath)
    except:
        # Sometims windows doesnt permit os functions which can be ignored
        print("Permision denied for removing file. Continung without removing")    
        pass

# Check if outlook running, if not start the app
def outlook_is_running():
    import win32ui
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        print("Outlook already running")
        return True
    except win32ui.error:
        print("Outlook is not running")
        return False

# Load --main
if __name__ == "__main__":
    if not outlook_is_running():
        os.startfile("outlook")
        print("Starting Outlook App...")
    # Subject -- Alert    
    saveattachments("Alert")    

