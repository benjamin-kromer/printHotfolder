import win32api
import win32print
import os
import time
import shutil
import re
from datetime import datetime

#all_printers = win32print.EnumPrinters(2)

pdf_dir = "Y:\\HOTFOLDER_DRUCK\\INPUT\\"
archiv = "Y:\\HOTFOLDER_DRUCK\\ARCHIV\\"
problem = "Y:\\HOTFOLDER_DRUCK\\PROBLEMJOBS\\"

def checkForValidFiles():
    files = []
    filesInFolder = os.listdir(pdf_dir)
    for f in filesInFolder:
        if f[-3:] == "pdf":
            files.append(f)
    return files

def getActualTime():
    now = datetime.now()
    t = now.strftime("%d.%m.%Y, %H:%M:%S")
    return t

def deleteFile(f):
    while f in os.listdir(pdf_dir):
        try:              
            time.sleep(3)
            print getActualTime()+" REMOVING "+f+" FROM INPUT FOLDER!"
            os.remove(os.path.join(pdf_dir,f))
        except Exception as e:
            print e
            time.sleep(5)
    return None

while True:
    files = checkForValidFiles()
    if len(files) > 0:
        time.sleep(6) # WARTEN BIS SWITCH DEN PREFIX ENTFERNT HAT 
        files = checkForValidFiles() # ANSCHLIESSEND ORDNER NEU EINLESEN
        for f in files:
            pattern = re.compile('\_([a-zA-Z0-9\s\-]+)\.pdf')
            match = pattern.search(f)
            try:
                printer = match.group(1)
                defaultPrinter = win32print.GetDefaultPrinter()
                if defaultPrinter != printer:
                    win32print.SetDefaultPrinter(printer)
                print getActualTime()+" PRINTING FILE "+ f +" on "+printer
                win32api.ShellExecute(0,"print", os.path.join(pdf_dir,f), None,  ".",  0)
                print getActualTime()+" COPY "+f+" to ARCHIV!"
                shutil.copy(os.path.join(pdf_dir,f),os.path.join(archiv,f))
                deleteFile(f)
                print getActualTime()+" FILE "+f+" SUCCESSFULLY PRINTED!"
            except Exception as e:
                print e
                shutil.copy(os.path.join(pdf_dir,f),os.path.join(problem,f))
                deleteFile(f)