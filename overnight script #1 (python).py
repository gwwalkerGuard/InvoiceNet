import os
import re
import shutil
import subprocess
import sys
import warnings
from datetime import date, datetime, timedelta
from email.policy import strict
from pathlib import Path
import numpy as np
import O365
import pandas as pd
import pyodbc
import PyPDF2
import pytesseract
import win32com.client
from O365 import Account
from O365.utils.token import FileSystemTokenBackend
from pdf2image import convert_from_path
from PIL import Image
from pip import main
import fnmatch
import predict
from decouple import config

"""main funcs"""

def latestEmailDate(folder):
    latest = None
    for message in folder.get_messages(1):
        latest = message.created.date()
        latest = latest - timedelta(days=1)
    return latest
    
def authenticate(credentials):
    account = Account(credentials)
    if account.authenticate(scopes=['basic', 'message_all']):
        print('Authenticated!')

def refreshToken():
    O365.Connection.refresh_token

def folderinit(__class):
    locations = ["images","multipage","unreadables","predictions"]
    path = __class.path 
    if os.path.isdir(path):
        return False
    else: 
        os.makedirs(__class.path)
        for location in locations:
            new=f"{__class.path}\\{location}"
            os.mkdir(new)
        return True
   
def cleanFolder(path): 
    filelist = [ f for f in os.listdir(path) if f.endswith((".png",".jpg",".BMP")) ]
    if len(filelist)!=0:
        for f in filelist:
            os.remove(os.path.join(path, f))
    else: 
        print("folder already exists")
    
def getBatch(Automation):
    if folderinit(Automation):    
        for messages in Automation.invoices.get_messages(query=Automation.query,download_attachments=True): #get_message(download_attachements=False) 
            if messages.has_attachments:
                for attachemnets in messages.attachments:
                    attachemnets.save(location=Automation.path)
                    Automation.attachList.append(attachemnets.name)

    cleanFolder(Automation.path)

def sortedFiles(__class):
    
    memos = [file for file in os.listdir(__class.path) if 'memo' in file.lower()]
    for file in memos:
        shutil.move(os.path.join(__class.path, file), os.path.join(__class.path,"unreadables"))
    
    filenames = (file for file in os.listdir(__class.path) if file.endswith((".pdf")))
    for filename in filenames:

        file = open(os.path.join(__class.path, filename), 'rb')
        pdfReader = PyPDF2.PdfFileReader(file)
        totalPages = pdfReader.getNumPages()
        file.close()
        
        if totalPages > 1:    
            shutil.move(os.path.join(__class.path, filename), os.path.join(__class.path,"multipage"))
            
    filenames = (file for file in os.listdir(__class.path) if file.endswith((".pdf")))
    return filenames,__class.path

def passToModel(location):
    execution = ["python","predict.py","--field invoice_number vendor_name invoice_date total_amount",'--data_dir "{}"'.format(location),'--pred_dir "{}"'.format(os.path.join(location, "predictions"))]
    print(' '.join(execution))
    subprocess.run(' '.join(execution),check=True)
    
# def passToModel(files):
#     for file in files:

"""Test Functions"""
       
def testAPI(credentials):
    account = Account(credentials)
    m = account.new_message()
    m.to.add('walkerGa@guardsmangroup.com')
    m.subject = 'Testing!'
    m.body = "George Best quote: I've stopped drinking, but only while I'm asleep."
    m.send()

"""Class Functions"""

class Automation:
    def __init__(self):
        self.tk = FileSystemTokenBackend(token_filename="o365_token",token_path="/")
        # self.credentials =  #('c28e3ca7-785d-4bde-b9cc-7c62cdd30566', 'rxP8Q~xz2Zv9O8QViYHoEUOEhOEDdkZY2RK_TbgD') 
        # print(self.credentials)

        self.account = Account(config('credentials',default=''))
        self.mailbox = self.account.mailbox()
        self.inbox = self.mailbox.get_folder(folder_name='Inbox')
        self.invoices = self.inbox.get_folder(folder_name='Invoices')
        self.lateDate = latestEmailDate(self.invoices)
        self.path = os.getcwd()+"\\"+"Invoices"+"\\"+self.lateDate.strftime("%d %B, %Y")
        self.query = self.invoices.new_query().on_attribute('created_date_time').greater_equal(pd.to_datetime(self.lateDate)).less(pd.to_datetime(self.lateDate+timedelta(days=1)))
        self.attachList = []

"""main script"""

if __name__ == "__main__":
    emailAccount = Automation()
    # getBatch(emailAccount)
    # files,location = sortedFiles(emailAccount)
    # passToModel(location)
    
    

