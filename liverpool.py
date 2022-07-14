from random import Random
import win32com.client
from datetime import date, timedelta
import calendar
import datetime
import os
import glob
import openpyxl

#Empty dirty file folder for speed
dir = "V:\Data & Analytics\Dirty_Files"

dir_contents = os.listdir(dir)

for file in dir_contents:
    os.remove(dir + '\\' +file)

# set up connection to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# for sub folder, add <.folder("your folder name")>
inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items

#Todays Date
today_date = str(datetime.date.today())

def delete_yesterdays_data(hotel_code, data):
    path = 'V:\\Data & Analytics\\Scheduling\\' + hotel_code + '\\' + data
    dir = os.listdir(path)
    for file in dir:
        lastmodified = os.stat(path + "\\" + file).st_mtime
        date_modified = str(datetime.date.fromtimestamp(lastmodified))
        if date_modified != today_date:
            os.remove(path + "\\" + file)
            print('file deleted')
        else:
            print('No files to delete')
        
delete_yesterdays_data('Liverpool files', 'Forecast')

delete_yesterdays_data('Jewellrey Quarter files', 'Forecast')

my_date = date.today()

dow = calendar.day_name[my_date.weekday()]

for email in messages:
    if dow == 'Monday':
        saturday = my_date + timedelta(days=-2)
        sunday = my_date + timedelta(days=-1)
        weekend_dates = [str(saturday), str(sunday), str(my_date)]
        if "operation audit" in email.subject.lower() and str(email.senton.date()) in weekend_dates:
            print(str(email.senton.date()))
            attachments = email.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment)
            attachment.SaveASfile("V:\\Data & Analytics\\Dirty_Files"+ '\\' + attachment_name)
        elif "bhxbn rexmax" in email.subject.lower() and str(email.senton.date()) == today_date:
                print(email.subject)
                attachments = email.Attachments
                attachment = attachments.Item(1)
                attachment_name = str(attachment)
                attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\Jewellrey Quarter files\\Forecast"+ '\\' + attachment_name)
                print(attachment_name + " " + "has been saved")
        elif "lpllc revmax" in email.subject.lower() and str(email.senton.date()) == today_date:
            print(email.subject)
            attachments = email.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment)
            attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\Liverpool files\\Forecast"+ '\\' + attachment_name)
            print(attachment_name + " " + "has been saved")
        else:
            pass
    else:
        if "operation audit" in email.subject.lower() and str(email.senton.date()) == today_date:
            print(email.subject)
            attachments = email.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment)
            attachment.SaveASfile("V:\\Data & Analytics\\Dirty_Files"+ '\\' + attachment_name)
            print(attachment_name + " " + "has been saved")
        elif "bhxbn rexmax" in email.subject.lower() and str(email.senton.date()) == today_date:
            print(email.subject)
            attachments = email.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment)
            attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\Jewellrey Quarter files\\Forecast"+ '\\' + attachment_name)
            print(attachment_name + " " + "has been saved")
        elif "lpllc revmax" in email.subject.lower() and str(email.senton.date()) == today_date:
            print(email.subject)
            attachments = email.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment)
            attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\Liverpool files\\Forecast"+ '\\' + attachment_name)
            print(attachment_name + " " + "has been saved")
        else:
            pass

#Connects to excel and finds the appropiate files
o = win32com.client.Dispatch("Excel.Application")
o.Visible = False

input_dir = r"V:\\Data & Analytics\\Dirty_Files"
output_dir = r"V:\\Accounts\\2022\\Hilton Hotels\\Data"
files = glob.glob(input_dir + "/*.xls")

#Moves and renames the corrupt files making them uncorrupt
for filename in files:
    file = os.path.basename(filename)
    output = output_dir + '\\' + file.replace('.xls','.xlsx')
    wb = o.Workbooks.Open(filename)
    wb.ActiveSheet.SaveAs(output,51)
    wb.Close(True)

path = "V:\\Accounts\\2022\\Hilton Hotels\\Data"
files = os.listdir(path)
ops_audits = [f for f in files if f[-4:] == 'xlsx']

for f in ops_audits:
    ss=openpyxl.load_workbook(path + '\\' + f)
    sheet_name = str(ss.get_sheet_names()[0])
    if sheet_name != "Sheet1":
        ss_sheet = ss[sheet_name]
        ss_sheet.title = 'Sheet1'
        ss.save(path + '\\' + f)
        print('sheet name has been changed')