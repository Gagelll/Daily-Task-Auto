# pip install pywin32  #if you not installed yet
import win32com.client
from datetime import datetime
import datetime
import os
import shutil
import re
import wget

# set up connection to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# for sub folder, add <.folder("your folder name")>
inbox = outlook.GetDefaultFolder(6).folders("Procurement")

# Access to the email in the inbox
messages = inbox.Items

#Todays Date
today_date = str(datetime.date.today())

try:    
    for message in messages:
        sent_date = str(message.senton.date())
        if sent_date == today_date:
            print(sent_date)
            x = message.body
            link = re.findall('<(.*)>', x)
            file = str(link[0])
            print(file)
            file_url = file
            file_name = wget.download(file_url, 'V:\Data & Analytics\Procurement')
except AttributeError:
    print("No Sender")

#gets body of email
# x = message.Body

# #takes string between <>
# link = re.findall('<(.*)>', x)

# file = str(link[0])

#Target Directories
path = "V:\Data & Analytics\Procurement"
GoodsAnalysis = "V:\Data & Analytics\Purchasing\Goods Purchased Data\Good Purchased Analysis"
ApprovalAudit =  "V:\Data & Analytics\Purchasing\Goods Purchased Data\Approval Audit"
Goods = "V:\Data & Analytics\Purchasing\Goods Purchased Data\GoodsPurchased"

#List Directory
files = os.listdir(path)

#loop through folder and move correct files
try:
    for filename in files:
        if "GoodsPurchasedAnalysis" in filename:
            shutil.move(path + '\\' + filename, GoodsAnalysis + '\\' + filename)
            print(filename + ' ' + 'has been moved')
        elif "ApprovalAudit" in filename:
            shutil.move(path +'\\' + filename, ApprovalAudit + '\\' + filename)
            print(filename + ' ' + 'has been moved')
        elif "GoodsPurchased" in filename:
            shutil.move(path + '\\' + filename, Goods + '\\' + filename)
            print(filename + ' ' + 'has been moved')
        else:
            print('File Ignored')
except AttributeError:
    print('Error')

today_date = str(datetime.date.today())

path = 'V:\Data & Analytics\Purchasing\Goods Purchased Data\Approval Audit'

audits = os.listdir(path)

for file in audits:
    lastmodified = os.stat(path + "\\" + file).st_mtime
    date_modified = date_modified = str(datetime.date.fromtimestamp(lastmodified))
    if date_modified !=  today_date:
        os.remove(path + "\\" + file)
        print(file + ' ' + 'deleted')
    else:
        print(file + ' ' + 'is present')

today = datetime.date.today().isoweekday()

if today == 1 or today == 3:

    goods_purchased_path = 'V:\Data & Analytics\Purchasing\Goods Purchased Data\GoodsPurchased'

    gp = os.listdir(goods_purchased_path)

    for file in gp:
            file_name = str(file)
            lastmodified = os.stat(goods_purchased_path + "\\" + file).st_mtime
            date_modified = date_modified = str(datetime.date.fromtimestamp(lastmodified))
            if file_name == 'GoodsPurchased_2020.csv' or file_name == 'GoodsPurchased_2019.csv':
                pass
            else:
                if date_modified !=  today_date:
                    os.remove(goods_purchased_path + "\\" + file)
                    print(file + ' ' + 'deleted')
                else:
                    print(file + ' ' + 'is present')
else:
    print('Wrong Day')

analysis_path = 'V:\Data & Analytics\Purchasing\Goods Purchased Data\Good Purchased Analysis'

gp = os.listdir(analysis_path)

for file in gp:
        file_name = str(file)
        lastmodified = os.stat(analysis_path + "\\" + file).st_mtime
        date_modified = date_modified = str(datetime.date.fromtimestamp(lastmodified))
        if file_name == 'GoodsPurchasedAnalysis_2020.csv' or file_name == 'GoodsPurchasedAnalysis_2019.csv':
            pass
        else:
            if date_modified !=  today_date:
                os.remove(analysis_path + "\\" + file)
                print(file + ' ' + 'deleted')
            else:
                print(file + ' ' + 'is present')