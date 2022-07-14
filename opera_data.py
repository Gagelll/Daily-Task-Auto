# pip install pywin32  #if you not installed yet
import win32com.client
import datetime
import os

# set up connection to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# for sub folder, add <.folder("your folder name")>
inbox = outlook.GetDefaultFolder(6).folders("Chardon")

# Access to the email in the inbox
messages = inbox.Items

#Todays Date
today_date = str(datetime.date.today())

try:
    for message in messages:
        #check the date of the email to only pull todays files
        if str(message.senton.date()) == today_date:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment).lower()
            if 'ediap' in attachment_name:
                if 'forecast' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\EDIAP files\\Forecast"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'statistics' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\EDIAP files\\Actuals"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'manager_report' in attachment_name:
                    attachment.SaveASfile("V:\\Accounts\\2022\\Chardon\\Holiday Inn Express Edinburgh Airport\\Data and Analytics\\Managers Flash\\Current Month"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'trial_balance' in attachment_name:
                    attachment.SaveASfile("V:\\Accounts\\2022\\Chardon\\Holiday Inn Express Edinburgh Airport\\Data and Analytics\\Trial Balance\\Current Month"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'data' and 'blocks' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Opera Data"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                else:
                    (attachment_name + ' ' + 'has been ignored')
            elif 'glwth' in attachment_name:
                if 'forecast' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\GLWTH files\\Forecast"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'statistics' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\GLWTH files\\Actuals"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'manager_report' in attachment_name:
                    attachment.SaveASfile("V:\\Accounts\\2022\\Chardon\\Holiday Inn Express Theatreland Glasgow\\Data and Analytics\\Managers Flash\\Current Month"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'trial_balance' in attachment_name:
                    attachment.SaveASfile("V:\\Accounts\\2022\\Chardon\\Holiday Inn Express Theatreland Glasgow\\Data and Analytics\\Trial Balance\\Current Month"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'data' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Opera Data"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'blocks' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Opera Data"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                else:
                    (attachment_name + ' ' + 'has been ignored')
            elif 'glwgc' in attachment_name:
                if 'forecast' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\GLWGC files\\Forecast"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'statistics' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Scheduling\\GLWGC files\\Actuals"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'manager_report' in attachment_name:
                    attachment.SaveASfile("V:\\Accounts\\2022\\Chardon\\Holiday Inn Theatreland Glasgow\\Data and Analytics\\Managers Flash\\Current Month"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'trial_balance' in attachment_name:
                    attachment.SaveASfile("V:\\Accounts\\2022\\Chardon\\Holiday Inn Theatreland Glasgow\\Data and Analytics\\Trial Balance\\Current Month"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'data' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Opera Data"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                elif 'blocks' in attachment_name:
                    attachment.SaveASfile("V:\\Data & Analytics\\Opera Data"+ '\\' + attachment_name)
                    print(attachment_name + ' ' + 'has been saved')
                else:
                    (attachment_name + ' ' + 'has been ignored')
except AttributeError:
    print("Email doesn't exist")

def delete_yesterdays_data(hotel_code, data):
    path = 'V:\\Data & Analytics\\Scheduling\\' + hotel_code + ' ' + 'Files\\' + data
    dir = os.listdir(path)
    for file in dir:
        lastmodified = os.stat(path + "\\" + file).st_mtime
        date_modified = str(datetime.date.fromtimestamp(lastmodified))
        if date_modified != today_date:
            os.remove(path + "\\" + file)
            print('file deleted')
        else:
            print('No files to delete')

delete_yesterdays_data('GLWGC', 'Actuals')

delete_yesterdays_data('GLWGC', 'Forecast')

delete_yesterdays_data('GLWTH', 'Actuals')

delete_yesterdays_data('GLWTH', 'Forecast')

delete_yesterdays_data('EDIAP', 'Actuals')

delete_yesterdays_data('EDIAP', 'Forecast')

def deputy_delete(path):
    dir = os.listdir(path)
    for file in dir:
        lastmodified = os.stat(path + "\\" + file).st_mtime
        date_modified = str(datetime.date.fromtimestamp(lastmodified))
        if date_modified != today_date:
            os.remove(path + "\\" + file)
            print('file deleted')
        else:
            print('No files to delete')

deputy_delete('V:\Data & Analytics\HR Folder\Deputy\Output')
