import datetime
import os

def delete_yesterdays_data(hotel_code, data):
    today_date = str(datetime.date.today())
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

def deputy_delete(path):
    today_date = str(datetime.date.today())
    dir = os.listdir(path)
    for file in dir:
        lastmodified = os.stat(path + "\\" + file).st_mtime
        date_modified = str(datetime.date.fromtimestamp(lastmodified))
        if date_modified != today_date:
            os.remove(path + "\\" + file)
            print('file deleted')
        else:
            print('No files to delete')