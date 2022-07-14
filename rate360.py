# pip install pywin32  #if you not installed yet
import win32com.client
import datetime
import tkinter as tk
import os

# set up connection to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# for sub folder, add <.folder("your folder name")>
inbox = outlook.GetDefaultFolder(6).folders("MyFolder")

# Access to the email in the inbox
messages = inbox.Items

#Todays Date
today_date = str(datetime.date.today())

for iteration, message in enumerate(messages):
    sent_date = str(message.senton.date())
    if sent_date == today_date:
        attachments = message.Attachments
        attachment = attachments.Item(1)
        attachment_name = str(attachment)
        attachment.Saveasfile("V:\Data & Analytics\Rate360\Datav2" + "\\" + attachment_name)
        print('saved')

dir = "V:\Data & Analytics\Rate360\Datav2"
count = 0
#iterate directory
for path in os.listdir(dir):
    if os.path.isfile(os.path.join(dir,path)):
        count+=1
count = str(count)

root= tk.Tk() 
 
canvas1 = tk.Canvas(root, width = 300, height = 300)
canvas1.pack()

label1 = tk.Label(root, text=(dir + ' ' + 'file count:', count))
canvas1.create_window(150, 150, window=label1)

root.mainloop()