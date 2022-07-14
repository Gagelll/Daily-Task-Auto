import win32com.client
import datetime
import tkinter as tk 

# set up connection to outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# for sub folder, add <.folder("your folder name")>
inbox = outlook.GetDefaultFolder(6)

messages = inbox.Items

#Todays Date
today_date = str(datetime.date.today())

for email in messages:
    if "test" in email.subject.lower() and str(email.senton.date()) == today_date:
        print(email.subject)
        attachments = email.Attachments
        attachment = attachments.Item(1)
        attachment_name = str(attachment)
        attachment.SaveASfile("V:\\Data & Analytics\\Dirty_Files"+ '\\' + attachment_name)
        print(attachment_name + " " + "has been saved")

root= tk.Tk() 
 
canvas1 = tk.Canvas(root, width = 300, height = 300)
canvas1.pack()

label1 = tk.Label(root, text=(attachment_name + " " + "has been saved"))
canvas1.create_window(150, 150, window=label1)

root.mainloop()