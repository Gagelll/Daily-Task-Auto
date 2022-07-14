import win32com.client

outlook = win32com.client.Dispatch('outlook.application')

mail = outlook.CreateItem(0)

mail.To = 'lorimergage@gmail.com'
mail.Subject = 'Python email test'
mail.HTMLBody = '<h3>This is a test</h3>'
mail.Body = "This is a test"

mail.Send()