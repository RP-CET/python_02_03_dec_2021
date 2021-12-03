## REMEMBER to perform a pip install pywin32 first.

import win32com.client

outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = '<<Replace receiver email address>>'
mail.Subject = 'Sample Email'
mail.HTMLBody = '<h3>This is an email sent from Outlook...</h3>'
mail.Body = "This is the normal body"
#mail.Attachments.Add('c:\\sample.xlsx')
#mail.Attachments.Add('c:\\sample2.xlsx')
#mail.CC = 'somebody@company.com'
mail.Send()
print("Sent")