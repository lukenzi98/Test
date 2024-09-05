import win32com.client as win32

def sendingEmails():
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)

    recipients = ["luka.stojanovic@fisglobal.com"]
    
    for recipient in recipients:
        mail.Recipients.Add(recipient)

    mail.Subject = 'Python script test'
    mail.Body = """
    Test
                """
    # attachment_path = r'path/to/attachment.txt'
    # mail.Attachments.Add(attachment_path)
    mail.Send()

sendingEmails()