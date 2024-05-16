""" Ancel Carson
    Napps Technology Comporation
    10/14/2020
    Email Test.py: Test to try sending emails via Python
"""
#Libraries
import win32com.client as win32

#Variables

#Functions
" Main Finction "
def sendEmail():
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'acarson@nappstech.com'
    mail.Subject = 'Test Email'
    mail.Body = 'This is a test email'
    mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    #attachment  = "Path to the attachment"
    #mail.Attachments.Add(attachment)

    mail.Send()

" Checks if this program is beiong called "
if __name__ == "__main__":
    sendEmail()

