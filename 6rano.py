import win32com.client as win32
from win32com.client import Dispatch
import os


body = """
CONFIDENTIAL
"""

# This portion deals with the mail part
outlook = win32.Dispatch('outlook.application')
# Workaround to send the mail from a different mail id-----------  skip this part if you want to send mail from your default mail box
sendfromAC = None
for oacc in outlook.Session.Accounts:
            if oacc.SmtpAddress == "karol.sawicki@CONFIDENTIAL":  # Mail id from which to send the mail
                sendfromAC = oacc
                break
    # ----------------------------------------------------------------

mail = outlook.CreateItem(0)

    # ----------------------------------------------------------------
if sendfromAC:
        #Msg.SendUsingAccount = oacctouse
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
    # ----------------------------------------------------------------
   
    
    
    mail.To = 'CONFIDENTIAL@example.com;'
    mail.Cc = 'CONFIDENTIAL@example.com;'
    mail.Subject = '[CONFIDENTIAL] CONFIDENTIAL'

    mail.HTMLBody = body
        
    mail.Display(False)
   
    