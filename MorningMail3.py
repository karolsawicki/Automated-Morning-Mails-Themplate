import win32com.client as win32
from win32com.client import Dispatch
import os


# This portion deals with the mail part
outlook = win32.Dispatch('outlook.application')
# Workaround to send the mail from a different mail id-----------  skip this part if you want to send mail from your default mail box
sendfromAC = None
for oacc in outlook.Session.Accounts:
            if oacc.SmtpAddress == "karol.sawicki@example.com":  # Mail id from which to send the mail
                sendfromAC = oacc
                break
    # ----------------------------------------------------------------

mail = outlook.CreateItem(0)

    # ----------------------------------------------------------------
if sendfromAC:
      
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, sendfromAC))
    # ----------------------------------------------------------------

    mail.To = 'CONFIDENTIAL@example.com;'
    mail.Cc = 'CONFIDENTIAL@example.com;'
    mail.Subject = '[CONFIDENTIAL] CONFIDENTIAL'

    mail.HTMLBody = mail.HTMLBody + "<BR>Cześć,<b> </b>" \
        + "<BR><BR>Noc spokojna, kolejka bez zmian. </b>"\
        + "<BR><BR> Pozdrawiam / Regards"\
        + "<BR> <b> Karol Sawicki </b>"\
        + "<BR CONFIDENTIAL "\
        + "<BR> CONFIDENTIAL"\
        + "<BR> CONFIDENTIAL"\
        + "<BR> CONFIDENTIAL"\
        + "<BR> e-mail:  CONFIDENTIAL"
         
    mail.Display(False)


